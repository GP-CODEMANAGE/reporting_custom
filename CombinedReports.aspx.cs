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
using System.IO;
using System.Reflection;
using System.Net;
using System.Text;
using Spire.Xls;
using System.Drawing;
using System.Data.Common;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using org.jfree.chart;
using org.jfree.chart.block;
using org.jfree.chart.plot;
using org.jfree.chart.title;

//using org.jfree.ui;

public partial class CombinedReports : System.Web.UI.Page
{
    Boolean fbCheckExcel = false;
    public StreamWriter sw = null;
    string strDescription = string.Empty;
    //bool bProceed = true;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName;
    public enum ReportType
    {
        PortfolioConsChart = 1,
        CommitmentSchedule = 2,
        AllocationGroupPieChart = 3,
        InvestmentObjectiveChart = 4,
        OverallPieChart = 5
    }
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            mvShowReport.ActiveViewIndex = 0;
            FillReportflag();
            fillHousehold();
        }
        
    }
    public string getFinalSp(ReportType Type)
    {
        String lsSQL = "";
        string houseHold = "";
        if (ddlHousehold.SelectedValue != "0")
        {
            houseHold = "'" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "'";
        }
        string AsOfDate = txtAsofdate.Text.Trim() == "" ? "null" : "'" + txtAsofdate.Text.Trim() + "'";
        object CashFlag = ddlCash.SelectedValue == "0" ? "null" : ddlCash.SelectedValue;
        object ReportFlag = ddlReportFlg.SelectedValue == "0" ? "null" : ddlReportFlg.SelectedValue;
        object AllocationGroup = ddlAllocationGroup.SelectedValue == "0" ? "null" : "'" + ddlAllocationGroup.SelectedItem.Text.Replace("'", "") + "'";
        object Report1and2 = ddlReport1and2.SelectedValue == "" ? "null" : ddlReport1and2.SelectedValue;
        object AllAsset = ddlAllAsset.SelectedValue == "0" ? "null" : ddlAllAsset.SelectedValue;
        //if (ddlHousehold.SelectedValue != "")
        //{
            if (Type == ReportType.PortfolioConsChart)
            {
                lsSQL = "exec SP_R_CONSTRUCTIONCHART '" + ddlHousehold.SelectedItem.Text + "'," + AsOfDate + ", null, " + AllocationGroup + "";//"exec SP_R_COMMITMENTSCHEDULE 'Kuppenheimer Family', '20110430', null, 0, 0, null, 0, 1";
            }
            else if (Type == ReportType.CommitmentSchedule)
            {
                lsSQL = "exec SP_R_COMMITMENTSCHEDULE '" + ddlHousehold.SelectedItem.Text + "'," + AsOfDate + ", null, 0,0, " + AllocationGroup + ",0,1";
            }
            else if (Type == ReportType.InvestmentObjectiveChart)
            {
                lsSQL = "exec GreshamPartners_MSCRM.dbo.sp_R_Investment_Objective_Chart_excel_SMA_New  @HouseholdName  = " + houseHold + ", @AsofDate = " + AsOfDate + ", @GreshamAdvisedFlagTxt = 'TIA',@AllocGroupName = " + AllocationGroup + "";
            }
            else if (Type == ReportType.AllocationGroupPieChart)
            {
                lsSQL = "exec SP_R_CONSTRUCTIONCHART '" + ddlHousehold.SelectedItem.Text + "'," + AsOfDate + ", null, null";
                //lsSQL = "exec GreshamPartners_MSCRM.dbo.SP_R_PositionswithRollUpGroup_excel_SMA   @HouseholdName  = '" + houseHold + ", @AsofDate = " + AsOfDate + ", @ReportGroupFlg = 1";
            }
            else if (Type == ReportType.OverallPieChart)
            {
                lsSQL = "exec SP_R_CONSTRUCTIONCHART '" + ddlHousehold.SelectedItem.Text + "'," + AsOfDate + ", null, " + AllocationGroup + "";
                //lsSQL = "exec GreshamPartners_MSCRM.dbo.SP_R_PositionswithRollUpGroup_excel_SMA  @HouseholdName='" + ddlHousehold.SelectedItem.Text + "',@AsofDate=" + AsOfDate + ",@GreshamAdvisedFlagTxt= 'TIA'";  
            }
        //}
        //else
        //{
        //    if (Type == ReportType.PortfolioConsChart)
        //    {
        //        lsSQL = "exec SP_R_CONSTRUCTIONCHART '" + ddlHousehold.SelectedItem.Text + "'," + AsOfDate + ", null, " + AllocationGroup + "";//"SP_R_Advent_Report_Other";
        //    }
        //    else if(Type == ReportType.CommitmentSchedule)
        //    {
        //        lsSQL = "exec SP_R_COMMITMENTSCHEDULE '" + ddlHousehold.SelectedItem.Text + "'," + AsOfDate + ", null, " + CashFlag + "," + ReportFlag + ", " + AllocationGroup + "," + Report1and2 + "," + AllAsset + "";//"SP_R_Advent_Report_Other";
        //    }
        //    else if (Type == ReportType.InvestmentObjectiveChart)
        //    {
        //         lsSQL = "exec GreshamPartners_MSCRM.dbo.sp_R_Investment_Objective_Chart_excel_SMA_New  @HouseholdName  = " + houseHold + ", @AsofDate = " + AsOfDate + ", @GreshamAdvisedFlagTxt = 'TIA',@AllocGroupName = " + AllocationGroup + "";
        //    }
        //    else if (Type == ReportType.AllocationGroupPieChart)
        //    {
        //        lsSQL = "exec GreshamPartners_MSCRM.dbo.SP_R_PositionswithRollUpGroup_excel_SMA   @HouseholdName  = '" + houseHold + "', @AsofDate = " + AsOfDate + ", @ReportGroupFlg = 1";

        //    }
        //    else if (Type == ReportType.OverallPieChart)
        //    {
        //        lsSQL = "exec GreshamPartners_MSCRM.dbo.SP_R_PositionswithRollUpGroup_excel_SMA  @HouseholdName='" + ddlHousehold.SelectedItem.Text + "',@AsofDate=" + AsOfDate + ",@GreshamAdvisedFlagTxt= 'TIA'";  
        //    }
        //}
        return lsSQL;
    }


    public void FillReportflag()
    {
        //ddlReportGroupflag.Items.Clear();
        //ddlReportgroupflag2.Items.Clear();
        //ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("All", "null"));
        //ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("All", "null"));

        //ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("Yes", "1"));
        //ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("Yes", "1"));

        //ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("No", "0"));
        //ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("No", "0"));
        //ddlReportGroupflag.SelectedValue = "1";
        //ddlReportgroupflag2.SelectedValue = "null";

    }
    public void fillHousehold()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("sp_s_Get_HouseHoldName");
        ddlHousehold.Items.Clear();
        ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }
    public void fillContact()
    {
        DB clsDB = new DB();
        ddlReportFlg.Items.Clear();

        DataSet loDataset = clsDB.getDataSet("sp_r_Household_contact_list @Householdname ='" + ddlHousehold.SelectedItem + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlReportFlg.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }
    public void AllocationGroup()
    {
        DB clsDB = new DB();
        ddlAllocationGroup.Items.Clear();
        drpAllocationGroupTitle.Items.Clear();
        ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        DataSet loDataset = clsDB.getDataSet("SP_S_Advent_Allocation_Group  @Householdname ='" + ddlHousehold.SelectedItem + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"])));
        }

    }

    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        // fillContact();
        AllocationGroup();
        fillHouseholdTitle();
        lblError.Text = "";
    }
    protected void drpAllocationGroupTitle_SelectedIndexChanged(object sender, EventArgs e)
    {
        fillGroupAllocationTitle();
    }
    public void fillGroupAllocationTitle()
    {
        DB clsDB = new DB();
        drpAllocationGroupTitle.Items.Clear();
        //drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        if (!String.IsNullOrEmpty(ddlAllocationGroup.SelectedValue))
        {
            //drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", ""));
            //  Response.Write("SP_S_AllocationGroupTitle  @AllocationGroupName ='" + ddlAllocationGroup.SelectedValue.Replace("'", "''") + "'");
            DataSet loDataset = clsDB.getDataSet("SP_S_AllocationGroupTitle  @AllocationGroupName ='" + ddlAllocationGroup.SelectedValue.Replace("'", "''") + "'");
            for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
            {
                if (Convert.ToString(loDataset.Tables[0].Rows[liCounter]["Column1"]) == "0")
                {
                    drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
                }
                else
                {
                    drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
                }
            }
        }

    }
    public void fillHouseholdTitle()
    {
        DB clsDB = new DB();
        drpHouseHoldReportTitle.Items.Clear();
        if (!String.IsNullOrEmpty(ddlHousehold.SelectedValue))
        {
            //drpHouseHoldReportTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", ""));
            //    Response.Write(" SP_S_HouseHoldTitle @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'","''") + "'");
            DataSet loDataset = clsDB.getDataSet(" SP_S_HouseHoldTitle @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "'");
            for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
            {
                drpHouseHoldReportTitle.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
            }
        }

    }

    protected void Button1_Click(object sender, EventArgs e)
    {

        DateTime loDatetme1 = new DateTime();
        loDatetme1 = Convert.ToDateTime(txtAsofdate.Text);
        DateTime loFindEndofday1 = new DateTime(loDatetme1.Year, loDatetme1.Month, 1).AddMonths(1).AddDays(-1);
        lblError.Text = "";

        if (RadioButton1.Checked)
        {
            fbCheckExcel = false;
            mvShowReport.ActiveViewIndex = 1;
        }
        if (RadioButton2.Checked)
        {
            try
            {
                fbCheckExcel = true;

            }
            catch (Exception ex)
            {
                Response.Write(ex.ToString());
                Response.Write(ex.StackTrace);
            }
        }


        clsCombinedReports obj = new clsCombinedReports();

        if (ddlHousehold.SelectedValue != "0")
        {
            if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
            {
                obj.lsFamiliesName = ddlAllocationGroup.SelectedItem.Text;
            }
            else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
            {
                obj.lsFamiliesName = drpHouseHoldReportTitle.SelectedItem.Text;
            }
            else
            {
                obj.lsFamiliesName = drpAllocationGroupTitle.SelectedItem.Text;
            }
        }

        if (ddlHousehold.SelectedValue != "0")
        {
            obj.HouseHoldText = "" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "";
        }
        obj.AsOfDate = txtAsofdate.Text.Trim() == "" ? "null" : "" + txtAsofdate.Text.Trim() + "";
        obj.AllocationGroupText = ddlAllocationGroup.SelectedValue == "0" ? "" : "" + ddlAllocationGroup.SelectedItem.Text.Replace("'", "''") + "";
 

        if (rdbtnPDF.Checked)
        {
            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
            string DestinationFileName = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\Gresham_" + strGUID + ".pdf";
           
            int count = 0;
            for (int i = 0; i < lstReport.Items.Count; i++)
            {
                if (lstReport.Items[i].Selected == true)
                {
                    count = count + 1;
                }
            }
            if (count > 0)
            {
                string[] SourceFileName = new string[0];
                string FileName=string.Empty;
                if (lstReport.Items[0].Selected == true)
                {
                    SourceFileName = new string[5];
                }
                else
                    SourceFileName = new string[count];

                for (int i = 0; i < count; i++)
                {
                    if (lstReport.Items[0].Selected == true && i == 0)
                    {
                        SourceFileName[0] = obj.generatePortfolioConstChart().Replace("No Record Found","");
                        SourceFileName[1] = obj.generateCommittmentSchReport().Replace("No Record Found", "");
                        SourceFileName[2] = obj.generateInvestmentObjectiveChart().Replace("No Record Found", "");
                        SourceFileName[3] = obj.generateAllocationGroupPieChart().Replace("No Record Found", "");
                        SourceFileName[4] = obj.generateOverAllPieChart().Replace("No Record Found", "");
                        i = count + 1;
                    }
                    else
                    {
                        if (lstReport.Items[1].Selected == true)               //overall piechart
                        {
                            //SourceFileName[i] = generateCommittmentSchReport();
                            //i++;
                            SourceFileName[i] = obj.generatePortfolioConstChart().Replace("No Record Found", "");
                            string File = SourceFileName[i];
                            if (File == "No Record Found" || Convert.ToString(File) == "")
                            {
                                lblError.Text = "No Record Found";
                                return;
                            }
                            i++;
                        }
                        if (lstReport.Items[2].Selected == true)                //Allocation Group Pie Chart
                        {
                            //SourceFileName[i] = generateAllocationGroupPieChart();
                            //i++;
                            SourceFileName[i] = obj.generateCommittmentSchReport().Replace("No Record Found", "");
                            string File = SourceFileName[i];
                            if (File == "No Record Found" || Convert.ToString(File) == "")
                            {
                                lblError.Text = "No Record Found";
                                return;
                            }

                            i++;
                        }
                        if (lstReport.Items[3].Selected == true)                //Portfolio Construction Chart
                        {
                            //SourceFileName[i] = generatePortfolioConstChart();
                            //i++;
                            SourceFileName[i] = obj.generateInvestmentObjectiveChart().Replace("No Record Found", "");
                            string File = SourceFileName[i];
                            if (File == "No Record Found" || Convert.ToString(File) == "")
                            {
                                lblError.Text = "No Record Found";
                                return;
                            }
                            i++;
                        }
                        if (lstReport.Items[4].Selected == true)                //Investment Objective Chart
                        {
                            //SourceFileName[i] = generateInvestmentObjectiveChart();
                            //i++;
                            SourceFileName[i] = obj.generateAllocationGroupPieChart().Replace("No Record Found", "");
                            string File = SourceFileName[i];
                            if (File == "No Record Found" || Convert.ToString(File) == "")
                            {
                                lblError.Text = "No Record Found";
                                return;
                            }
                            i++;
                        }
                        if (lstReport.Items[5].Selected == true)                 //Commitment Schedule
                        {
                            SourceFileName[i] = obj.generateOverAllPieChart().Replace("No Record Found", "");
                            string File = SourceFileName[i];
                            if (File == "No Record Found" || Convert.ToString(File) == "")
                            {
                                lblError.Text = "No Record Found";
                                return;
                            }
                            i++;
                        }
                    }
                }

                PDFMerge PDF = new PDFMerge();
                PDF.MergeFiles(DestinationFileName, SourceFileName);

                try
                {

                    Response.Write("<script>");
                    string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/Gresham_" + strGUID + ".pdf";
                    Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    Response.Write("</script>");

                }
                catch (Exception exc)
                {
                    Response.Write(exc.Message);
                }
            }
        }
    }
    #region otherCode
    void Page_PreInit(object sender, System.EventArgs args)
    {
        gvReport.SkinID = "gvReportSkin";
    }

    protected void btnBack_Click(object sender, EventArgs e)
    {
        mvShowReport.ActiveViewIndex = 0;
    }

    protected void BtnExport_Click(object sender, EventArgs e)
    {
        fbCheckExcel = true;
    }

    public void ExportExcel()
    {
        String FileName = "report";

        Response.Clear();

        Response.Write("<html xmlns:o=\"urn:schemasmicrosoft-com:office:office\" xmlns:x=\"urn:schemas-microsoftcom:office:excel\">");

        Response.Write("<head>");
        Response.Write("<!--[if gte mso 9]><xml>");
        Response.Write("<x:ExcelWorkbook>");
        Response.Write("<x:ExcelWorksheets>");
        Response.Write(" <x:ExcelWorksheet>");
        Response.Write(" <x:Name>report</x:Name>");
        Response.Write(" <x:WorksheetOptions>");
        Response.Write("<x:PageSetup><Layout x:Orientation=\"Landscape\"/><x:/PageSetup>");
        Response.Write(" <x:DisplayPageBreak/>");
        Response.Write(" <x:Print>");
        //  Response.Write(" <x:BlackAndWhite/>");
        //Response.Write(" <x:DraftQuality/>");
        Response.Write(" <x:ValidPrinterInfo/>");
        Response.Write(" <x:PaperSizeIndex>5</x:PaperSizeIndex>");
        Response.Write(" <x:Scale>85</x:Scale>");
        Response.Write(" <x:HorizontalResolution>600</x:HorizontalResolution>");
        // Response.Write(" <x:Gridlines/>");
        //Response.Write(" <x:RowColHeadings/>");
        // Response.Write("(<x:RepeatedRows>$1:$6<x:RepeatedCols>");
        // Response.Write("<x:Formula>=report!$6:$6</x:Formula>");
        Response.Write(" </x:Print>");
        Response.Write(" </x:WorksheetOptions>");
        Response.Write(" </x:ExcelWorksheet>");
        Response.Write(" </x:ExcelWorksheets>");
        Response.Write("</x:ExcelWorkbook>");
        Response.Write("<x:ExcelName>");
        Response.Write("<x:Name>Print_Titles</x:Name>");
        Response.Write("<x:SheetIndex>0</x:SheetIndex>");

        Response.Write("<x:Formula>=report!$A:$G,report!$1:$9</x:Formula>");

        Response.Write("</x:ExcelName>");
        Response.Write("</xml><![endif]-->");

        ///       Response.Write("<style>body {font-family:Frutiger 55 Roman;font-size:8pt} .PercentageDecimal{	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0%_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0%\\\\\\) ;}  .ddcblk { border-bottom:1pt solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}    .whiteclass {	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .greyclass {	background-color:#D8D8D8;} .BackgroundColor{	background-color:#B7DDE8;}.dummyheader{padding-left:5px; }.dummy{ border-top:1pt solid #000000;}.Title {	font-family:Frutiger 55 Roman; font-size:18px;	font-weight:normal;	text-decoration:none;}.gvReportss {border-bottom:1pt solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssNo {border-bottom:1pt solid #ffffff;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssBlack {border-bottom:thin solid #000000;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .ddcblkss {border-bottom:thin solid #000000;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .ddcblksswhite {border-bottom:thin solid #ffffff;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}");

        //      Response.Write("<style>body {font-family:Frutiger 55 Roman;font-size:8pt} .PercentageDecimal{	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0%_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0%\\\\\\) ;}  .ddcblk { border-bottom:thin solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}    .whiteclass {	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .greyclass {	background-color:#D8D8D8;} .BackgroundColor{	background-color:#B7DDE8;}.dummyheader{padding-left:5px; }.dummy{ border-top:thin solid #000000;}.Title {	font-family:Frutiger 55 Roman; font-size:18px;	font-weight:normal;	text-decoration:none;}.gvReportss {border-bottom:thin solid #F2F2F2;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssNo {border-bottom:thin solid #ffffff;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssBlack {border-bottom:thin solid #000000;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .ddcblkss {border-bottom:thin solid #000000;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .ddcblksswhite {border-bottom:thin solid #ffffff;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}");
        Response.Write("<style> @page  {   margin:.5in .25in .5in .25in; mso-horizontal-page-align:center; mso-header-margin:.25in; mso-footer-margin:.25in; mso-footer-color:red;;mso-footer-data : '&C&\\0022Frutiger 55 Roman\\,Regular\\0022&8 Page &P of &N &R&\\0022Frutiger 55 Roman\\,italic\\0022&8  &KD8D8D8&D, &T'   } body {font-family:Frutiger 55 Roman;font-size:8pt} .PercentageDecimal{	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0%_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0%\\\\\\) ;}  .ddcblk { border-bottom:.5pt hairline #F2F2F2;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}    .whiteclass {	background-color:#ffffff;mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .greyclass {	background-color:#D8D8D8;} .BackgroundColor{	background-color:#B7DDE8;}.dummyheader{padding-left:5px;height:16px }.dummy{ border-top:thin solid #000000;}.Title {	font-family:Frutiger 55 Roman; font-size:18px;	font-weight:normal;	text-decoration:none;}.gvReportss {border-bottom:.5pt hairline #F2F2F2;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssNo {border-bottom:thin solid #ffffff;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .gvReportssBlack {border-bottom:thin solid #000000;mso-number-format:\\#\\,\\#\\#0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\\\\\) ;} .ddcblkss {border-bottom:thin solid #000000;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;} .ddcblksswhite {border-bottom:thin solid #ffffff;  mso-number-format:\\#\\,\\#\\#0\\.0_\\)\\;\\[Black\\]\\\\\\(\\#\\,\\#\\#0\\.0\\\\\\) ;}");

        Response.Write(".familyname { font-family:Frutiger 55 Roman;font-size:14pt;font-weight:bold;height:18.0pt; } ");
        Response.Write("ht25px { height:25px; } .assetdistribution { font-family:Frutiger 55 Roman;font-size:12pt; } ");
        Response.Write(".assDate { font-family:Frutiger 55 Roman;font-size:10pt;font-style:italic; } ");


        Response.Write("</style> ");
        Response.Write("</head>");
        Response.Write("<body>");
        Response.Charset = "";
        Response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xls", FileName));
        // Response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xlsx", FileName));

        //Response.Cache.SetCacheability(HttpCacheability.NoCache);
        Response.ContentType = "application/vnd.ms-excel";

        System.IO.StringWriter stringWrite = new System.IO.StringWriter();
        System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
        gvReport.RenderControl(htmlWrite);
        Response.Write(stringWrite.ToString());
        Response.Write("</body>");
        Response.Write("</html>");
        Response.End();

    }

    public void ExportGVtoExcel(GridView gvexcel, string filename)
    {
        HttpResponse response = HttpContext.Current.Response;
        gvexcel.AllowPaging = false;
        gvexcel.AllowSorting = false;

        response.Clear();
        response.Charset = "";
        response.ContentType = "application/vnd.ms-excel";
        response.AddHeader("content-disposition", string.Format("attachment;filename={0}.xlsx", filename));

        using (StringWriter sw = new StringWriter())
        {
            using (HtmlTextWriter htw = new HtmlTextWriter(sw))
            {
                gvexcel.RenderControl(htw);

                response.Write(sw.ToString());
                response.End();
            }
        }
    }
    public override void VerifyRenderingInServerForm(Control control)
    {
        //this method requires for exportGridtoexcel function
    }

    public void grd_clientview_onitemcommand(object sender, GridViewRowEventArgs e)
    {

        if (e.Row.RowType == DataControlRowType.Header && !String.IsNullOrEmpty(txtAsofdate.Text))
        {


            GridView oGridView = (GridView)sender;
            GridViewRow oGridViewRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
            oGridViewRow.BorderWidth = Unit.Pixel(0);

            String lsfamilyName = ddlHousehold.SelectedItem.Text;

            int liCommaCounter = lsfamilyName.IndexOf(",");
            int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
            if (liCommaCounter > 0 && liSpaceCounter > 0)
                lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
            else
                lsfamilyName = lsfamilyName;
            if (ddlAllocationGroup.SelectedValue != "")
            {
                lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
            }
            //if (!String.IsNullOrEmpty(drpHouseHoldReportTitle.SelectedValue))
            //    lsfamilyName = drpHouseHoldReportTitle.SelectedValue;

            //if (!String.IsNullOrEmpty(drpAllocationGroupTitle.SelectedValue))
            //    lsfamilyName = drpAllocationGroupTitle.SelectedValue;


            System.Web.UI.WebControls.Table loTable = new System.Web.UI.WebControls.Table();
            loTable.Width = Unit.Percentage(100);
            loTable.HorizontalAlign = HorizontalAlign.Center;
            TableCell oTableCell = new TableCell();

            TableRow loRow = new TableRow();
            TableCell loCell = new TableCell();
            loCell.Text = "";
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);

            loRow = new TableRow();
            loRow.Height = Unit.Pixel(25);
            loCell = new TableCell();
            loCell.Text = lsfamilyName;
            loCell.CssClass = "familyname";
            loCell.Height = Unit.Pixel(25);
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);

            loRow = new TableRow();
            loCell = new TableCell();
            loCell.Text = "ASSET DISTRIBUTION";
            if (ddlAllAsset.SelectedItem.ToString() != "Horizontal")
                loCell.Text = "ASSET DISTRIBUTION COMPARISON";
            loCell.CssClass = "assetdistribution";
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);



            loRow = new TableRow();
            loCell = new TableCell();
            loCell.Text = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "<span style=\"color: #ffffff;\">.</span>" + "</span>";
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loCell.CssClass = "assDate";
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);


            loRow = new TableRow();
            loCell = new TableCell();
            loCell.Text = "";
            loCell.ColumnSpan = gvReport.Columns.Count - 1;
            loCell.HorizontalAlign = HorizontalAlign.Center;
            loRow.Cells.Add(loCell);
            loTable.Rows.Add(loRow);




            //loRow = new TableRow();
            //loCell = new TableCell();
            //loCell.Text = "<br><span style=\"font-family:Frutiger 55 Roman;	font-size:12pt;\">" + "<span  style=\"font-family:Frutiger 55 Roman;	font-size:14pt;font-weight:bold;\" >" + lsfamilyName + "</span>";
            //loCell.Text = oTableCell.Text + "<br>ASSET DISTRIBUTION";
            //loCell.Text = oTableCell.Text + "</span><br><span style=\"font-family:Frutiger 55 Roman;	font-size:10pt;font-style:italic;\">" + " " + Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "<span style=\"color: #ffffff;\">.</span><br> <span style=\"color: #ffffff;\">.</span> " + "</span>";
            //loCell.HorizontalAlign = HorizontalAlign.Center;
            //loRow.Cells.Add(loCell);
            //loTable.Rows.Add(loRow);







            oTableCell.Controls.Add(loTable);
            oTableCell.Height = Unit.Pixel(25);
            oTableCell.ColumnSpan = gvReport.Columns.Count - 1;
            oTableCell.HorizontalAlign = HorizontalAlign.Center;
            oGridViewRow.CssClass = "ht25px";
            oGridViewRow.Cells.Add(oTableCell);
            oGridView.Controls[0].Controls.AddAt(0, oGridViewRow);


        }
    }

    private DataSet AddTotals(DataSet lodataset)
    {
        DataTable table = new DataTable("GrandTot");
        table = lodataset.Tables[0].Clone();
        DataRow dr = table.NewRow();//lodataset.Tables[0].NewRow();

        for (int i = 0; i < lodataset.Tables.Count; i++)
        {
            if (lodataset.Tables[i].Rows.Count > 0)
            {
                for (int j = 0; j < lodataset.Tables[i].Rows.Count; j++)
                {

                    if (Convert.ToString(lodataset.Tables[i].Rows[j]["_OrderNmb"]).ToLower() == "3")
                    {
                        for (int k = 0; k < lodataset.Tables[0].Columns.Count; k++)
                        {
                            if (lodataset.Tables[i].Columns[k].ColumnName.Equals("Commitment"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }
                            else if (lodataset.Tables[i].Columns[k].ColumnName.Equals("CalledToDate"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }
                            else if (lodataset.Tables[i].Columns[k].ColumnName.Equals("ReinvestedDistributionToDate"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }
                            else if (lodataset.Tables[i].Columns[k].ColumnName.Equals("TotalInvestedToDate"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }
                            else if (lodataset.Tables[i].Columns[k].ColumnName.Equals("RemainingCommitment"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }
                            else if (lodataset.Tables[i].Columns[k].ColumnName.Equals("ExpectedRemaining"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }
                            else if (lodataset.Tables[i].Columns[k].ColumnName.Equals("ExpectedRemainingCallsCurrentQuarter"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }
                            else if (lodataset.Tables[i].Columns[k].ColumnName.Equals("ExpectedRemainingCallsNextQuarter"))
                            {
                                if (Convert.ToString(dr[lodataset.Tables[i].Columns[k].ColumnName]) == "")
                                    dr[lodataset.Tables[i].Columns[k].ColumnName] = 0.0M;

                                dr[lodataset.Tables[i].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[i].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[i].Rows[j][k]);
                            }

                        }
                    }
                }
            }


        }



        dr["_OrderNmb"] = 7;
        dr["Investment"] = "Total Commtments";
        //dr["Firm"] = "Grand Total";
        table.Rows.Add(dr);
        //table.Rows.Add(dr);
        table.TableName = "GrandTot";
        lodataset.Tables.Add(table);


        lodataset.AcceptChanges();

        return lodataset;
    }

  public void SetBorder(Cell foCell, bool IsTop, bool IsBottom, bool IsLeft, bool IsRight)
    {
        if (IsTop == true)
        {
            foCell.BorderWidthTop = 1F;
            foCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsBottom == true)
        {
            foCell.BorderWidthBottom = 1F;
            foCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsLeft == true)
        {
            foCell.BorderWidthLeft = 1F;
            foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsRight == true)
        {
            foCell.BorderWidthRight = 1F;
            foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        //if (TopBottom == true)
        //{
        //    foCell.BorderWidthBottom = 1F;
        //    foCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthLeft = 0.2F;
        //    foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthRight = 0.2F;
        //    foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        //}
        //else
        //{
        //    foCell.BorderWidthTop= 1F;
        //    foCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthLeft = 0.2F;
        //    foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthRight = 0.2F;
        //    foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        //}

    }
    #endregion

    public string generateInvestmentObjectiveChart()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = String.Empty;
        //String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);
        String lsSQL = getFinalSp(ReportType.InvestmentObjectiveChart);
        // Response.Write(lsSQL);
        newdataset = clsDB.getDataSet(lsSQL);

        if (newdataset.Tables[0].Rows.Count > 0)
        {
            DataRow dr = newdataset.Tables[0].NewRow();

            for (int j = 0; j < newdataset.Tables[0].Rows.Count; j++)
            {
                for (int k = 0; k < newdataset.Tables[0].Columns.Count; k++)
                {
                    if (newdataset.Tables[0].Columns[k].ColumnName.Contains("Current Portfolio %"))
                    {
                        if (Convert.ToString(dr[newdataset.Tables[0].Columns[k].ColumnName]) == "")
                            dr[newdataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                        dr[newdataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[newdataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(newdataset.Tables[0].Rows[j][k]);
                    }

                    if (newdataset.Tables[0].Columns[k].ColumnName.Contains("Suggested Allocation"))
                    {
                        if (Convert.ToString(dr[newdataset.Tables[0].Columns[k].ColumnName]) == "")
                            dr[newdataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                        dr[newdataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[newdataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(newdataset.Tables[0].Rows[j][k]);
                    }
                }
            }
            dr["_LineFlg"] = 2;
            //dr["_FundName"] = "Total";
            newdataset.Tables[0].Rows.Add(dr);
            newdataset.AcceptChanges();
        }
        if (newdataset.Tables[0].Rows.Count < 1)
        {
            lblError.Text = "No Record Found";
            return "No Record Found";
        }

        DataSet loInsertblankRow = newdataset.Copy();

        newdataset = loInsertblankRow.Clone();

        // string strGUID = Guid.NewGuid().ToString();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        // strGUID = strGUID.Substring(0, 5);
        // String fsFinalLocation = @"C:\Reports\" + strGUID + ".xls";

        String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\IOC_" + strGUID + ".xls";
        int liBlankCounter = 0;

        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Columns.Count; liBlankRow++)
        {
            if (liBlankRow != 0 && liBlankRow != 2 && liBlankRow != 3 && liBlankRow != 4 && liBlankRow != 12)
            {
                loInsertblankRow.Tables[0].Columns.RemoveAt(liBlankRow);

                for (int i = 0; i < loInsertblankRow.Tables[0].Columns.Count; i++)
                {
                    if (i > 3)
                    {
                        loInsertblankRow.Tables[0].Columns.RemoveAt(i);

                        for (int j = 0; j < loInsertblankRow.Tables[0].Columns.Count; j++)
                        {
                            if (j > 3)
                            {
                                loInsertblankRow.Tables[0].Columns.RemoveAt(j);
                            }
                        }
                    }
                }
            }


        }

        loInsertblankRow.Tables[0].AcceptChanges();
        DataSet lodataset = new DataSet();
        lodataset = loInsertblankRow.Copy();
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
        loTable.Cellpadding = 0f;
        loTable.Cellspacing = 0f;


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

                loInsertdataset.AcceptChanges();
                setHeaderInvestmentObjective(document, loInsertdataset, newdataset);
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
                    if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1)
                    {
                        lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(lsFormatedString));
                    }
                    else
                    {
                        lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(lsFormatedString));
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
                loCell.Leading = 4f;//6

                loCell.UseBorderPadding = true;

                //  if (lodataset.Tables[0].Rows[liRowCount]["_Ssi_TabFlg"].ToString() == "True" && lodataset.Tables[0].Rows[liRowCount]["_Ssi_UnderlineFlg"].ToString() != "True")


                if (liColumnCount != 0 && liColumnCount != 3)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                }
                else if (liColumnCount == 3)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                }

                if (liColumnCount == 1)
                {
                    if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "1")
                    {

                        loCell.EnableBorderSide(2);
                    }
                    else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        string CurrentAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                        //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                        lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                        //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                    }
                }
                else if (liColumnCount == 2)
                {
                    if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "1")
                    {

                        loCell.EnableBorderSide(2);
                    }
                    else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        string CurrentAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                        //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                        lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                        //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                    }
                }

                loCell.Add(lochunk);
                loTable.AddCell(loCell);
            }

            try
            {
                if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
                {
                    document.Add(loTable);
                    liCurrentPage = liCurrentPage + 1;
                    //document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                }
            }
            catch (Exception Ex)
            {

            }
        }

        document.Close();

        FileInfo loFile = new FileInfo(ls);
        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        return fsFinalLocation.Replace(".xls", ".pdf");

        //if (loInsertdataset.Tables[0].Rows.Count > 0)
        //{
        //    document.Close();

        //    FileInfo loFile = new FileInfo(ls);
        //    try
        //    {
        //        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

        //        Response.Write("<script>");
        //        string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
        //        Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        //        Response.Write("</script>");

        //    }
        //    catch (Exception exc)
        //    {
        //        Response.Write(exc.Message);
        //    }
        //}
    }

    public string generateAllocationGroupPieChart()
    {
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\AL_" + strGUID + ".xls";

        DataSet newdataset;
        DB clsDB = new DB();
        String lsSQL = getFinalSp(ReportType.AllocationGroupPieChart);
        newdataset = clsDB.getDataSet(lsSQL);
        DataTable table = newdataset.Tables[0];
        AppDomain.CurrentDomain.Load("JCommon");
        org.jfree.data.general.DefaultPieDataset myDataSet = new org.jfree.data.general.DefaultPieDataset();
        if (table.Rows.Count > 0)
        {
            if (Convert.ToString(table.Rows[0][2]) != "0")
                myDataSet.setValue("Cash and Equivalents", Math.Round(Convert.ToDouble(table.Rows[0][2]), 1));
            if (Convert.ToString(table.Rows[1][2]) != "0")
                myDataSet.setValue("Fixed Income", Math.Round(Convert.ToDouble(table.Rows[1][2]), 1));
            if (Convert.ToString(table.Rows[2][2]) != "0")
                myDataSet.setValue("Domestic Equity", Math.Round(Convert.ToDouble(table.Rows[2][2]), 1));
            if (Convert.ToString(table.Rows[3][2]) != "0")
                myDataSet.setValue("International Equity", Math.Round(Convert.ToDouble(table.Rows[3][2]), 1));
            if (Convert.ToString(table.Rows[4][2]) != "0")
                myDataSet.setValue("Global Opportunistic", Math.Round(Convert.ToDouble(table.Rows[4][2]), 1));
            if (Convert.ToString(table.Rows[5][2]) != "0")
                myDataSet.setValue("Hedged Strategies", Math.Round(Convert.ToDouble(table.Rows[5][2]), 1));
            if (Convert.ToString(table.Rows[7][2]) != "0")
                myDataSet.setValue("Liquid Real Assets", Math.Round(Convert.ToDouble(table.Rows[7][2]), 1));
            if (Convert.ToString(table.Rows[8][2]) != "0")
                myDataSet.setValue("Illiquid Real Assets", Math.Round(Convert.ToDouble(table.Rows[8][2]), 1));
            if (Convert.ToString(table.Rows[9][2]) != "0")
                myDataSet.setValue("Private Equitys", Math.Round(Convert.ToDouble(table.Rows[9][2]), 1));
        }
        JFreeChart pieChart = ChartFactory.createPieChart3D(lsFamiliesName + "\n" + lsDateName, myDataSet, false, true, false);

        pieChart.setBackgroundPaint(java.awt.Color.white);
        pieChart.setBorderVisible(false);

        pieChart.setTitle(new org.jfree.chart.title.TextTitle("", new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12)));
        //TextTitle subtitle = new TextTitle("ALLOCATION GROUP PIE CHART");
        //subtitle.setFont(new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 11));
        //TextTitle subtitleDate = new TextTitle("lsDateName");
        //subtitleDate.setFont(new java.awt.Font("Frutiger55", java.awt.Font.ITALIC, 9));
        //pieChart.addSubtitle(subtitle);
        //pieChart.addSubtitle(subtitleDate);

        PiePlot ColorConfigurator = (PiePlot)pieChart.getPlot();
        ColorConfigurator.setLabelBackgroundPaint(System.Drawing.Color.White);// ColorConfigurator.getLabelPaint()
        ColorConfigurator.setLabelOutlinePaint(System.Drawing.Color.White);
        ColorConfigurator.setLabelShadowPaint(System.Drawing.Color.White);
      
        ColorConfigurator.setLabelFont(new System.Drawing.Font("Frutiger55", 10));
        ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardPieSectionLabelGenerator("{0} =  {1}%"));
        //ColorConfigurator.setLabelGenerator(new CustomPieSectionLabelGenerator());
        ColorConfigurator.setCircular(false);

        java.util.List keys = myDataSet.getKeys();

        for (int i = 0; i < keys.size(); i++)
        {
            if (i == 0)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4271A5"));
            if (i == 1)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#AD4542"));
            if (i == 2)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#84A24A"));
            if (i == 3)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#73558C"));
            if (i == 4)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4296AD"));
            if (i == 5)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#DE8239"));
            if (i == 6)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#94A6CE"));
            if (i == 7)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CE9294"));
            if (i == 8)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#B5CB94"));
        }
        ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
        java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
        ChartUtilities.writeChartAsPNG(jos, pieChart, 800, 350);
        //CreatePDFFile();
    
        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        lsTotalNumberofColumns = 4 + "";
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(4, 4);   // 2 rows, 2 columns        
        setTableProperty(loTable);

        if (ddlHousehold.SelectedValue != "0")
        {
            if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
            {
                lsFamiliesName = ddlAllocationGroup.SelectedItem.Text;
            }
            else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
            {
                lsFamiliesName = drpHouseHoldReportTitle.SelectedItem.Text;
            }
            else
            {
                lsFamiliesName = drpAllocationGroupTitle.SelectedItem.Text;
            }
        }
        if (txtAsofdate.Text != "")
            lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";

        Chunk lochunk = new Chunk(lsFamiliesName, setFontsAll(12, 1, 0));
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + "ALLOCATION GROUP PIE CHART", setFontsAll(11, 0, 0));

        loCell.Add(lochunk);
        
        loCell.HorizontalAlignment = 1;

        lochunk = new Chunk("\n" + lsDateName, setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        loCell.Add(lochunk);
        loCell.Border = 0;
        loCell.Colspan = 4;
        //   loCell.Add(loParagraph);
        loTable.AddCell(loCell);

        document.Add(loTable);

        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(fsFinalLocation.Replace(".xls", ".png"));

        //Give space before image
        jpg.SpacingBefore = 2f;
        //Give some space after the image
        //jpg.SpacingAfter = 1f;

        document.Add(jpg); //add an image to the created pdf document
        document.Close();
        FileInfo loFile = new FileInfo(ls);
        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    public string generateOverAllPieChart()
    {
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";

        DataSet newdataset;
        DB clsDB = new DB();
        String lsSQL = getFinalSp(ReportType.OverallPieChart);
        newdataset = clsDB.getDataSet(lsSQL);
        DataTable table = newdataset.Tables[0];
        AppDomain.CurrentDomain.Load("JCommon");
        org.jfree.data.general.DefaultPieDataset myDataSet = new org.jfree.data.general.DefaultPieDataset();
        if (table.Rows.Count > 0)
        {
            if (Convert.ToString(table.Rows[0][2]) != "0")
                myDataSet.setValue("Cash and Equivalents", Math.Round(Convert.ToDouble(table.Rows[0][2]), 1)); 
            if (Convert.ToString(table.Rows[1][2]) != "0")
                myDataSet.setValue("Fixed Income", Math.Round(Convert.ToDouble(table.Rows[1][2]), 1)); 
            if (Convert.ToString(table.Rows[2][2]) != "0")
                myDataSet.setValue("Domestic Equity", Math.Round(Convert.ToDouble(table.Rows[2][2]), 1)); 
            if (Convert.ToString(table.Rows[3][2]) != "0")
                myDataSet.setValue("International Equity", Math.Round(Convert.ToDouble(table.Rows[3][2]), 1)); 
            if (Convert.ToString(table.Rows[4][2]) != "0")
                myDataSet.setValue("Global Opportunistic", Math.Round(Convert.ToDouble(table.Rows[4][2]), 1)); 
            if (Convert.ToString(table.Rows[5][2]) != "0")
                myDataSet.setValue("Hedged Strategies", Math.Round(Convert.ToDouble(table.Rows[5][2]), 1)); 
            if (Convert.ToString(table.Rows[7][2]) != "0")
                myDataSet.setValue("Liquid Real Assets", Math.Round(Convert.ToDouble(table.Rows[7][2]), 1)); 
            if (Convert.ToString(table.Rows[8][2]) != "0")
                myDataSet.setValue("Illiquid Real Assets", Math.Round(Convert.ToDouble(table.Rows[8][2]), 1)); 
            if (Convert.ToString(table.Rows[9][2]) != "0")
                myDataSet.setValue("Private Equitys", Math.Round(Convert.ToDouble(table.Rows[9][2]), 1));
        }
        JFreeChart pieChart = ChartFactory.createPieChart3D(lsFamiliesName + "\n" + lsDateName, myDataSet, false, true, false);

        pieChart.setBackgroundPaint(java.awt.Color.white);
        pieChart.setBorderVisible(false);

        pieChart.setTitle(new org.jfree.chart.title.TextTitle("", new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12)));
        //TextTitle subtitle = new TextTitle("ALLOCATION GROUP PIE CHART");
        //subtitle.setFont(new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 11));
        //TextTitle subtitleDate = new TextTitle("lsDateName");
        //subtitleDate.setFont(new java.awt.Font("Frutiger55", java.awt.Font.ITALIC, 9));
        //pieChart.addSubtitle(subtitle);
        //pieChart.addSubtitle(subtitleDate);

        PiePlot ColorConfigurator = (PiePlot)pieChart.getPlot();
        ColorConfigurator.setLabelBackgroundPaint(System.Drawing.Color.White);// ColorConfigurator.getLabelPaint()
        ColorConfigurator.setLabelOutlinePaint(System.Drawing.Color.White);
        ColorConfigurator.setLabelShadowPaint(System.Drawing.Color.White);

        ColorConfigurator.setLabelFont(new System.Drawing.Font("Frutiger55", 10));

        ColorConfigurator.setCircular(false);
        ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardPieSectionLabelGenerator("{0} =  {1}%"));
        // ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardPieSectionLabelGenerator("{0}", new java.text.DecimalFormat("#,##0.00"), java.text.NumberFormat.getPercentInstance()));
        //ColorConfigurator.setLabelGenerator(new CustomPieSectionLabelGenerator());

        java.util.List keys = myDataSet.getKeys();

        for (int i = 0; i < keys.size(); i++)
        {
            if (i == 0)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4271A5"));
            if (i == 1)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#AD4542"));
            if (i == 2)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#84A24A"));
            if (i == 3)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#73558C"));
            if (i == 4)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4296AD"));
            if (i == 5)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#DE8239"));
            if (i == 6)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#94A6CE"));
            if (i == 7)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CE9294"));
            if (i == 8)
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#B5CB94"));
        }
        ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
        java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
        ChartUtilities.writeChartAsPNG(jos, pieChart, 800, 350);
        //CreatePDFFile();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        if (ddlHousehold.SelectedValue != "0")
        {
            if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
            {
                lsFamiliesName = ddlAllocationGroup.SelectedItem.Text;
            }
            else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
            {
                lsFamiliesName = drpHouseHoldReportTitle.SelectedItem.Text;
            }
            else
            {
                lsFamiliesName = drpAllocationGroupTitle.SelectedItem.Text;
            }
        }
        if (txtAsofdate.Text != "")
            lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";

        lsTotalNumberofColumns = 4 + "";
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(4, 4);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk lochunk = new Chunk(lsFamiliesName, setFontsAll(12, 1, 0));
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + "Overall Pie Chart", setFontsAll(11, 0, 0));

        loCell.Add(lochunk);
        
        //loCell.HorizontalAlignment = Align.CENTER;
        loCell.HorizontalAlignment = 1;

        lochunk = new Chunk("\n" + lsDateName, setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        loCell.Add(lochunk);
        loCell.Border = 0;
        loCell.Colspan = 4;
        //   loCell.Add(loParagraph);
        loTable.AddCell(loCell);

        document.Add(loTable);
        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png");
       
        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(fsFinalLocation.Replace(".xls", ".png"));

        //Give space before image
        jpg.SpacingBefore = 2f;
        //Give some space after the image
        //jpg.SpacingAfter = 1f;

        document.Add(jpg); //add an image to the created pdf document
        document.Close();
        FileInfo loFile = new FileInfo(ls);
        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    public string generatePortfolioConstChart()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = String.Empty;
        //String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);
        String lsSQL = getFinalSp(ReportType.PortfolioConsChart);
        // Response.Write(lsSQL);
        newdataset = clsDB.getDataSet(lsSQL);
        DataTable table = newdataset.Tables[0];
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\PC_" + strGUID + ".xls";

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        lsTotalNumberofColumns = 13 + "";
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(13, 15);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        setTableProperty(loTable);

        int liTotalPage = 1;// (newdataset.Tables[0].Rows.Count / liPageSize);
        int liCurrentPage = 0;
        liPageSize = 38;

        iTextSharp.text.Chunk lochunk = new Chunk();

        if (table.Rows.Count > 0)
        {
            Double DirectionalPerc = Math.Round(Convert.ToDouble(table.Rows[2][2]) + Convert.ToDouble(table.Rows[3][2]), 1);
            Double NonDirPerc = Math.Round(Convert.ToDouble(table.Rows[4][2]) + Convert.ToDouble(table.Rows[5][2]), 1);
            string DirectValue = String.Format("{0:0,0}", Convert.ToDouble(table.Rows[2][1]) + Convert.ToInt32(table.Rows[3][1]));
            string NonDirectValue = String.Format("{0:0,0}", Convert.ToDouble(table.Rows[4][1]) + Convert.ToInt32(table.Rows[5][1]));
            Double TotalPerc = Math.Round(DirectionalPerc + NonDirPerc, 1);
            string CorePortfTotal = String.Format("{0:0,0}", Convert.ToDouble(DirectValue) + Convert.ToDouble(NonDirectValue));

            double CashValue = Convert.ToDouble(table.Rows[0][1]);
            double FixedIncomeValue = Convert.ToDouble(table.Rows[1][1]);
            double DomesticEqueityValue = Convert.ToDouble(table.Rows[2][1]);
            double GlobalOppValue = Convert.ToDouble(table.Rows[4][1]);
            double InternationEqValue = Convert.ToDouble(table.Rows[3][1]);
            double HedgeValue = Convert.ToDouble(table.Rows[5][1]);
            double ConcentratedValue = Convert.ToDouble(table.Rows[6][1]);
            double LiquidValue = Convert.ToDouble(table.Rows[7][1]);
            double ILiquidValue = Convert.ToDouble(table.Rows[8][1]);
            double PrivateEqValue = Convert.ToDouble(Convert.ToDouble(table.Rows[9][1]));
            int CashPerc = Convert.ToInt32(Math.Round(Convert.ToDouble(table.Rows[0][2])));
            int FixedIncomPerc = Convert.ToInt32(Math.Round(Convert.ToDouble(table.Rows[1][2])));
            int ConcPerc = Convert.ToInt32(Math.Round(Convert.ToDouble(table.Rows[6][2])));
            int LiquidPerc = Convert.ToInt32(Math.Round(Convert.ToDouble(table.Rows[7][2])));
            int ILLiquidPerc = Convert.ToInt32(Math.Round(Convert.ToDouble(table.Rows[8][2])));
            int PrivatePerc = Convert.ToInt32(Math.Round(Convert.ToDouble(table.Rows[9][2])));

            Double rtgh = CashValue + FixedIncomeValue + Convert.ToDouble(CorePortfTotal) + ConcentratedValue + LiquidValue;
            string LiquidAssetValue = String.Format("{0:0,0}", rtgh);
            string ILLiquidAssetValue = String.Format("{0:0,0}", ILiquidValue + PrivateEqValue);
            Double LiquidAssetPerc = 0;     // TotalPerc + CashPerc + FixedIncomPerc + ConcPerc + LiquidPerc;
            Double ILLiquidAssetPerc = 0;  // ILLiquidPerc + PrivatePerc;

            Double dLiquidAssetPerc = Math.Round(Convert.ToDouble(table.Rows[2][2]) + Convert.ToDouble(table.Rows[3][2]) + Convert.ToDouble(table.Rows[4][2]) + Convert.ToDouble(table.Rows[5][2]) + Convert.ToDouble(table.Rows[0][2]) + Convert.ToDouble(table.Rows[1][2]) + Convert.ToDouble(table.Rows[6][2]) + Convert.ToDouble(table.Rows[7][2]), 1);
            Double dILLiquidAssetPerc = Math.Round(Convert.ToDouble(table.Rows[8][2]) + Convert.ToDouble(table.Rows[9][2]), 1);
            LiquidAssetPerc = dLiquidAssetPerc;
            ILLiquidAssetPerc = dILLiquidAssetPerc;

            string FinalValue = String.Format("{0:0,0}", Convert.ToDouble(LiquidAssetValue) + Convert.ToDouble(ILLiquidAssetValue));

            if (txtAsofdate.Text != "")
                lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";

            for (int i = 0; i < 16; i++)
            {
                int colsize = 13;
                for (int j = 0; j < colsize; j++)
                {
                    // string Text = "i=" + i.ToString() +": j="+ j.ToString()+",";// Convert.ToString(newdataset.Tables[0].Rows[i][j]);
                    string Text = "";// i.ToString() + ":" + j.ToString();
                    string lsfamilyName = "";
                    //if (ddlAllocationGroup.SelectedValue != "0")
                    //{
                    //    lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
                    //}
                    if (ddlHousehold.SelectedValue != "0")
                    {
                        if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
                        {
                            lsFamiliesName = ddlAllocationGroup.SelectedItem.Text;
                        }
                        else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
                        {
                            lsFamiliesName = drpHouseHoldReportTitle.SelectedItem.Text;
                        }
                        else
                        {
                            lsFamiliesName = drpAllocationGroupTitle.SelectedItem.Text;
                        }
                    }
                    if (i == 0 && j == 0)
                    {
                        lochunk = new Chunk(lsFamiliesName + Text + "", setFontsAll(12, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Colspan = 13;// newdataset.Tables[0].Columns.Count;
                        j = j + 12;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        //liCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.Leading = 10F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 1 && j == 0)
                    {
                        lochunk = new Chunk("PORTFOLIO CONSTRUCTION" + Text + "", setFontsAll(11, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        lochunk = new Chunk("\n" + lsDateName, setFontsAll(9, 0, 1));
                        loCell.Add(lochunk);

                        loCell.Colspan = 13;// newdataset.Tables[0].Columns.Count;
                        j = j + 12;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
                        loCell.Leading = 11F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 0)
                    {
                        lochunk = new Chunk("Cash" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        SetBorder(loCell, false, true, false, false);
                        loCell.Add(lochunk);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 2)
                    {
                        lochunk = new Chunk("Fixed Income" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 2;// newdataset.Tables[0].Columns.Count;

                        SetBorder(loCell, false, true, false, false);
                        //liCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        // loCell.Add(lochunk);
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 4)
                    {
                        lochunk = new Chunk("Core Portfolio	" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Colspan = 2;// newdataset.Tables[0].Columns.Count;
                        j = j + 1;
                        SetBorder(loCell, false, true, false, false);
                        //liCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        // loCell.Add(lochunk);
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 7)
                    {
                        lochunk = new Chunk("Concentrated" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, false, true, false, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 9)
                    {
                        lochunk = new Chunk("Real Assets" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Colspan = 2;
                        j = j + 1;
                        SetBorder(loCell, false, true, false, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 12)
                    {
                        lochunk = new Chunk("Private Equity" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 2;
                        SetBorder(loCell, false, true, false, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 4)
                    {
                        lochunk = new Chunk("Directional" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 3;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 5)
                    {
                        lochunk = new Chunk("Non-Directional" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 3;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 0)
                    {
                        lochunk = new Chunk("Cash\n\n\n" + Math.Round(Convert.ToDouble(table.Rows[0][2]), 1) + Text + "%", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, false, true, true);
                        //loCell.BorderWidthTop = 0.2F;
                        //loCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        //loCell.BorderWidthLeft = 0.2F;
                        //loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        //loCell.BorderWidthRight = 0.2F;
                        //loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        //iTextSharp.text.Color objColor = new iTextSharp.text.Color(215,228,118,255);
                        //objColor.
                        //loCell.BackgroundColor =System.Draw//
                        // PdfSpotColor objCol= new PdfSpotColor("pdfcolor",11,new CMYKColor(3f,0f,8f,5f));
                        //PdfContentByte canvas= new PdfContentByte(
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#EAF1DD"));
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 0)//Cash Below
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#EAF1DD"));
                        SetBorder(loCell, false, true, true, true);
                        //loCell.BorderWidthBottom = 1F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        //loCell.BorderWidthLeft = 0.2F;
                        //loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        //loCell.BorderWidthRight = 0.2F;
                        //loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 2)
                    {
                        lochunk = new Chunk("Fixed Income\n\n\n" + Math.Round(Convert.ToDouble(table.Rows[1][2]), 1) + Text + "%", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#D7E4BC"));
                        SetBorder(loCell, true, false, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 2)//Fixed Income Below
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#D7E4BC"));
                        SetBorder(loCell, false, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 4)
                    {
                        string Domequity = String.Format("{0:0,0}", Convert.ToDouble(table.Rows[2][1]));
                        lochunk = new Chunk("Domestic Equity\n\n" + Math.Round(Convert.ToDouble(table.Rows[2][2]), 1) + Text + "%\n$" + Domequity + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B8CCE4"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 5)
                    {
                        string GlobOpp = String.Format("{0:0,0}", Convert.ToDouble(table.Rows[4][1]));
                        lochunk = new Chunk("Global Opportunistic\n\n" + Math.Round(Convert.ToDouble(table.Rows[4][2]), 1) + Text + "%\n$" + GlobOpp + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B8CCE4"));

                        SetBorder(loCell, true, true, false, true);
                        //SetBorder(loCell, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 7)
                    {
                        lochunk = new Chunk("Concentrated Positions\n\n" + Math.Round(Convert.ToDouble(table.Rows[6][2]), 1) + Text + "%", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#DBE5F1"));

                        SetBorder(loCell, true, false, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 7)//Concetrated Beloww
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#DBE5F1"));

                        SetBorder(loCell, false, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 9)
                    {
                        lochunk = new Chunk("Liquid Real Assets\n\n" + Math.Round(Convert.ToDouble(table.Rows[7][2]), 1) + Text + "%", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#E5E0EC"));
                        SetBorder(loCell, true, false, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 9)// Liquid Real Assets Below
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#E5E0EC"));
                        SetBorder(loCell, false, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 10)
                    {
                        lochunk = new Chunk("Illiquid Real Assets\n\n" + Math.Round(Convert.ToDouble(table.Rows[8][2]), 1) + Text + "%", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, false, false, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 10)//Illiquid Real Assets Below
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, false, true, false, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 12)
                    {
                        lochunk = new Chunk("Private Equity\n\n\n" + Math.Round(Convert.ToDouble(table.Rows[9][2]), 1) + Text + "%", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#DDD9C3"));
                        SetBorder(loCell, true, false, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 12)//rivate Equity Below
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#DDD9C3"));
                        SetBorder(loCell, false, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }

                    else if (i == 5 && j == 4) //International Equity Heading
                    {
                        string Intquity = String.Format("{0:0,0}", Convert.ToDouble(table.Rows[3][1]));
                        lochunk = new Chunk("International Equity\n\n" + Math.Round(Convert.ToDouble(table.Rows[3][2]), 1) + Text + "%\n$" + Intquity + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B8CCE4"));

                        SetBorder(loCell, false, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 5) //Hedged Strategies Heading
                    {
                        string HedgedStr = String.Format("{0:0,0}", Convert.ToDouble(table.Rows[5][1]));
                        lochunk = new Chunk("Hedged Strategies\n\n" + Math.Round(Convert.ToDouble(table.Rows[5][2]), 1) + Text + "%\n$" + HedgedStr + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B8CCE4"));
                        SetBorder(loCell, false, true, false, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }

                    else if (i == 6 && j == 4) //Directional %
                    {
                        lochunk = new Chunk(DirectionalPerc + "%" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#F7F3F7"));
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        //loCell.Colspan = 3;
                        loCell.BorderWidthTop = 2F;
                        loCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 5) //Non Directional %
                    {
                        lochunk = new Chunk(NonDirPerc + "%" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#F7F3F7"));
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        //loCell.Colspan = 3;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 7 && j == 4) // Directional Value
                    {
                        lochunk = new Chunk("$" + DirectValue + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        // loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#F7F3F7"));
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        // loCell.Colspan = 3;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 7 && j == 5) // Non Directional Value
                    {
                        lochunk = new Chunk("$" + NonDirectValue + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#F7F3F7"));
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        //loCell.Colspan = 3;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 8 && j == 4) //Core Portfolio %
                    {
                        lochunk = new Chunk(TotalPerc + "%" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        // loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#F7F3F7"));
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 2;
                        j = j + 1;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 9 && j == 0) //Cash Value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:0,0}", Convert.ToDouble(table.Rows[0][1])) + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 9 && j == 2)//Fixed Income Value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:0,0}", Convert.ToDouble(table.Rows[1][1])) + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 9 && j == 4)// Core Portfolio value
                    {
                        lochunk = new Chunk("$" + CorePortfTotal + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 2;
                        j = j + 1;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 9 && j == 7)// Concentrated value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:0,0}", Convert.ToDouble(table.Rows[6][1])) + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 9 && j == 9)// Liquid Real Assets value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:0,0}", Convert.ToDouble(table.Rows[7][1])) + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        //loCell.BorderWidthBottom = 1F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 9 && j == 10)// Illiquid Real Assets value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:0,0}", Convert.ToDouble(table.Rows[8][1])) + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        //loCell.BorderWidthBottom = 1F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 9 && j == 12)// Private equity value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:0,0}", Convert.ToDouble(table.Rows[9][1])) + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 10 && j == 0)// | border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 10;
                        j = j + 9;
                        loCell.BorderWidthLeft = 1F;
                        loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthBottom = 1F;
                        loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.fi = 5.2;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 10 && j == 10)// | border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 3;
                        j = j + 2;
                        loCell.BorderWidthLeft = 1F;
                        loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthBottom = 1F;
                        loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        loTable.AddCell(loCell);
                    }
                    //liquid asset above | border start
                    else if (i == 11 && j == 0)// -- border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 5;
                        j = j + 4;
                        //loCell.BorderWidthTop = 1F;
                        //loCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 11 && j == 5)// -- border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 6;
                        j = j + 5;
                        //loCell.BorderWidthLeft = 1F;
                        //loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 11 && j == 11)// -- border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 2;
                        j = j + 1;
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    //liquid asset above | border end
                    else if (i == 12 && j == 0)// Liquid Assets Heading
                    {
                        lochunk = new Chunk("Liquid Assets\n$" + LiquidAssetValue + "\n" + LiquidAssetPerc + "" + Text + "%", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 10;
                        j = j + 9;
                        loCell.Border = 0;
                        //loCell.BorderWidthBottom = 1F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 12 && j == 10)// illLiquid Assets Heading
                    {
                        lochunk = new Chunk("Illiquid Assets\n$" + ILLiquidAssetValue + "\n" + ILLiquidAssetPerc + "%" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 3;
                        j = j + 2;
                        loCell.Border = 0;
                        //loCell.BorderWidthBottom = 1F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }

                    else if (i == 13 && j == 0)// -- border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 13;
                        j = j + 12;
                        loCell.BorderWidthLeft = 1F;
                        loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthBottom = 1F;
                        loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        loTable.AddCell(loCell);
                    }

                    else if (i == 14 && j == 0)// -- border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 6;
                        j = j + 5;
                        //loCell.BorderWidthLeft = 1F;
                        //loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        // loCell.Height =iTextSharp.text.Cell.
                        loTable.AddCell(loCell);
                    }
                    else if (i == 14 && j == 6)// -- border
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 7;
                        j = j + 6;
                        //loCell.BorderWidthLeft = 1F;
                        //loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.BorderWidthRight = 1F;
                        loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        // loCell.Height =iTextSharp.text.Cell.
                        loTable.AddCell(loCell);
                    }
                    else if (i == 15 && j == 0)// Total Portfolio Heading
                    {
                        lochunk = new Chunk("Total Portfolio\n$" + FinalValue + "" + Text + "", setFontsAll(10, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 13;
                        j = j + 12;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }

                    else
                    {
                        lochunk = new Chunk(Text, setFontsAll(10, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 2;// newdataset.Tables[0].Columns.Count;

                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        //liCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        // loCell.Add(lochunk);
                        loTable.AddCell(loCell);
                    }
                }
            }
            document.Add(loTable);
            document.Close();

            FileInfo loFile = new FileInfo(ls);
            loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            //Response.Write("<script>");
            //string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
            //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
            //Response.Write("</script>");
            return fsFinalLocation.Replace(".xls", ".pdf");
        }
        else
        {
            lblError.Text = "Record not found";
            return "Record not found";
        }
        //document.Add(loTable);

        //if (newdataset.Tables[0].Rows.Count > 0)
        //{
        //    document.Close();

        //    FileInfo loFile = new FileInfo(ls);
        //    try
        //    {
        //        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

        //        Response.Write("<script>");
        //        string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
        //        Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        //        Response.Write("</script>");
        //        return fsFinalLocation.Replace(".xls", ".pdf");
        //    }
        //    catch (Exception exc)
        //    {
        //        Response.Write(exc.Message);
        //    }
        //}
    }

    public string generateCommittmentSchReport()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = "*We have not included remaining commitments for certain funds where the General Partner has indicated further capital calls are highly unlikely.  However, it is possible that small capital calls may be made in the future by certain of these General Partners.";
        //String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);
        String lsSQL = getFinalSp(ReportType.CommitmentSchedule);
        // Response.Write(lsSQL);
        newdataset = clsDB.getDataSet(lsSQL);

        newdataset = AddTotals(newdataset);
        DataTable table = newdataset.Tables[0].Clone();

        for (int i = 0; i < newdataset.Tables.Count; i++)
        {
            if (newdataset.Tables[i].Rows.Count > 0)
            {
                if (i == 0)
                {
                    if (newdataset.Tables[0].Rows.Count > 1)
                    {
                        if (newdataset.Tables[1].Rows.Count < 1)
                        {
                            for (int l = 0; l < newdataset.Tables[0].Rows.Count; l++)
                            {
                                if (Convert.ToString(newdataset.Tables[0].Rows[l]["_OrderNmb"]) == "3")
                                {
                                    newdataset.Tables.Remove(newdataset.Tables[1]);
                                    //newdataset.Tables[0].Rows[l].BeginEdit();
                                    ////newdataset.Tables[0].Rows[l]["Investment"] = "";
                                    //newdataset.Tables[0].Rows[l]["Commitment"] = 0.00;
                                    //newdataset.Tables[0].Rows[l]["CalledToDate"] = 0.00;
                                    //newdataset.Tables[0].Rows[l]["ReinvestedDistributionToDate"] = 0.00;
                                    //newdataset.Tables[0].Rows[l]["TotalInvestedToDate"] = 0.00;
                                    //newdataset.Tables[0].Rows[l]["RemainingCommitment"] = 0.00;
                                    //newdataset.Tables[0].Rows[l]["ExpectedRemaining"] = 0.00;
                                    //newdataset.Tables[0].Rows[l]["ExpectedRemainingCallsCurrentQuarter"] = 0.00;
                                    //newdataset.Tables[0].Rows[l]["ExpectedRemainingCallsNextQuarter"] = 0.00;

                                    //newdataset.Tables[0].Rows[l].EndEdit();
                                    newdataset.Tables[0].AcceptChanges();

                                    table.Merge(newdataset.Tables[i], true);
                                    table.Rows.Add(table.NewRow());// Add new Blank row
                                }
                            }
                        }
                        else
                        {
                            table.Merge(newdataset.Tables[i], true);

                            table.Rows.Add(table.NewRow());// Add new Blank row

                            DataRow newHeaderRow = table.NewRow();
                            newHeaderRow["Investment"] = "PROPOSED/CONFIRMED - PENDING CLOSE";

                            newHeaderRow["_OrderNmb"] = 6;//to style new row

                            table.Rows.Add(newHeaderRow);
                        }
                    }
                }
                else if (i == 1) // for table 1
                {
                    //DataRow newHeaderRow = table.NewRow();
                    //newHeaderRow["Investment"] = "Proposed/Confirmed - Pending Close";
                    //newHeaderRow["_OrderNmb"] = 6;//to style new row
                    ////table.Rows.Add(newHeaderRow);
                    table.Merge(newdataset.Tables[i], true);
                    //table.Rows.Add(table.NewRow());// Add new Blank row

                }
                else //if(newdataset.Tables[i+1].Rows.Count >0)
                {
                    //table.Rows.Add(table.NewRow());// Add new Blank row
                    DataRow rowStyle = table.NewRow();
                    rowStyle["_OrderNmb"] = 8;
                    table.Rows.Add(rowStyle);


                    table.Merge(newdataset.Tables[i], true);
                    table.Rows.Add(table.NewRow());// Add new Blank row

                    DataRow newHeaderRow = table.NewRow();
                    newHeaderRow["Investment"] = "PROPOSED/CONFIRMED - PENDING CLOSE";

                    newHeaderRow["_OrderNmb"] = 6;//to style new row

                    table.Rows.Add(newHeaderRow);

                }

            }
            else if (newdataset.Tables[i].Rows.Count < 1)
            {
                if (table.Rows.Count > 0)
                {
                    table.Rows[table.Rows.Count - 1].Delete();
                }

            }

        }

        if (table.Rows.Count > 0)
        {
            table.Rows[table.Rows.Count - 1].Delete();
        }

        table.AcceptChanges();

        if (table.Rows.Count < 3)
        {
            lblError.Text = "Record not found";
            return "Record not found";
        }
        //lodataset.Tables.Add(table);
        DataTable loInsertblankRow = table.Copy();
        //loInsertblankRow.Tables.Add(table);
        //lodataset.Tables[0].Clear();
        table.Clear();
        table = null;
        table = loInsertblankRow.Clone();

        // string strGUID = Guid.NewGuid().ToString();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        // strGUID = strGUID.Substring(0, 5);
        // String fsFinalLocation = @"C:\Reports\" + strGUID + ".xls";

        String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";
        int liBlankCounter = 0;



        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Rows.Count; liBlankRow++)
        {
            if (liBlankRow != 0 && loInsertblankRow.Rows[liBlankRow]["_OrderNmb"].ToString() == "2" || loInsertblankRow.Rows[liBlankRow]["_OrderNmb"].ToString() == "3")
            {
                //if (!String.IsNullOrEmpty(fsSPriorDate) && loInsertblankRow.Tables[0].Rows.Count - 1 != liBlankRow)
                if (loInsertblankRow.Rows.Count - 1 != liBlankRow)
                {
                    // DataRow newCustomersRow = loInsertblankRow.NewRow();
                    // newCustomersRow[0] = "";
                    //// newCustomersRow[1] = 100.00;
                    // loInsertblankRow.Rows.Add(newCustomersRow);
                    // liBlankCounter = liBlankCounter + 1;
                }

            }

            if (loInsertblankRow.Rows.Count > 1)
            {
                table.ImportRow(loInsertblankRow.Rows[liBlankRow]);
            }
        }
        table.AcceptChanges();
        DataSet lodataset = new DataSet();
        lodataset.Tables.Add(table);
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
        loTable.Cellpadding = 0f;
        loTable.Cellspacing = 0f;


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

                loInsertdataset.AcceptChanges();
                setHeader(document, loInsertdataset, newdataset);
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
                    if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1)
                    {
                        lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                    }
                    else
                    {
                        lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
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
                loCell.Leading = 4f;//6

                loCell.UseBorderPadding = true;

                //  if (lodataset.Tables[0].Rows[liRowCount]["_Ssi_TabFlg"].ToString() == "True" && lodataset.Tables[0].Rows[liRowCount]["_Ssi_UnderlineFlg"].ToString() != "True")


                if (liColumnCount != 0)
                {
                    loCell.HorizontalAlignment = 2;
                }


                //// new underline  code
                if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "1")
                {
                    loCell.BorderColorBottom = iTextSharp.text.Color.BLACK;
                }
                else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "0")
                {
                    loCell.DisableBorderSide(-1);
                }
                /////

                /*=========START WITH BOLD AND SUPERBOLD FLAG========*/
                if (checkTrue(lodataset, liRowCount, "_OrderNmb") || checkTrue(lodataset, liRowCount, "_OrderNmb"))
                {
                    lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                    try
                    {
                        if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1)
                        {
                            lsFormatedString = String.Format("${0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
                        }
                        else
                        {
                            lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                        }
                    }
                    catch
                    {

                    }

                    //changed on 02/25/2011
                    //lochunk = new Chunk(lsFormatedString, Font9Bold());
                    lochunk = new Chunk(lsFormatedString, Font8Bold());
                    #region Commented
                    if (!lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE"))
                    {
                        //changed on 02/25/2011
                        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                        lochunk = new Chunk(lsFormatedString, Font8Bold());
                        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                        if (lsFormatedString.Length > 25)
                        {
                            if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
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
                    #endregion

                }
                else
                {
                    if (liColumnCount == 0 && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                    {
                        String abc = "" + lodataset.Tables[0].Rows[liRowCount][0].ToString();
                        //changed on 02/25/2011
                        //lochunk = new Chunk(abc, Font9Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                        lochunk = new Chunk(abc, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));

                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "")
                        {
                            //loCell.EnableBorderSide(0);
                            lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "6")
                        {
                            checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(216, 216, 216));
                            lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]), setFontsAll(8, 1, 0));
                            //lochunk.SetBackground(iTextSharp.text.Color.LIGHT_GRAY);#B7DDE8 new iTextSharp.text.Color(216, 216, 216)
                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "3")
                        {

                            if (liRowCount == lodataset.Tables[0].Rows.Count - 5)
                            {
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loCell.VerticalAlignment = 4;
                                loCell.Leading = 10f;
                                lsFormatedString = "TOTAL PROPOSED/CONFIRMED COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                            }
                            else if (liRowCount == lodataset.Tables[0].Rows.Count - 4)
                            {
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loCell.VerticalAlignment = 4;
                                loCell.Leading = 10f;
                                lsFormatedString = "TOTAL PROPOSED/CONFIRMED COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                            }
                            else //if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                            {
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loCell.VerticalAlignment = 4;
                                loCell.Leading = 10f;
                                lsFormatedString = "TOTAL COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                            }

                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "2")
                        {
                            if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                            {
                                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 0, 0));
                            }
                            else
                            {
                                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 0, 0));
                            }

                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "7")
                        {
                            if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                            {
                                loCell.VerticalAlignment = 4;
                                loCell.Leading = 10f;
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                lsFormatedString = "TOTAL COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                            }
                            else
                            {
                                //loCell.EnableBorderSide(1);
                                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 1, 0));
                            }

                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "8")
                        {
                            loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                            //lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                        }
                    }
                    else if (liColumnCount != 0 && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                    {
                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "2")
                        {
                            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                            {
                                try
                                {
                                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                    lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                }
                                catch
                                {

                                }
                            }
                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "3")
                        {
                            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                            {
                                try
                                {
                                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                    lochunk = new Chunk(lsFormatedString, Font19Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                }
                                catch
                                {

                                }
                            }
                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "6")
                        {

                            try
                            {
                                checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(216, 216, 216));
                                lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]), setFontsAll(9, 1, 0));
                                //lochunk.SetBackground(iTextSharp.text.Color.LIGHT_GRAY);
                            }
                            catch
                            { }
                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "7")
                        {
                            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                            {
                                try
                                {
                                    //loCell.EnableBorderSide(1);

                                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                    lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                }
                                catch
                                {

                                }
                            }
                        }
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "8")
                        {
                            loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                            //lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                        }


                    }



                }
                if (checkTrue(lodataset, liRowCount, "_OrderNmb") && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                {
                    if (liColumnCount == 0)
                    {
                        String abc = "          " + "          " + lodataset.Tables[0].Rows[liRowCount][0].ToString();
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
                //checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(183, 221, 232));
                //====added on 28Feb2011 to change font size for total====
                if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
                {
                    if (liColumnCount != 0)
                    {
                        lochunk = new Chunk(lsFormatedString, Font7Bold());
                    }
                }
                /*=====END=====*/

                if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
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

            try
            {
                if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
                {
                    document.Add(loTable);
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                }
            }
            catch (Exception Ex)
            {

            }
        }
        document.Close();

        FileInfo loFile = new FileInfo(ls);
        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        return fsFinalLocation.Replace(".xls", ".pdf");
        //if (loInsertdataset.Tables[0].Rows.Count > 0)
        //{
        //    document.Close();

        //    FileInfo loFile = new FileInfo(ls);
        //    try
        //    {
        //        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

        //        Response.Write("<script>");
        //        string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
        //        Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        //        Response.Write("</script>");

        //        return fsFinalLocation.Replace(".xls", ".pdf");
        //    }
        //    catch (Exception exc)
        //    {
        //        Response.Write(exc.Message);
        //    }
        //}


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
            if (checkTrue(foDataset, fiRowCount, "_OrderNmb") || checkTrue(foDataset, fiRowCount, "_OrderNmb") || checkTrue(foDataset, fiRowCount, "_OrderNmb"))
            {
                setBottomWidthWhite(foCell);
            }
            if (checkTrue(foDataset, fiRowCount + 1, "_OrderNmb") || checkTrue(foDataset, fiRowCount + 1, "_OrderNmb") || checkTrue(foDataset, fiRowCount + 1, "_OrderNmb"))
            {
                setBottomWidthWhite(foCell);
            }
            else
            {
                foCell.BorderWidthBottom = 0.1F;
                foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);
                //foCell.BorderColorBottom = new iTextSharp.text.Color(121, 121, 121);
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


    public iTextSharp.text.Font setFontsAll1(int size, int bold)
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


        return font;
        #endregion
    }


    public void setHeaderInvestmentObjective(Document foDocument, DataSet loInsertdataset, DataSet loDatatset)
    {

        DataSet OldDataset = new DataSet();
        OldDataset = loDatatset.Copy();

        for (int liNewdataset = loInsertdataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        {
            if (loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
            }
        }

        DataSet AddDataset = loInsertdataset.Copy();
        AddDataset.AcceptChanges();
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(AddDataset.Tables[0].Columns.Count, 4);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk loParagraph = new Chunk();


        //////// set header new addition for pdf
        string lsfamilyName = "";
        if (ddlHousehold.SelectedValue != "0")
        {
            if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
            {
                lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
            }
            else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
            {
                lsfamilyName = drpHouseHoldReportTitle.SelectedItem.Text;
            }
            else
            {
                lsfamilyName = drpAllocationGroupTitle.SelectedItem.Text;
            }
        }




        if (txtAsofdate.Text != "")
            lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";

        /////////////

        //Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));

        Chunk lochunk = new Chunk(lsfamilyName, setFontsAll(12, 1, 0));
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + "Investment Policy Summary", setFontsAll(11, 0, 0));
        //loParagraph.Chunks.Add(lochunk);

        loCell.Add(lochunk);
        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        loCell.HorizontalAlignment = 1;



        lochunk = new Chunk("\n" + lsDateName, setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
        loTable.AddCell(loCell);

        Boolean lbCheckFoMarket = false;
        for (int liColumnCount = 0; liColumnCount < AddDataset.Tables[0].Columns.Count; liColumnCount++)
        {

            if (Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName) != "")
            {
                if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Current Portfolio %"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Current Portfolio %", "Current Allocation"), setFontsAll(7, 1, 0));
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Bench mark"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Bench mark", "Benchmark"), setFontsAll(7, 1, 0));
                }
                else
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName), setFontsAll(7, 1, 0));
                }


            }
            //}
            loCell = new Cell();

            loCell.Add(lochunk);
            loCell.Border = 0;

            loCell.NoWrap = true;//true;

            loCell.MaxLines = 2;
            loCell.Leading = -2F;
            if (liColumnCount != 0)
            {
                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
            }

            else
            {
                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
            }


            if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
            {
                loCell.Leading = 10f;//8
                loCell.MaxLines = 5;
                //loCell.Leading = 9f;
            }
            loCell.Leading = 10f;//8

            loCell.VerticalAlignment = 1; //5 ,6 bottom : WASTE VALUES - 3,4
            loTable.AddCell(loCell);

        }


        foDocument.Add(loTable);
        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        foDocument.Add(png);
    }

    public void setHeader(Document foDocument, DataSet loInsertdataset, DataSet loDatatset)
    {
        //iTextSharp.text.Table loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, 4);   // 2 rows, 2 columns        
        // DataSet OldDataset = loInsertdataset.Copy();

        DataTable table = loDatatset.Tables[0].Clone();
        for (int i = 0; i < loDatatset.Tables.Count; i++)
        {
            if (loDatatset.Tables[i].Rows.Count > 0)
            {

                if (i == 1)
                {
                    table.Rows.Add(table.NewRow());// Add new Blank row

                    //DataRow newHeaderRow = table.NewRow();
                    //newHeaderRow["Investment"] = "Proposed/Confirmed - Pending Close";

                    //newHeaderRow["_OrderNmb"] = 6;//to style new row

                    //table.Rows.Add(newHeaderRow);
                    table.Merge(loDatatset.Tables[i], true);


                }
                else
                {
                    table.Merge(loDatatset.Tables[i], true);

                    table.Rows.Add(table.NewRow());// Add new Blank row

                    DataRow newHeaderRow = table.NewRow();
                    newHeaderRow["Investment"] = "Proposed/Confirmed - Pending Close";

                    newHeaderRow["_OrderNmb"] = 6;//to style new row

                    table.Rows.Add(newHeaderRow);
                }
            }
            else if (loDatatset.Tables[i].Rows.Count < 1)
            {
                if (table.Rows.Count > 0)
                {
                    table.Rows[table.Rows.Count - 1].Delete();
                }
            }
        }

        if (table.Rows.Count > 0)
        {
            table.Rows[table.Rows.Count - 1].Delete();
        }

        table.AcceptChanges();

        if (table.Rows.Count < 2)
        {
            lblError.Text = "Record not found";
        }

        DataSet OldDataset = new DataSet();
        OldDataset.Tables.Add(table);

        for (int liNewdataset = loInsertdataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        {
            if (loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
            }
        }

        DataSet AddDataset = loInsertdataset.Copy();
        AddDataset.AcceptChanges();
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(AddDataset.Tables[0].Columns.Count, 4);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk loParagraph = new Chunk();


        //////// set header new addition for pdf
        string lsfamilyName = "";
        if (ddlHousehold.SelectedValue != "0")
        {
            if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
            {
                lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
            }
            else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
            {
                lsfamilyName = drpHouseHoldReportTitle.SelectedItem.Text;
            }
            else
            {
                lsfamilyName = drpAllocationGroupTitle.SelectedItem.Text;
            }
        }




        if (txtAsofdate.Text != "")
            lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";

        /////////////

        //Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));

        Chunk lochunk = new Chunk(lsfamilyName, setFontsAll(12, 1, 0));
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + "COMMITMENT SCHEDULE", setFontsAll(11, 0, 0));
        //loParagraph.Chunks.Add(lochunk);

        loCell.Add(lochunk);
        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        loCell.HorizontalAlignment = 1;



        lochunk = new Chunk("\n" + lsDateName, setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
        loTable.AddCell(loCell);

        if (loDatatset.Tables[0].Rows.Count > 0)
        {
            iTextSharp.text.Cell liCell = new Cell();
            Chunk lichunk = new Chunk(lsDistributionName, setFontsAll(10, 0, 0));
            lichunk = new Chunk("COMMITMENTS", setFontsAll(8, 1, 0));
            liCell.Add(lichunk);
            liCell.Colspan = loInsertdataset.Tables[0].Columns.Count;

            liCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
            liCell.Leading = 10f;
            liCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
            liCell.VerticalAlignment = 4;
            liCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
            loTable.AddCell(liCell);
        }
        else
        {
            iTextSharp.text.Cell liCell = new Cell();
            Chunk lichunk = new Chunk(lsDistributionName, setFontsAll(8, 0, 0));
            lichunk = new Chunk("PROPOSED/CONFIRMED - PENDING CLOSE", setFontsAll(8, 1, 0));
            liCell.Add(lichunk);
            liCell.Colspan = loInsertdataset.Tables[0].Columns.Count;

            liCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
            liCell.Leading = 10f;
            liCell.VerticalAlignment = 4;
            liCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
            liCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
            loTable.AddCell(liCell);
        }


        Boolean lbCheckFoMarket = false;
        for (int liColumnCount = 0; liColumnCount < AddDataset.Tables[0].Columns.Count; liColumnCount++)
        {
            //if (liColumnCount == 0)
            //{
            //    //changed on 02/25/2011
            //    //lochunk = new Chunk("", setFontsAll(9, 1, 0));
            //    lochunk = new Chunk("", setFontsAll(7, 1, 0));
            //}
            //else
            //{
            //changed on 02/25/2011
            if (Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName) != "")
            {
                if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("CalledToDate"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("CalledToDate", "Actual Calls to Date"), setFontsAll(7, 1, 0));
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("ReinvestedDistributionToDate"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ReinvestedDistributionToDate", "Reinvested Distribution"), setFontsAll(7, 1, 0));
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("TotalInvestedToDate"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("TotalInvestedToDate", "Total Amount Invested"), setFontsAll(7, 1, 0));
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("ExpectedRemaining"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemaining", "Expected Remaining Calls"), setFontsAll(7, 1, 0));
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("ExpectedRemainingCallsCurrentQuarter"))
                {
                    if (Convert.ToString(OldDataset.Tables[0].Rows[0]["_CurrentQuarter"]) != "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsCurrentQuarter", "Expected In Current Quarter"), setFontsAll(7, 1, 0));//Convert.ToString(OldDataset.Tables[0].Rows[0]["_CurrentQuarter"])
                    }
                    else if (Convert.ToString(OldDataset.Tables[0].Rows[1]["_CurrentQuarter"]) != "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsCurrentQuarter", "Expected In Current Quarter"), setFontsAll(7, 1, 0));
                    }
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("ExpectedRemainingCallsNextQuarter"))
                {
                    if (Convert.ToString(OldDataset.Tables[0].Rows[0]["_NextQuarter"]) != "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsNextQuarter", "Expected In Next Quarter"), setFontsAll(7, 1, 0));//Convert.ToString(OldDataset.Tables[0].Rows[0]["_NextQuarter"])
                    }
                    else if (Convert.ToString(OldDataset.Tables[0].Rows[1]["_NextQuarter"]) != "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsNextQuarter", "Expected In Next Quarter"), setFontsAll(7, 1, 0));
                    }
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("RemainingCommitment"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("RemainingCommitment", "Remaining Commitment"), setFontsAll(7, 1, 0));
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Investment"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Investment", "Investment*"), setFontsAll(7, 1, 0));
                }
                else
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName), setFontsAll(7, 1, 0));
                }
                //lochunk = RemainingCommitment new Chunk(Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Replace(" Market Value", ""), setFontsAll(9, 1, 0));

            }
            //}
            loCell = new Cell();

            loCell.Add(lochunk);
            loCell.Border = 0;

            loCell.NoWrap = true;//true;

            loCell.MaxLines = 2;
            loCell.Leading = -2F;
            if (liColumnCount != 0)
            {
                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            }
            else
            {
                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            }


            if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
            {
                loCell.Leading = 10f;//8
                loCell.MaxLines = 5;
                //loCell.Leading = 9f;
            }
            loCell.Leading = 10f;//8

            loCell.VerticalAlignment = 1;//5 ,6 bottom : WASTE VALUES - 3,4
            loTable.AddCell(loCell);

        }


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
            case "22":
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
                int[] headerwidths4 = { 15, 13, 13, 16 };
                fotable.SetWidths(headerwidths4);
                fotable.Width = 100;
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
                int[] headerwidths9 = { 27, 9, 9, 9, 9, 9, 9, 9, 7 };
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
            case "21":
                int[] headerwidths12 = { 7, 7, 12, 7, 7, 7, 7, 35, 7, 7, 7, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                fotable.SetWidths(headerwidths12);
                fotable.Width = 150; break;
            case "13":
                int[] headerwidths13 = { 10, 2, 15, 2, 20, 20, 2, 15, 5, 15, 15, 2, 15 };
                fotable.SetWidths(headerwidths13);
                fotable.Width = 100; break;
            case "14":
                int[] headerwidths14 = { 30, 9 };
                fotable.SetWidths(headerwidths14);
                fotable.Width = 100;
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
            case "29":
                int[] headerwidths20 = { 30, 9 };
                fotable.SetWidths(headerwidths20);
                fotable.Width = 39;
                break;

        }
    }

    public Boolean checkTrue(DataSet foDataset, int fiRowCount, String fsField)
    {
        Boolean lblReturn = false;
        if (foDataset.Tables[0].Rows.Count > 0)
        {
            if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString().ToUpper() == "TRUE")
            {
                lblReturn = true;
            }
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
        return setFontsAll(7, 0, 0, new iTextSharp.text.Color(165, 165, 165));
        //return setFontsAll(9, 0, 0, new iTextSharp.text.Color(175, 175, 175));
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

    public iTextSharp.text.Font Font18Bold(String fsTest)
    {
        if (fsTest == "test")
        {
            return setFontsAll(7, 1, 0);
        }
        else
        {
            return setFontsAll(7, 1, 0);
        }
    }

    public iTextSharp.text.Font Font19Bold(String fsTest)
    {
        if (fsTest == "test")
        {
            return setFontsAll1(7, 1);
        }
        else
        {
            return setFontsAll1(7, 1);
        }
    }

    public iTextSharp.text.Font Font7Bold()
    {
        return setFontsAll(7, 1, 0);
    }

    public void checkTrue(DataSet foDataset, int fiRowCount, String fsField, Cell foCell, iTextSharp.text.Color foColor)
    {

        if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString() == "6")
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


        //loCell = new Cell();
        ////loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
        //loCell.Leading = 15f;//25f
        //loCell.HorizontalAlignment = 2;
        //loCell.BorderWidth = 0;
        //loCell.Add(loChunk);
        //fotable.AddCell(loCell);

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

    protected void ddlAllocationGroup_SelectedIndexChanged(object sender, EventArgs e)
    {
        fillGroupAllocationTitle();
    }

    private void SetComboValue()
    {
        if (ddlHousehold.SelectedValue != "0")
        {
            if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
            {
                lsFamiliesName = ddlAllocationGroup.SelectedItem.Text;
            }
            else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
            {
                lsFamiliesName = drpHouseHoldReportTitle.SelectedItem.Text;
            }
            else
            {
                lsFamiliesName = drpAllocationGroupTitle.SelectedItem.Text;
            }
        }
        if (txtAsofdate.Text != "")
            lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";
    }
}

//public class CustomPieSectionLabelGenerator : org.jfree.chart.labels.PieSectionLabelGenerator
//{
//    /* other stuff... */


//    public String generateSectionLabel(org.jfree.data.general.PieDataset dataset, java.lang.Comparable keys)
//    {
//        StringBuilder label = new StringBuilder();
        
//        if (dataset.getItemCount() > 0 && keys != null)
//        {
//            /* I want to display the key
//             * but I also want to disply the value as an integer.
//             */
//            label.Append(keys);
//            label.Append(" = ");
//            label.Append(dataset.getValue(keys).floatValue());
//            label.Append("%");
//        }

//        return label.ToString();
//    }

//    public java.text.AttributedString generateAttributedSectionLabel(org.jfree.data.general.PieDataset dataset, java.lang.Comparable keys)
//    {
//        // TODO Auto-generated method stub
//        return null;
//    }

//}
