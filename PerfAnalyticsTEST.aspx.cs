using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Linq;


using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using java.awt.image;
using org.jfree.chart.entity;
using org.jfree.chart.encoders;
using org.jfree.chart.labels;
using org.jfree.data.category;
using System.Globalization;
using org.jfree.util;
using org.jfree.chart.annotations;
using org.jfree.data;
using java.io;
using System.Drawing;
using System.Web.UI.DataVisualization.Charting;
using Microsoft.IdentityModel.Claims;
using System.Threading;

public partial class PerfAnalytics1 : System.Web.UI.Page
{
    GeneralMethods clsGM = new GeneralMethods();

    string ColorTIA1 = "#558ED5"; //Blue
    string ColorNetInvestedCap = "#77933C"; //Green
    string ColorTIA2 = "#558ED5"; //Blue
    string ColorInflationAdjInvCap = "#E46C0A";//Orange
    double extendedrangePerc = 1.05;
    double max1 = 0.0;
    double min1 = 0.0;
    int SelectedRptCnt = 0;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            BindHouseHolds();
            Bind_AssetClass();
        }
    }

    public void BindHouseHolds()
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

    public void Bind_AssetClass()
    {
        string sqlstr = "SP_S_ASSET_CLASS";
        clsGM.getListForBindListBox(lstAssetClass, sqlstr, "sas_name", "sas_assetclassId");

        lstAssetClass.Items.Insert(0, "All");
        lstAssetClass.Items[0].Value = "0";
        //lstAssetClass.SelectedIndex = 0;


        for (int i = 0; i < lstAssetClass.Items.Count; i++)
        {
            if (lstAssetClass.Items[i].Value.ToString() == "e2a78beb-d604-de11-a38c-001d09665e8f")//Domestic Equity
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "028b5efb-d604-de11-a38c-001d09665e8f")//Fixed Income
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "8413896b-4925-df11-b686-001d09665e8f")//Global Opportunistic
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "c2a2d71c-d704-de11-a38c-001d09665e8f")//Illiquid Real Assets
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "42b39247-d704-de11-a38c-001d09665e8f")//International Equity
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "0332530a-1ad3-df11-9789-0019b9e7ee05")//Liquid Real Assets
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "2287692a-d704-de11-a38c-001d09665e8f")//Low Volatility Hedged Strategies
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "02ffe912-d704-de11-a38c-001d09665e8f")//Private Equity
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "9776259d-0392-4de0-8a12-0399724abf8d") //Cash and Equivalents
                lstAssetClass.Items[i].Selected = true;

            if (lstAssetClass.Items[i].Value.ToString().ToUpper() == "C1B9A3B8-D578-E511-9418-005056A0567E")//Diversified growth
                lstAssetClass.Items[i].Selected = true;

            if (lstAssetClass.Items[i].Value.ToString().ToUpper() == "106A7EA0-7D76-E511-9418-005056A0567E")//Emerging markets
                lstAssetClass.Items[i].Selected = true;

            if (lstAssetClass.Items[i].Value.ToString().ToUpper() == "C0845309-7D1D-E411-8A68-0019B9E7EE05")//Global equity
                lstAssetClass.Items[i].Selected = true;
        }

        //lstAssetClass.Items[3].Selected = true;
        //lstAssetClass.Items[4].Selected = true;
        //lstAssetClass.Items[5].Selected = true;
        //lstAssetClass.Items[6].Selected = true;
        //lstAssetClass.Items[7].Selected = true;
        //lstAssetClass.Items[9].Selected = true;
        //lstAssetClass.Items[12].Selected = true;
        //lstAssetClass.Items[16].Selected = true;
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        //GeneratePDF();
        GeneratePDFNew();
    }

    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        //string sqlstr = string.Empty;

        //sqlstr = "[sp_s_Get_GroupName_Only] @HouseHoldNameTxt='" + ddlHousehold .SelectedItem.Text+ "'";
        //BindDropdown(ddlGroup, sqlstr, "groupid", "groupname");

        DB clsDB = new DB();

        //Bind Group Name 
        DataSet loDataset = clsDB.getDataSet("[SP_S_HOUSEHOLD_GROUPNAME] @HouseHoldName='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@TiaFlg =2");
        ddlGroup.Items.Clear();
        ddlGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][1])));
        }

        //Allocation Group 
        //DataSet loDataset1 = clsDB.getDataSet("[SP_S_Advent_Allocation_Group] @HouseholdName='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@Flag = 1");
        //ddlAllocationGrp.Items.Clear();
        //ddlAllocationGrp.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        //for (int liCounter = 0; liCounter < loDataset1.Tables[0].Rows.Count; liCounter++)
        //{
        //    ddlAllocationGrp.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset1.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset1.Tables[0].Rows[liCounter][1])));
        //}

        //TIA Group
        DataSet loDataset2 = clsDB.getDataSet("[SP_S_HOUSEHOLD_GROUPNAME] @HouseholdName='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@Flag = NULL,@TiaFlg = 1");
        ddlTIAGrp.Items.Clear();
        ddlTIAGrp.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        for (int liCounter = 0; liCounter < loDataset2.Tables[0].Rows.Count; liCounter++)
        {
            ddlTIAGrp.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset2.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset2.Tables[0].Rows[liCounter][1])));
        }

    }

    private void GeneratePDFNew()
    {





        string ReportOpFolder = string.Empty;
        string ContactFolderName = string.Empty;
        string ParentFolder = string.Empty;
        string TempFolderPath = string.Empty;
        try
        {


            if (!chkrpt1.Checked && !chkrpt3.Checked && !chkrpt4.Checked)
            {
                lblError.Text = "Please select atleast one report";
                return;
            }
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





            DB clsDB = new DB();
            int rpt1 = 0, rpt2 = 0, rpt3 = 0;
            clsCombinedReports objCombinedReports = new clsCombinedReports();
            string[] SourceFileName = new string[3];

            string Grp = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";
            DataSet dsGrpName = clsDB.getDataSet("SP_S_ReportRollupGroupAllocationName @RRGName =" + Grp + "");
            string Familyname = dsGrpName.Tables[0].Rows[0][0].ToString();
            objCombinedReports.lsFamiliesName = Familyname;

            for (int i = 0; i <= 3; i++)
            {
                if (chkrpt1.Checked && i == 0)
                {
                    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "";
                    objCombinedReports.ReportRollupGroupIdName = TIAGrp;
                    objCombinedReports.GreshamAdvisedFlag = "TIA";
                }
                else
                {
                    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "" + ddlGroup.SelectedValue.Replace("'", "''") + "";
                    objCombinedReports.ReportRollupGroupIdName = GrpName;
                    objCombinedReports.GreshamAdvisedFlag = "GA";
                }

                objCombinedReports.HouseHoldText = ddlHousehold.SelectedItem.Text.Replace("'", "''");
                objCombinedReports.AsOfDate = txtAsofdate.Text;
                string strAssetClass = lstAssetClass.SelectedValue == "0" ? "" + GetAllItemsTextFromListBox(lstAssetClass, false) + "" : "" + GetAllItemsTextFromListBox(lstAssetClass, true) + "";
                objCombinedReports.AssetClassCSV = strAssetClass;
                //  DataSet dsTableRpt3 = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt3Table1));
                objCombinedReports.TempFolderPath = TempFolderPath;


                if (chkrpt1.Checked && rpt1 == 0)
                { SourceFileName[i] = objCombinedReports.generatePerfAnalyticsRpt1(); rpt1 = 1; }
                else if (chkrpt3.Checked && rpt2 == 0)
                { SourceFileName[i] = objCombinedReports.generatePerfAnalyticsRpt3(); rpt2 = 1; }
                else if (chkrpt4.Checked && rpt3 == 0)
                { SourceFileName[i] = objCombinedReports.generatePerfAnalyticsRpt4(); rpt3 = 1; }
            }
            PDFMerge PDF = new PDFMerge();
            Random rand = new Random();
            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmssfff") + rand.Next().ToString();

            PDF.MergeFiles(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\Perf_Analytics_" + strGUID + ".pdf", SourceFileName);

            // objCombinedReports.IncDate = dsTableRpt3.Tables[0].Rows[0]["InceptionDate"].ToString();
            try
            {
                ////loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                Response.Write("<script>");

                Response.Write("window.open('ViewReport.aspx?Perf_Analytics_" + strGUID + ".pdf', 'mywindow')");
                Response.Write("</script>");
            }
            catch (Exception exc)
            {
                Response.Write(exc.Message);
            }
        }
        catch (Exception ex)
        {
        }
        finally
        {
            if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
            {
                //lg.AddinLogFile(Session["Filename"].ToString(), "FOLDER DELETE--> " + ReportOpFolder + "\\" + ParentFolder + "-----" + DateTime.Now);
                Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
            }

        }
    }


    private void GeneratePDF()
    {
        if (!chkrpt1.Checked && !chkrpt3.Checked && !chkrpt4.Checked)
        {
            lblError.Text = "Please select atleast one report";
            return;
        }
        //Create Document with margin and size
        Random rnd = new Random();
        string date2 = System.DateTime.Today.ToString();

        string strGUID = DateTime.Parse(date2).ToString("yyyyMMdd") + "_PERF_ANALYTICS_" + DateTime.Now.ToString("yyyy-MM-dd-HHmmssfff") + "_" + rnd;
        String fsFinalLocation = Server.MapPath("~/ExcelTemplate/TempFolder/" + strGUID + ".pdf");

        iTextSharp.text.Document pdoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate(), -23, -20, 43, 0);//10,10
        String ls = Server.MapPath("~/ExcelTemplate/TempFolder/ls_" + strGUID + ".pdf");

        PdfWriter writer = PdfWriter.GetInstance(pdoc, new FileStream(ls, FileMode.Create));

        pdoc.Open();

        DateTime asofDT = Convert.ToDateTime(txtAsofdate.Text);

        string AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);
        DB clsDB = new DB();
        string Grp = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";
        DataSet dsGrpName = clsDB.getDataSet("SP_S_ReportRollupGroupAllocationName @RRGName =" + Grp + "");
        string Familyname = dsGrpName.Tables[0].Rows[0][0].ToString();


        if (chkrpt1.Checked)
        {
            #region REPORT 1

            //Creating Table for Heading -- Family NAme 
            PdfPTable LoHeader = new PdfPTable(1);
            //Table -- two charts 
            PdfPTable LoCharts = new PdfPTable(3);
            //Table -- Table Report 
            PdfPTable LoTempTableRpt = new PdfPTable(3);
            //Table -- Table Report 
            PdfPTable LoTableRpt = new PdfPTable(5);
            //Table -- Footer
            PdfPTable LoFooter = new PdfPTable(1);

            int[] widthHeader = { 100 };
            LoHeader.SetWidths(widthHeader);

            int[] widthChart = { 50, 1, 50 };
            LoCharts.SetWidths(widthChart);

            int[] widthTableRpt = { 3, 40, 28, 28, 3 };
            LoTableRpt.SetWidths(widthTableRpt);

            int[] widthTempTableRpt = { 28, 43, 28 };
            LoTempTableRpt.SetWidths(widthTempTableRpt);

            int[] widthFooter = { 100 };
            LoFooter.SetWidths(widthFooter);

            //Defining Width 
            LoHeader.TotalWidth = 100f;
            LoCharts.TotalWidth = 100f;
            LoTableRpt.TotalWidth = 100f;
            LoTempTableRpt.TotalWidth = 100f;
            LoFooter.TotalWidth = 100f;

            LoCharts.WidthPercentage = 100f;
            LoTableRpt.WidthPercentage = 100f;
            LoTempTableRpt.WidthPercentage = 100f;
            LoFooter.WidthPercentage = 100f;

            //Cell - ReportName (Center) <<FAMILY NAME>>
            PdfPCell loFamilyName = new PdfPCell();
            //Cell - Heading (Row2) <<TOTAL INVESTMENT ASSETS>>
            PdfPCell loHeadingRow2 = new PdfPCell();
            //Cell - Heading (Row3) <<Are My Assets Keeping Pace with Inflation?>>
            PdfPCell loHeadingRow3 = new PdfPCell();
            //Cell - Date (Row4) <<Month dd,yyyy>>
            PdfPCell loHeadingRow4 = new PdfPCell();
            //Cell - (Row5) <<Chart>> --> Table --> Chart1 --> Gap --> Chart2
            PdfPCell loHeadingRow5 = new PdfPCell();
            //Cell - (Row6) <<Table Report>> --> Table 
            PdfPCell loHeadingRow6 = new PdfPCell();
            //Cell - (Row7) <<Table -- Footer>> --> Table
            PdfPCell loHeadingRow7 = new PdfPCell();

            Paragraph HeadingFamilyName = new Paragraph(Familyname, setFontsAll(14, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph HeadingRow2 = new Paragraph("TOTAL INVESTMENT ASSETS", setFontsAll(10, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph HeadingRow3 = new Paragraph("Are My Investment Assets Keeping Pace with Inflation?", setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
            Paragraph HeadingRow4 = new Paragraph(AsOfDate, setFontsAll(10, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            //Cell Styles
            loFamilyName.Border = 0;
            HeadingFamilyName.SetAlignment("center");

            loHeadingRow2.Border = 0;
            HeadingRow2.SetAlignment("center");

            loHeadingRow3.Border = 0;
            HeadingRow3.SetAlignment("center");

            loHeadingRow4.Border = 0;
            HeadingRow4.SetAlignment("center");


            loFamilyName.AddElement(HeadingFamilyName);
            loHeadingRow2.AddElement(HeadingRow2);
            loHeadingRow3.AddElement(HeadingRow3);
            loHeadingRow4.AddElement(HeadingRow4);

            loHeadingRow2.PaddingTop = -4f;
            loHeadingRow3.PaddingTop = -5f;
            loHeadingRow4.PaddingTop = -5f;


            LoHeader.AddCell(loFamilyName);
            LoHeader.AddCell(loHeadingRow2);
            LoHeader.AddCell(loHeadingRow3);
            LoHeader.AddCell(loHeadingRow4);

            /*** Chart -- Section Start ***/

            //Cell - Left Chart
            PdfPCell CellChart1 = new PdfPCell();
            //Cell - Space between two Charts
            PdfPCell CellSpace1 = new PdfPCell();
            //Cell -- Right Chart
            PdfPCell CellChart2 = new PdfPCell();

            iTextSharp.text.Image chartimg1 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");
            iTextSharp.text.Image chartimg2 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");

            string filename1 = getLineChartReport1();
            string filename2 = filename1.Replace("OP_", "A_OP_");
            // string filename2 = filename1;
            //Document document = new Document();

            //JFreeChart chart = TEMP();
            //BufferedImage bufferedImage = chart.createBufferedImage(500, 500);
            //iTextSharp.text.Image image1 = iTextSharp.text.Image.GetInstance(bufferedImage);

            Paragraph PR1Row1Cell1 = new Paragraph("Total Investment Assets vs. Net Invested Capital", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            PR1Row1Cell1.SetAlignment("center");
            PR1Row1Cell1.SpacingAfter = 2f;
            PR1Row1Cell1.Leading = 10f;

            Paragraph PR1Row1Cell3 = new Paragraph("Total Investment Assets vs. Inflation Adj. Net Invested Capital", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            PR1Row1Cell3.SetAlignment("center");
            PR1Row1Cell3.SpacingAfter = 2f;
            PR1Row1Cell3.Leading = 10f;


            chartimg1 = iTextSharp.text.Image.GetInstance(filename1);
            chartimg2 = iTextSharp.text.Image.GetInstance(filename2);
            //chartimg.ScalePercent(65);
            chartimg1.ScalePercent(25);
            chartimg2.ScalePercent(25);

            CellChart1.AddElement(PR1Row1Cell1);
            CellChart1.AddElement(chartimg1);
            CellChart2.AddElement(PR1Row1Cell3);
            CellChart2.AddElement(chartimg2);


            CellChart1.HorizontalAlignment = Element.ALIGN_LEFT;
            CellChart1.HorizontalAlignment = Element.ALIGN_LEFT;

            Paragraph PSpace = new Paragraph(" ", setFontsAll(14, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));

            CellSpace1.AddElement(PSpace);
            CellSpace1.Border = 0;
            CellChart1.BorderWidth = 1.5f;
            CellChart2.BorderWidth = 1.5f;

            LoCharts.AddCell(CellChart1);
            LoCharts.AddCell(CellSpace1);
            LoCharts.AddCell(CellChart2);

            LoCharts.HorizontalAlignment = Element.ALIGN_LEFT;
            loHeadingRow5.AddElement(LoCharts);
            loHeadingRow5.Border = 0;
            loHeadingRow5.HorizontalAlignment = Element.ALIGN_LEFT;
            loHeadingRow5.PaddingTop = 20f;

            LoHeader.AddCell(loHeadingRow5);

            /*** Chart -- Section end ***/

            /*** Table Report ***/

            string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";

            DataSet dsTableRpt = clsDB.getDataSet("SP_R_ASSET_VALUE_NEW_GA_BASEDATA @GroupName =" + TIAGrp + ",@PositionGAFlagTxt='TIA' ,@TrxnGAFlagTxt= 'TIA' ,@AsOfDate = '" + txtAsofdate.Text + "'");

            string YearToDateMnyRow1 = MoneyFormat(dsTableRpt.Tables[0].Rows[0]["YearToDateMny"].ToString());
            string YearToDateMnyRow2 = MoneyFormat(dsTableRpt.Tables[0].Rows[1]["YearToDateMny"].ToString());
            string YearToDateMnyRow3 = MoneyFormat(dsTableRpt.Tables[0].Rows[2]["YearToDateMny"].ToString());
            string YearToDateMnyRow4 = MoneyFormat(dsTableRpt.Tables[0].Rows[3]["YearToDateMny"].ToString());
            string YearToDateMnyRow5 = MoneyFormat(dsTableRpt.Tables[0].Rows[4]["YearToDateMny"].ToString());

            string InceptionMnyRow1 = MoneyFormat(dsTableRpt.Tables[0].Rows[0]["InceptionMny"].ToString());
            string InceptionMnyRow2 = MoneyFormat(dsTableRpt.Tables[0].Rows[1]["InceptionMny"].ToString());
            string InceptionMnyRow3 = MoneyFormat(dsTableRpt.Tables[0].Rows[2]["InceptionMny"].ToString());
            string InceptionMnyRow4 = MoneyFormat(dsTableRpt.Tables[0].Rows[3]["InceptionMny"].ToString());
            string InceptionMnyRow5 = MoneyFormat(dsTableRpt.Tables[0].Rows[4]["InceptionMny"].ToString());

            string InceptionDt = dsTableRpt.Tables[0].Rows[4]["InceptionDt"].ToString();

            PdfPCell CellTRGap = new PdfPCell();

            PdfPCell CellTRrow1 = new PdfPCell();
            PdfPCell CellTRrow2 = new PdfPCell();
            PdfPCell CellTRrow3 = new PdfPCell();

            PdfPCell CellTRrow4Col1 = new PdfPCell();
            PdfPCell CellTRrow4Col2 = new PdfPCell();
            PdfPCell CellTRrow4Col3 = new PdfPCell();

            PdfPCell CellTRrow5Col1 = new PdfPCell();
            PdfPCell CellTRrow5Col2 = new PdfPCell();
            PdfPCell CellTRrow5Col3 = new PdfPCell();

            PdfPCell CellTRrow6Col1 = new PdfPCell();
            PdfPCell CellTRrow6Col2 = new PdfPCell();
            PdfPCell CellTRrow6Col3 = new PdfPCell();

            PdfPCell CellTRrow7Col1 = new PdfPCell();
            PdfPCell CellTRrow7Col2 = new PdfPCell();
            PdfPCell CellTRrow7Col3 = new PdfPCell();

            PdfPCell CellTRrow8Col1 = new PdfPCell();
            PdfPCell CellTRrow8Col2 = new PdfPCell();
            PdfPCell CellTRrow8Col3 = new PdfPCell();

            PdfPCell CellTRrow9Col1 = new PdfPCell();
            PdfPCell CellTRrow9Col2 = new PdfPCell();
            PdfPCell CellTRrow9Col3 = new PdfPCell();

            Paragraph PTRGap = new Paragraph(" ", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PTRHeading = new Paragraph("Total Investment Assets", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow2 = new Paragraph("as of " + AsOfDate, setFontsAll(8, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow3 = new Paragraph("Since        ", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow4Col1 = new Paragraph(" ", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow4Col2 = new Paragraph("Year-to Date", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow4Col3 = new Paragraph("  " + InceptionDt, setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow5Col1 = new Paragraph("Beginning Value", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow5Col2 = new Paragraph(YearToDateMnyRow1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow5Col3 = new Paragraph(InceptionMnyRow1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PTRRow6Col1 = new Paragraph("Contributions/Withdrawals", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow6Col2 = new Paragraph(YearToDateMnyRow2, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow6Col3 = new Paragraph(InceptionMnyRow2, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PTRRow7Col1 = new Paragraph("Net Invested Capital", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow7Col2 = new Paragraph(YearToDateMnyRow3, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow7Col3 = new Paragraph(InceptionMnyRow3, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PTRRow8Col1 = new Paragraph("Increase/Decrease", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow8Col2 = new Paragraph(YearToDateMnyRow4, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow8Col3 = new Paragraph(InceptionMnyRow4, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PTRRow9Col1 = new Paragraph("Ending Value", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow9Col2 = new Paragraph(YearToDateMnyRow5, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PTRRow9Col3 = new Paragraph(InceptionMnyRow5, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            PTRHeading.Leading = 10;
            PTRRow2.Leading = 10;
            PTRRow3.Leading = 10;



            PTRRow4Col1.Leading = 10;
            PTRRow4Col2.Leading = 10;
            PTRRow4Col3.Leading = 10;

            PTRRow5Col1.Leading = 5;
            PTRRow5Col2.Leading = 5;
            PTRRow5Col3.Leading = 5;

            //PTRRow6Col1.Leading = 7;
            //PTRRow6Col2.Leading = 7;
            //PTRRow6Col3.Leading = 7;

            PTRRow7Col1.Leading = 5;
            PTRRow7Col2.Leading = 5;
            PTRRow7Col3.Leading = 5;

            PTRRow8Col1.Leading = 1;
            PTRRow8Col2.Leading = 1;
            PTRRow8Col3.Leading = 1;

            PTRRow9Col1.Leading = 7;
            PTRRow9Col2.Leading = 7;
            PTRRow9Col3.Leading = 7;


            //PTRRow6Col1.SpacingAfter = 3f;
            //PTRRow6Col2.SpacingAfter = 3f;
            //PTRRow6Col3.SpacingAfter = 3f;

            //PTRRow8Col1.SpacingAfter = 3f;
            //PTRRow8Col2.SpacingAfter = 3f;
            //PTRRow8Col3.SpacingAfter = 3f;

            //CellTRrow1.FixedHeight = 15f;
            //CellTRrow2.FixedHeight = 15f;

            CellTRrow1.Border = 0;
            CellTRrow2.Border = 0;
            CellTRrow3.Border = 0;

            //CellTRrow3.FixedHeight = 12f;
            //CellTRrow3.PaddingBottom = -5f;

            //CellTRrow1.PaddingTop = 5f;
            //CellTRrow2.PaddingTop = 5f;
            //CellTRrow3.PaddingTop = 5f;

            //CellTRrow7Col1.PaddingTop = -7f;
            //CellTRrow7Col2.PaddingTop = -7f;
            //CellTRrow7Col2.PaddingTop = -7f;

            CellTRrow1.PaddingBottom = -3f;
            CellTRrow2.PaddingBottom = -3f;
            CellTRrow3.PaddingBottom = -3f;


            CellTRrow4Col1.PaddingBottom = -15f;
            CellTRrow4Col2.PaddingBottom = -15f;
            CellTRrow4Col3.PaddingBottom = -15f;

            CellTRrow5Col1.PaddingTop = 5f;
            CellTRrow5Col2.PaddingTop = 5f;
            CellTRrow5Col3.PaddingTop = 5f;

            CellTRrow5Col1.PaddingBottom = -20f;
            CellTRrow5Col2.PaddingBottom = -20f;
            CellTRrow5Col3.PaddingBottom = -20f;

            CellTRrow6Col1.PaddingTop = -5f;
            CellTRrow6Col2.PaddingTop = -5;
            CellTRrow6Col3.PaddingTop = -5;

            CellTRrow6Col1.PaddingBottom = -25f;
            CellTRrow6Col2.PaddingBottom = -25f;
            CellTRrow6Col3.PaddingBottom = -25f;

            CellTRrow7Col1.PaddingBottom = -10f;
            CellTRrow7Col1.PaddingBottom = -10f;
            CellTRrow7Col1.PaddingBottom = -10f;


            CellTRrow7Col1.PaddingTop = 5f;
            CellTRrow7Col2.PaddingTop = 5f;
            CellTRrow7Col3.PaddingTop = 5f;

            //CellTRrow6Col1.PaddingTop = 2f;
            //CellTRrow6Col2.PaddingTop = 5f;
            //CellTRrow6Col3.PaddingTop = 5f;

            //CellTRrow7Col1.PaddingTop = 5f;
            //CellTRrow7Col2.PaddingTop = 5f;
            //CellTRrow7Col3.PaddingTop = 5f;

            CellTRrow8Col1.PaddingTop = 10f;
            CellTRrow8Col2.PaddingTop = 10f;
            CellTRrow8Col3.PaddingTop = 10f;

            CellTRrow9Col1.PaddingTop = 5f;
            CellTRrow9Col2.PaddingTop = 5f;
            CellTRrow9Col3.PaddingTop = 5f;

            //CellTRrow4Col1.FixedHeight = 12f;
            //CellTRrow4Col2.FixedHeight = 12f;
            //CellTRrow4Col3.FixedHeight = 12f;

            CellTRrow4Col1.Border = 0;
            CellTRrow4Col2.Border = 0;
            CellTRrow4Col3.Border = 0;

            CellTRrow5Col1.Border = 0;
            CellTRrow5Col2.Border = 0;
            CellTRrow5Col3.Border = 0;

            CellTRrow6Col1.Border = 0;
            CellTRrow6Col2.Border = 0;
            CellTRrow6Col3.Border = 0;

            CellTRrow7Col1.Border = 0;
            CellTRrow7Col2.Border = 0;
            CellTRrow7Col3.Border = 0;

            CellTRrow8Col1.Border = 0;
            CellTRrow8Col2.Border = 0;
            CellTRrow8Col3.Border = 0;

            CellTRrow9Col1.Border = 0;
            CellTRrow9Col2.Border = 0;
            CellTRrow9Col3.Border = 0;




            PTRHeading.SetAlignment("center");
            PTRRow2.SetAlignment("center");
            PTRRow3.SetAlignment("right");

            PTRRow4Col2.SetAlignment("center");
            PTRRow4Col3.SetAlignment("center");

            PTRRow5Col1.SetAlignment("left");
            PTRRow5Col2.SetAlignment("right");
            PTRRow5Col3.SetAlignment("right");

            PTRRow6Col1.SetAlignment("left");
            PTRRow6Col2.SetAlignment("right");
            PTRRow6Col3.SetAlignment("right");

            PTRRow7Col1.SetAlignment("left");
            PTRRow7Col2.SetAlignment("right");
            PTRRow7Col3.SetAlignment("right");

            PTRRow8Col1.SetAlignment("left");
            PTRRow8Col2.SetAlignment("right");
            PTRRow8Col3.SetAlignment("right");

            PTRRow9Col1.SetAlignment("left");
            PTRRow9Col2.SetAlignment("right");
            PTRRow9Col3.SetAlignment("right");

            CellTRGap.AddElement(PTRGap);

            CellTRrow1.AddElement(PTRHeading);

            //CellTRrow1.FixedHeight = 16f;
            //CellTRrow2.FixedHeight = 15f;
            //CellTRrow3.FixedHeight = 15f;
            CellTRrow2.AddElement(PTRRow2);
            CellTRrow3.AddElement(PTRRow3);

            CellTRrow3.PaddingRight = 30f;

            CellTRGap.Border = 0;
            //CellTRrow3.PaddingRight = 10f;
            //CellTRrow4Col3.PaddingRight = 10f;
            //CellTRrow5Col3.PaddingRight = 10f;
            //CellTRrow6Col3.PaddingRight = 10f;
            //CellTRrow7Col3.PaddingRight = 10f;
            //CellTRrow8Col3.PaddingRight = 10f;
            //CellTRrow9Col3.PaddingRight = 10f;

            CellTRrow4Col1.AddElement(PTRRow4Col1);
            CellTRrow4Col2.AddElement(PTRRow4Col2);
            CellTRrow4Col3.AddElement(PTRRow4Col3);

            CellTRrow5Col1.AddElement(PTRRow5Col1);
            CellTRrow5Col2.AddElement(PTRRow5Col2);
            CellTRrow5Col3.AddElement(PTRRow5Col3);

            CellTRrow6Col1.AddElement(PTRRow6Col1);
            CellTRrow6Col2.AddElement(PTRRow6Col2);
            CellTRrow6Col3.AddElement(PTRRow6Col3);

            CellTRrow7Col1.AddElement(PTRRow7Col1);
            CellTRrow7Col2.AddElement(PTRRow7Col2);
            CellTRrow7Col3.AddElement(PTRRow7Col3);

            CellTRrow8Col1.AddElement(PTRRow8Col1);
            CellTRrow8Col2.AddElement(PTRRow8Col2);
            CellTRrow8Col3.AddElement(PTRRow8Col3);

            CellTRrow9Col1.AddElement(PTRRow9Col1);
            CellTRrow9Col2.AddElement(PTRRow9Col2);
            CellTRrow9Col3.AddElement(PTRRow9Col3);

            CellTRrow1.Colspan = 5;
            CellTRrow2.Colspan = 5;
            CellTRrow3.Colspan = 5;

            //CellTRrow1.PaddingTop = -7f;
            //CellTRrow2.PaddingTop = -5f;
            //CellTRrow3.PaddingTop = -5f;
            //CellTRrow3.PaddingRight = 35f;

            CellTRrow7Col1.BorderWidthTop = 1f;
            CellTRrow7Col2.BorderWidthTop = 1f;
            CellTRrow7Col3.BorderWidthTop = 1f;

            CellTRrow9Col1.BorderWidthTop = 2f;
            CellTRrow9Col2.BorderWidthTop = 2f;
            CellTRrow9Col3.BorderWidthTop = 2f;

            //CellTRrow6Col1.PaddingBottom = -4f;
            //CellTRrow6Col2.PaddingBottom = -4f;
            //CellTRrow6Col3.PaddingBottom = -4f;

            //CellTRrow7Col1.PaddingTop = -8f;
            //CellTRrow7Col2.PaddingTop = -8f;
            //CellTRrow7Col3.PaddingTop = -8f;

            //CellTRrow8Col1.PaddingBottom = 4f;
            //CellTRrow8Col2.PaddingBottom = 4f;
            //CellTRrow8Col3.PaddingBottom = 4f;

            //CellTRrow9Col1.PaddingTop = -1f;
            //CellTRrow9Col2.PaddingTop = -1f;
            //CellTRrow9Col3.PaddingTop = -1f;

            LoTableRpt.AddCell(CellTRrow1);
            LoTableRpt.AddCell(CellTRrow2);
            LoTableRpt.AddCell(CellTRrow3);

            LoTableRpt.AddCell(CellTRGap);
            LoTableRpt.AddCell(CellTRrow4Col1);
            LoTableRpt.AddCell(CellTRrow4Col2);
            LoTableRpt.AddCell(CellTRrow4Col3);
            LoTableRpt.AddCell(CellTRGap);

            LoTableRpt.AddCell(CellTRGap);
            LoTableRpt.AddCell(CellTRrow5Col1);
            LoTableRpt.AddCell(CellTRrow5Col2);
            LoTableRpt.AddCell(CellTRrow5Col3);
            LoTableRpt.AddCell(CellTRGap);

            LoTableRpt.AddCell(CellTRGap);
            LoTableRpt.AddCell(CellTRrow6Col1);
            LoTableRpt.AddCell(CellTRrow6Col2);
            LoTableRpt.AddCell(CellTRrow6Col3);
            LoTableRpt.AddCell(CellTRGap);

            LoTableRpt.AddCell(CellTRGap);
            LoTableRpt.AddCell(CellTRrow7Col1);
            LoTableRpt.AddCell(CellTRrow7Col2);
            LoTableRpt.AddCell(CellTRrow7Col3);
            LoTableRpt.AddCell(CellTRGap);

            LoTableRpt.AddCell(CellTRGap);
            LoTableRpt.AddCell(CellTRrow8Col1);
            LoTableRpt.AddCell(CellTRrow8Col2);
            LoTableRpt.AddCell(CellTRrow8Col3);
            LoTableRpt.AddCell(CellTRGap);

            LoTableRpt.AddCell(CellTRGap);
            LoTableRpt.AddCell(CellTRrow9Col1);
            LoTableRpt.AddCell(CellTRrow9Col2);
            LoTableRpt.AddCell(CellTRrow9Col3);
            LoTableRpt.AddCell(CellTRGap);


            // LoTableRpt.DefaultCell.Border = iTextSharp.text.Rectangle.TABLE;

            // loHeadingRow6.AddElement(LoTableRpt);

            /** Temp **/
            PdfPCell CellTempTRrow1 = new PdfPCell();
            PdfPCell CellTempTRrow2 = new PdfPCell();
            PdfPCell CellTempTRrow3 = new PdfPCell();


            Paragraph PTemp1 = new Paragraph(" ", setFontsAll(11, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));


            CellTempTRrow1.AddElement(PTemp1);
            CellTempTRrow2.AddElement(LoTableRpt);
            CellTempTRrow3.AddElement(PTemp1);

            CellTempTRrow1.Border = 0;
            // CellTempTRrow2.Border = 0;
            CellTempTRrow3.Border = 0;
            CellTempTRrow2.BorderWidth = 1.5f;


            LoTempTableRpt.AddCell(CellTempTRrow1);
            LoTempTableRpt.AddCell(CellTempTRrow2);
            LoTempTableRpt.AddCell(CellTempTRrow3);

            loHeadingRow6.AddElement(LoTempTableRpt);

            loHeadingRow6.PaddingTop = 13f;
            loHeadingRow6.Border = 0;
            LoHeader.AddCell(loHeadingRow6);


            /*** Footer Start ***/
            PdfPCell CellFooterRow1 = new PdfPCell();
            PdfPCell CellFooterRow2 = new PdfPCell();
            PdfPCell CellFooterRow3 = new PdfPCell();
            PdfPCell CellFooterRow4 = new PdfPCell();
            PdfPCell CellFooterRow5 = new PdfPCell();
            PdfPCell CellFooterRow6 = new PdfPCell();
            PdfPCell CellFooterRow7 = new PdfPCell();

            Phrase PFooterRow1 = new Phrase();
            Phrase PFooterRow2 = new Phrase();
            Phrase PFooterRow3 = new Phrase();
            Phrase PFooterRow4 = new Phrase();
            Phrase PFooterRow5 = new Phrase();
            Phrase PFooterRow6 = new Phrase();
            Phrase PFooterRow7 = new Phrase();

            Chunk PFooterRow1P1 = new Chunk("Total Investment Assets:", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow1P2 = new Chunk(" Includes all investment assets, Gresham advised and non-Gresham advised. Figure depicts overall asset base after spending and taxes. Excludes residences, personal property and", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow2P1 = new Chunk("\"below the line\" assets.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow3P1 = new Chunk("Net Invested Capital:", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow3P2 = new Chunk("Absolute amount of funds invested, not adjusted for market value changes, less withdrawals for non-investment purposes (e.g. homes).", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow4P1 = new Chunk("Infl. Adj. Initial TIA (Inflation Adjusted Initial Total Investment Assets): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow4P2 = new Chunk("Depicts the value of your overall investable asset base if it remained intact (funds not withdrawn for non-investment needs) ", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow5P1 = new Chunk("and adjusted for inflation. ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow6P1 = new Chunk("Infl. Adj. Net Invested Capital (Inflation Adjusted Net Invested Capital): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow6P2 = new Chunk("When compared to TIA, gap depicts asset growth over/under the rate of inflation. ", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow7P1 = new Chunk("Increase/Decrease: ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PFooterRow7P2 = new Chunk("The absolute value change in your investment asset base from one point in time to another. Not solely related to performance.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));

            PFooterRow1.Add(PFooterRow1P1);
            PFooterRow1.Add(PFooterRow1P2);
            PFooterRow2.Add(PFooterRow2P1);
            PFooterRow3.Add(PFooterRow3P1);
            PFooterRow3.Add(PFooterRow3P2);
            PFooterRow4.Add(PFooterRow4P1);
            PFooterRow4.Add(PFooterRow4P2);
            PFooterRow5.Add(PFooterRow5P1);
            PFooterRow6.Add(PFooterRow6P1);
            PFooterRow6.Add(PFooterRow6P2);
            PFooterRow7.Add(PFooterRow7P1);
            PFooterRow7.Add(PFooterRow7P2);

            CellFooterRow1.AddElement(PFooterRow1);
            //CellFooterRow1.AddElement(PFooterRow1P2);
            CellFooterRow2.AddElement(PFooterRow2);
            CellFooterRow3.AddElement(PFooterRow3);
            //CellFooterRow3.AddElement(PFooterRow3P2);
            //CellFooterRow4.AddElement(PFooterRow4P1);
            //CellFooterRow4.AddElement(PFooterRow4);
            // CellFooterRow5.AddElement(PFooterRow5);
            CellFooterRow6.AddElement(PFooterRow6);
            //CellFooterRow6.AddElement(PFooterRow6P2);
            CellFooterRow7.AddElement(PFooterRow7);
            //CellFooterRow7.AddElement(PFooterRow7P2);

            CellFooterRow1.PaddingTop = 0;
            CellFooterRow2.PaddingTop = -8f;
            CellFooterRow3.PaddingTop = -8f;
            CellFooterRow4.PaddingTop = -8f;
            CellFooterRow5.PaddingTop = -8f;
            CellFooterRow6.PaddingTop = -8f;
            CellFooterRow7.PaddingTop = -8f;

            CellFooterRow1.Border = 0;
            CellFooterRow2.Border = 0;
            CellFooterRow3.Border = 0;
            CellFooterRow4.Border = 0;
            CellFooterRow5.Border = 0;
            CellFooterRow6.Border = 0;
            CellFooterRow7.Border = 0;

            LoFooter.AddCell(CellFooterRow1);
            LoFooter.AddCell(CellFooterRow2);
            LoFooter.AddCell(CellFooterRow3);
            // LoFooter.AddCell(CellFooterRow4);
            LoFooter.AddCell(CellFooterRow5);
            LoFooter.AddCell(CellFooterRow6);
            LoFooter.AddCell(CellFooterRow7);

            LoFooter.WidthPercentage = 100f;
            LoFooter.TotalWidth = 100f;
            LoFooter.TotalWidth = 700;
            LoFooter.WriteSelectedRows(0, 7, 55, 100, writer.DirectContent);

            //loHeadingRow7.AddElement(LoFooter);
            //LoHeader.AddCell(loHeadingRow7);

            pdoc.Add(LoHeader);

            SelectedRptCnt++;
            #endregion
        }
        if (chkrpt3.Checked)
        {
            #region REPORT 3
            /** REPORT-3 **/

            //Creating Table for Heading -- Family NAme 
            PdfPTable LoR3Header = new PdfPTable(1);
            //Creating Table for ROW1 -- Chart 
            PdfPTable LoR3Row1 = new PdfPTable(1);
            //Creating Table for ROW2 -- Chart 
            PdfPTable LoR3Row2 = new PdfPTable(1);
            //Creating Table for ROW3 -- Chart 
            PdfPTable LoR3Row3 = new PdfPTable(3);
            PdfPTable LoR3Row3Temp = new PdfPTable(7);
            PdfPTable LoR3LoFooter = new PdfPTable(1);


            int[] widthR3Header = { 100 };
            LoR3Header.SetWidths(widthR3Header);

            int[] widthR3Row1 = { 100 };
            LoR3Row1.SetWidths(widthR3Row1);

            int[] widthR3Row2 = { 100 };
            LoR3Row2.SetWidths(widthR3Row2);

            LoR3Row2.SpacingBefore = 8f; //Gap between two tables 
            LoR3Row3.SpacingBefore = 8f; //Gap between two tables 

            int[] widthR3Row3 = { 15, 70, 15 };
            LoR3Row3.SetWidths(widthR3Row3);

            int[] widthR3Row3Temp = { 1, 32, 10, 10, 10, 12, 20 };
            LoR3Row3Temp.SetWidths(widthR3Row3Temp);

            int[] widthR3Footer = { 100 };
            LoR3LoFooter.SetWidths(widthR3Footer);

            LoR3Row3Temp.TotalWidth = 100f;
            LoR3Row3Temp.WidthPercentage = 100f;

            Paragraph PR3FamilyName = new Paragraph(Familyname, setFontsAll(14f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3HeadingRow2 = new Paragraph("GRESHAM ADVISED ASSETS", setFontsAll(10f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3HeadingRow3 = new Paragraph("How Have My Gresham Advised Assets Performed?", setFontsAll(12f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
            Paragraph PR3HeadingRow4 = new Paragraph(AsOfDate, setFontsAll(10f, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            //Cell - ReportName (Center) <<FAMILY NAME>>
            PdfPCell loR3FamilyName = new PdfPCell();
            //Cell - Heading (Row2) <<GRESHAM ADVISED ASSETS>>
            PdfPCell loR3HeadingRow2 = new PdfPCell();
            //Cell - Heading (Row3) <<How Have My Gresham Advised Assets Performed?>>
            PdfPCell loR3HeadingRow3 = new PdfPCell();
            //Cell - Date (Row4) <<Month dd,yyyy>>
            PdfPCell loR3HeadingRow4 = new PdfPCell();

            PdfPCell CellR3Chart1 = new PdfPCell();
            PdfPCell CellR3Chart2 = new PdfPCell();
            PdfPCell CellR3Row3Temp1 = new PdfPCell();
            PdfPCell CellR3Row3Temp2 = new PdfPCell();
            PdfPCell CellR3Row3Temp3 = new PdfPCell();


            //Cell Styles
            loR3FamilyName.Border = 0;
            PR3FamilyName.SetAlignment("center");

            loR3HeadingRow2.Border = 0;
            PR3HeadingRow2.SetAlignment("center");

            loR3HeadingRow3.Border = 0;
            PR3HeadingRow3.SetAlignment("center");

            loR3HeadingRow4.Border = 0;
            PR3HeadingRow4.SetAlignment("center");


            loR3FamilyName.AddElement(PR3FamilyName);
            loR3HeadingRow2.AddElement(PR3HeadingRow2);
            loR3HeadingRow3.AddElement(PR3HeadingRow3);
            loR3HeadingRow4.AddElement(PR3HeadingRow4);

            loR3HeadingRow2.PaddingTop = -4f;
            loR3HeadingRow3.PaddingTop = -5f;
            loR3HeadingRow4.PaddingTop = -5f;

            loR3HeadingRow4.PaddingBottom = 20f;

            LoR3Header.AddCell(loR3FamilyName);
            LoR3Header.AddCell(loR3HeadingRow2);
            LoR3Header.AddCell(loR3HeadingRow3);
            LoR3Header.AddCell(loR3HeadingRow4);

            iTextSharp.text.Image chartimg3 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");
            string filename3 = getLineChartReport3();
            chartimg3 = iTextSharp.text.Image.GetInstance(filename3);

            iTextSharp.text.Image chartimg4 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");
            // string filename4 = getBarChartTAB3();
            string filename4 = getBarChartReport3();
            chartimg4 = iTextSharp.text.Image.GetInstance(filename4);

            Paragraph PR3Row1Cell1 = new Paragraph("Growth of My Gresham Advised Assets (GAA)", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            PR3Row1Cell1.SetAlignment("center");
            PR3Row1Cell1.SpacingAfter = 2f;
            PR3Row1Cell1.Leading = 10f;

            Paragraph PR3Row3Cell1 = new Paragraph("Annual Performance of Gresham Advised Assets (GAA)", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            PR3Row3Cell1.SetAlignment("center");
            PR3Row3Cell1.SpacingAfter = 2f;
            PR3Row3Cell1.Leading = 10f;

            chartimg3.ScalePercent(25);
            chartimg4.ScalePercent(25);

            CellR3Chart1.AddElement(PR3Row1Cell1);
            CellR3Chart1.AddElement(chartimg3);
            CellR3Chart1.HorizontalAlignment = Element.ALIGN_LEFT;
            CellR3Chart1.BorderWidth = 1.5f;


            CellR3Chart2.AddElement(PR3Row3Cell1);
            CellR3Chart2.AddElement(chartimg4);
            CellR3Chart2.HorizontalAlignment = Element.ALIGN_LEFT;
            CellR3Chart2.BorderWidth = 1.5f;

            // LoCharts.HorizontalAlignment = Element.ALIGN_LEFT;
            LoR3Row1.AddCell(CellR3Chart1);
            LoR3Row2.AddCell(CellR3Chart2);

            PdfPCell CellR3TRGapLeft = new PdfPCell();
            PdfPCell CellR3TRGapRight = new PdfPCell();
            PdfPCell CellR3TRGapLeft5 = new PdfPCell();
            PdfPCell CellR3TRGapRight5 = new PdfPCell();

            PdfPCell CellR3TRRow1 = new PdfPCell();

            PdfPCell CellR3TRRow2Col1 = new PdfPCell();
            PdfPCell CellR3TRRow2Col2 = new PdfPCell(); //since

            PdfPCell CellR3TRRow3Col1 = new PdfPCell();
            PdfPCell CellR3TRRow3Col2 = new PdfPCell();
            PdfPCell CellR3TRRow3Col3 = new PdfPCell();
            PdfPCell CellR3TRRow3Col4 = new PdfPCell();
            PdfPCell CellR3TRRow3Col5 = new PdfPCell();
            PdfPCell CellR3TRRow3Col6 = new PdfPCell();

            PdfPCell CellR3TRRow4Col1 = new PdfPCell();
            PdfPCell CellR3TRRow4Col2 = new PdfPCell();
            PdfPCell CellR3TRRow4Col3 = new PdfPCell();
            PdfPCell CellR3TRRow4Col4 = new PdfPCell();
            PdfPCell CellR3TRRow4Col5 = new PdfPCell();
            PdfPCell CellR3TRRow4Col6 = new PdfPCell();

            PdfPCell CellR3TRRow5Col1 = new PdfPCell();
            PdfPCell CellR3TRRow5Col2 = new PdfPCell();
            PdfPCell CellR3TRRow5Col3 = new PdfPCell();
            PdfPCell CellR3TRRow5Col4 = new PdfPCell();
            PdfPCell CellR3TRRow5Col5 = new PdfPCell();
            PdfPCell CellR3TRRow5Col6 = new PdfPCell();

            string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

            string IncDate = "";
            string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
            DataSet dsTableRpt3 = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA  @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate ='" + txtAsofdate.Text + "' , @AnnPerfFlg = 1 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

            if (dsTableRpt3.Tables[0].Rows.Count < 1)
            {
                lblError.Text = "No Record Found";
                return;
            }

            IncDate = Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["InceptionDate"]);

            if (string.IsNullOrEmpty(IncDate))
            {
                lblError.Text = "No Record Found";
                return;
            }

            DataSet dsTableCPI = clsDB.getDataSet("exec SP_R_CPI_PERFORMANCE  @AsOfDate ='" + txtAsofdate.Text + "'  ,@SinceInceptDT ='" + IncDate.ToString() + "',@BenchMarkID	= '75D35570-F8BB-E211-8A81-0019B9E7EE05'");
            string str1year = "";
            string str3year = "";
            string str5year = "";
            string strSinceInc = "";
            string FamName = "";
            string str1yearCPI = "";
            string str3yearCPI = "";
            string str5yearCPI = "";
            string strSinceIncCPI = "";

            if (dsTableRpt3.Tables[0].Rows.Count > 0)
            {
                if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["1 Year"])))
                {
                    double num1Year = Convert.ToDouble(dsTableRpt3.Tables[0].Rows[0]["1 Year"].ToString());
                    if (num1Year == 0)
                        str1year = "N/A";
                    else
                        str1year = num1Year.ToString("P1", CultureInfo.InvariantCulture);
                }
                else
                    str1year = "N/A";

                if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["3 Year"])))
                {
                    double num3Year = Convert.ToDouble(dsTableRpt3.Tables[0].Rows[0]["3 Year"].ToString());
                    if (num3Year == 0)
                        str3year = "N/A";
                    else
                        str3year = num3Year.ToString("P1", CultureInfo.InvariantCulture);
                }
                else
                    str3year = "N/A";


                if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["5 Year"])))
                {
                    double num5Year = Convert.ToDouble(dsTableRpt3.Tables[0].Rows[0]["5 Year"].ToString());
                    if (num5Year == 0)
                        str5year = "N/A";
                    else
                        str5year = num5Year.ToString("P1", CultureInfo.InvariantCulture);
                }
                else
                    str5year = "N/A";

                if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["Since Inception"])))
                {
                    double numSinceInc = Convert.ToDouble(dsTableRpt3.Tables[0].Rows[0]["Since Inception"].ToString());
                    if (numSinceInc == 0)
                        strSinceInc = "N/A";
                    else
                        strSinceInc = numSinceInc.ToString("P1", CultureInfo.InvariantCulture);
                }
                else
                    strSinceInc = "N/A";

                FamName = dsTableRpt3.Tables[0].Rows[0][0].ToString();
                IncDate = dsTableRpt3.Tables[0].Rows[0]["InceptionDate"].ToString();
            }

            if (dsTableCPI.Tables[0].Rows.Count > 0)
            {

                if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["1 Year"])))
                {
                    double num1Year = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["1 Year"].ToString());
                    if (num1Year == 0)
                        str1yearCPI = "N/A";
                    else
                        str1yearCPI = num1Year.ToString("N1", CultureInfo.InvariantCulture) + "%";
                }
                else
                    str1yearCPI = "N/A";

                if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["3 Year"])))
                {
                    double num3Year = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["3 Year"].ToString());
                    if (num3Year == 0)
                        str3yearCPI = "N/A";
                    else
                        str3yearCPI = num3Year.ToString("N1", CultureInfo.InvariantCulture) + "%";
                }
                else
                    str3yearCPI = "N/A";

                if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["5 Year"])))
                {
                    double num5Year = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["5 Year"].ToString());
                    if (num5Year == 0)
                        str5yearCPI = "N/A";
                    else
                        str5yearCPI = num5Year.ToString("N1", CultureInfo.InvariantCulture) + "%";
                }
                else
                    str5yearCPI = "N/A";

                if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["Since Inception"])))
                {
                    double numSinceInc = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["Since Inception"].ToString());
                    if (numSinceInc == 0)
                        strSinceIncCPI = "N/A";
                    else
                        strSinceIncCPI = numSinceInc.ToString("N1", CultureInfo.InvariantCulture) + "%";
                }
                else
                    strSinceIncCPI = "N/A";

            }


            Paragraph PR3TRGap = new Paragraph(" ", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR3TRHeading = new Paragraph("Annualized Trailing Performance", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR3TRRow2Col1 = new Paragraph(" ", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow2Col2 = new Paragraph("Since", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR3TRRow3Col1 = new Paragraph(" ", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow3Col2 = new Paragraph("1 Year", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow3Col3 = new Paragraph("3 Year", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow3Col4 = new Paragraph("5 Year", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow3Col5 = new Paragraph(IncDate, setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR3TRRow4Col1 = new Paragraph(FamName, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow4Col2 = new Paragraph(str1year, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow4Col3 = new Paragraph(str3year, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow4Col4 = new Paragraph(str5year, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow4Col5 = new Paragraph(strSinceInc, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));


            Paragraph PR3TRRow5Col1 = new Paragraph("Inflation (CPI)", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow5Col2 = new Paragraph(str1yearCPI, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow5Col3 = new Paragraph(str3yearCPI, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow5Col4 = new Paragraph(str5yearCPI, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR3TRRow5Col5 = new Paragraph(strSinceIncCPI, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));



            PR3TRHeading.SetAlignment("center");

            PR3TRRow2Col1.SetAlignment("center");
            PR3TRRow2Col2.SetAlignment("center");

            PR3TRRow3Col1.SetAlignment("right");
            PR3TRRow3Col2.SetAlignment("center");
            PR3TRRow3Col3.SetAlignment("center");
            PR3TRRow3Col4.SetAlignment("center");
            PR3TRRow3Col5.SetAlignment("center");

            PR3TRRow5Col1.SetAlignment("right");
            PR3TRRow5Col2.SetAlignment("center");
            PR3TRRow5Col3.SetAlignment("center");
            PR3TRRow5Col4.SetAlignment("center");
            PR3TRRow5Col5.SetAlignment("center");

            PR3TRRow4Col1.SetAlignment("right");
            PR3TRRow4Col2.SetAlignment("center");
            PR3TRRow4Col3.SetAlignment("center");
            PR3TRRow4Col4.SetAlignment("center");
            PR3TRRow4Col5.SetAlignment("center");


            CellR3TRGapLeft.Border = 0;
            CellR3TRGapRight.Border = 0;
            CellR3TRRow1.Border = 0;
            CellR3TRRow2Col1.Border = 0;
            CellR3TRRow2Col2.Border = 0;

            CellR3TRRow2Col1.PaddingBottom = -1f;
            CellR3TRRow2Col2.PaddingBottom = -1f;

            CellR3TRRow3Col1.PaddingTop = -5f;
            CellR3TRRow3Col2.PaddingTop = -5f;
            CellR3TRRow3Col3.PaddingTop = -5f;
            CellR3TRRow3Col4.PaddingTop = -5f;
            CellR3TRRow3Col5.PaddingTop = -5f;

            CellR3TRRow4Col1.PaddingTop = -10f;
            CellR3TRRow4Col2.PaddingTop = -10f;
            CellR3TRRow4Col3.PaddingTop = -10f;
            CellR3TRRow4Col4.PaddingTop = -10f;
            CellR3TRRow4Col5.PaddingTop = -10f;

            CellR3TRRow5Col1.PaddingTop = -10f;
            CellR3TRRow5Col2.PaddingTop = -10f;
            CellR3TRRow5Col3.PaddingTop = -10f;
            CellR3TRRow5Col4.PaddingTop = -10f;
            CellR3TRRow5Col5.PaddingTop = -10f;

            CellR3TRRow5Col1.PaddingBottom = -15f;
            CellR3TRRow5Col2.PaddingBottom = -15f;
            CellR3TRRow5Col3.PaddingBottom = -15f;
            CellR3TRRow5Col4.PaddingBottom = -15f;
            CellR3TRRow5Col5.PaddingBottom = -20f;



            CellR3TRRow3Col1.Border = 0;
            CellR3TRRow3Col2.Border = 0;
            CellR3TRRow3Col3.Border = 0;
            CellR3TRRow3Col4.Border = 0;
            CellR3TRRow3Col5.Border = 0;

            CellR3TRRow4Col1.Border = 0;
            CellR3TRRow4Col2.Border = 0;
            CellR3TRRow4Col3.Border = 0;
            CellR3TRRow4Col4.Border = 0;
            CellR3TRRow4Col5.Border = 0;

            CellR3TRRow5Col1.Border = 0;
            CellR3TRRow5Col2.Border = 0;
            CellR3TRRow5Col3.Border = 0;
            CellR3TRRow5Col4.Border = 0;
            CellR3TRRow5Col5.Border = 0;

            CellR3TRRow1.AddElement(PR3TRHeading);
            CellR3TRRow1.Colspan = 7;

            CellR3TRRow2Col1.AddElement(PR3TRRow2Col1);
            CellR3TRRow2Col1.Colspan = 4;

            CellR3TRRow2Col2.AddElement(PR3TRRow2Col2);

            CellR3TRGapLeft.AddElement(PR3TRGap);
            CellR3TRGapRight.AddElement(PR3TRGap);

            CellR3TRGapLeft5.AddElement(PR3TRGap);
            CellR3TRGapRight5.AddElement(PR3TRGap);


            CellR3TRRow3Col1.AddElement(PR3TRRow3Col1);
            CellR3TRRow3Col2.AddElement(PR3TRRow3Col2);
            CellR3TRRow3Col3.AddElement(PR3TRRow3Col3);
            CellR3TRRow3Col4.AddElement(PR3TRRow3Col4);
            CellR3TRRow3Col5.AddElement(PR3TRRow3Col5);
            //CellR3TRRow3Col1.AddElement(PR3TRGap);

            CellR3TRRow4Col1.AddElement(PR3TRRow4Col1);
            CellR3TRRow4Col2.AddElement(PR3TRRow4Col2);
            CellR3TRRow4Col3.AddElement(PR3TRRow4Col3);
            CellR3TRRow4Col4.AddElement(PR3TRRow4Col4);
            CellR3TRRow4Col5.AddElement(PR3TRRow4Col5);

            CellR3TRRow5Col1.AddElement(PR3TRRow5Col1);
            CellR3TRRow5Col2.AddElement(PR3TRRow5Col2);
            CellR3TRRow5Col3.AddElement(PR3TRRow5Col3);
            CellR3TRRow5Col4.AddElement(PR3TRRow5Col4);
            CellR3TRRow5Col5.AddElement(PR3TRRow5Col5);

            CellR3TRGapRight5.Border = 0;
            CellR3TRGapLeft5.Border = 0;

            // CellR3TRRow5Col5.BackgroundColor = iTextSharp.text.Color.CYAN;
            CellR3TRRow5Col5.FixedHeight = 10f;
            CellR3TRRow5Col4.FixedHeight = 10f;
            CellR3TRRow5Col3.FixedHeight = 10f;
            CellR3TRRow5Col2.FixedHeight = 10f;
            CellR3TRRow5Col1.FixedHeight = 10f;
            CellR3TRGapLeft5.FixedHeight = 10f;
            CellR3TRGapRight5.FixedHeight = 10f;

            CellR3TRRow2Col1.FixedHeight = 15f;
            CellR3TRRow2Col2.FixedHeight = 12f;


            LoR3Row3Temp.AddCell(CellR3TRRow1);

            LoR3Row3Temp.AddCell(CellR3TRGapLeft5);
            LoR3Row3Temp.AddCell(CellR3TRRow2Col1);
            LoR3Row3Temp.AddCell(CellR3TRRow2Col2);
            LoR3Row3Temp.AddCell(CellR3TRGapRight5);

            LoR3Row3Temp.AddCell(CellR3TRGapLeft);
            LoR3Row3Temp.AddCell(CellR3TRRow3Col1);
            LoR3Row3Temp.AddCell(CellR3TRRow3Col2);
            LoR3Row3Temp.AddCell(CellR3TRRow3Col3);
            LoR3Row3Temp.AddCell(CellR3TRRow3Col4);
            LoR3Row3Temp.AddCell(CellR3TRRow3Col5);
            LoR3Row3Temp.AddCell(CellR3TRGapRight);

            LoR3Row3Temp.AddCell(CellR3TRGapLeft);
            LoR3Row3Temp.AddCell(CellR3TRRow4Col1);
            LoR3Row3Temp.AddCell(CellR3TRRow4Col2);
            LoR3Row3Temp.AddCell(CellR3TRRow4Col3);
            LoR3Row3Temp.AddCell(CellR3TRRow4Col4);
            LoR3Row3Temp.AddCell(CellR3TRRow4Col5);
            LoR3Row3Temp.AddCell(CellR3TRGapRight);

            LoR3Row3Temp.AddCell(CellR3TRGapLeft5);
            LoR3Row3Temp.AddCell(CellR3TRRow5Col1);
            LoR3Row3Temp.AddCell(CellR3TRRow5Col2);
            LoR3Row3Temp.AddCell(CellR3TRRow5Col3);
            LoR3Row3Temp.AddCell(CellR3TRRow5Col4);
            LoR3Row3Temp.AddCell(CellR3TRRow5Col5);
            LoR3Row3Temp.AddCell(CellR3TRGapRight5);

            // LoR3Row3Temp.SpacingAfter = -2f;


            CellR3Row3Temp1.Border = 0;
            CellR3Row3Temp3.Border = 0;
            CellR3Row3Temp2.AddElement(LoR3Row3Temp);
            CellR3Row3Temp2.BorderWidth = 1.5f;

            LoR3Row3.AddCell(CellR3Row3Temp1);
            LoR3Row3.AddCell(CellR3Row3Temp2);
            LoR3Row3.AddCell(CellR3Row3Temp3);

            PdfPCell CellR3FooterRow1 = new PdfPCell();
            PdfPCell CellR3FooterRow2 = new PdfPCell();
            PdfPCell CellR3FooterRow3 = new PdfPCell();
            PdfPCell CellR3FooterRow4 = new PdfPCell();

            Phrase PR3FooterRow1 = new Phrase();
            Phrase PR3FooterRow2 = new Phrase();
            Phrase PR3FooterRow3 = new Phrase();
            Phrase PR3FooterRow4 = new Phrase();


            Chunk PR3FooterRow1P1 = new Chunk("Gresham Advised Assets (GAA):", setFontsAll(7.5f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR3FooterRow1P2 = new Chunk(" All Gresham advised investments except cash prior to 1/1/2014 (includes non-marketable strategies).", setFontsAll(7.5f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR3FooterRow2P1 = new Chunk("Net Invested Capital:", setFontsAll(7.5f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR3FooterRow2P2 = new Chunk(" Total dollars invested, less amounts withdrawn.", setFontsAll(7.5f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR3FooterRow3P1 = new Chunk("Performance:", setFontsAll(7.5f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR3FooterRow3P2 = new Chunk(" Performance is shown net of all manager fees but gross of Gresham's fee, which covers a wide range of services. ", setFontsAll(7.5f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR3FooterRow4P1 = new Chunk("Infl. Adj. Net Invested Capital (Inflation Adjusted Net Invested Capital): ", setFontsAll(7.5f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR3FooterRow4P2 = new Chunk("When compared to GA, gap depicts performance relative to the rate of inflation.", setFontsAll(7.5f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));


            PR3FooterRow1.Add(PR3FooterRow1P1);
            PR3FooterRow1.Add(PR3FooterRow1P2);
            PR3FooterRow2.Add(PR3FooterRow2P1);
            PR3FooterRow2.Add(PR3FooterRow2P2);
            PR3FooterRow3.Add(PR3FooterRow3P1);
            PR3FooterRow3.Add(PR3FooterRow3P2);
            PR3FooterRow4.Add(PR3FooterRow4P1);
            PR3FooterRow4.Add(PR3FooterRow4P2);


            CellR3FooterRow1.AddElement(PR3FooterRow1);
            CellR3FooterRow2.AddElement(PR3FooterRow2);
            CellR3FooterRow3.AddElement(PR3FooterRow3);
            CellR3FooterRow4.AddElement(PR3FooterRow4);

            CellR3FooterRow1.Border = 0;
            CellR3FooterRow2.Border = 0;
            CellR3FooterRow3.Border = 0;
            CellR3FooterRow4.Border = 0;


            CellR3FooterRow2.PaddingTop = -8f;
            CellR3FooterRow3.PaddingTop = -8f;
            CellR3FooterRow4.PaddingTop = -8f;

            LoR3LoFooter.AddCell(CellR3FooterRow1);
            LoR3LoFooter.AddCell(CellR3FooterRow2);
            // LoR3LoFooter.AddCell(CellR3FooterRow4);
            LoR3LoFooter.AddCell(CellR3FooterRow3);




            // LoR3LoFooter.WidthPercentage = 100f;
            // LoR3LoFooter.TotalWidth = 100f;
            LoR3LoFooter.TotalWidth = 700;

            LoR3LoFooter.WriteSelectedRows(0, 3, 55, -100, writer.DirectContent);

            if (SelectedRptCnt > 0)
                pdoc.NewPage();

            pdoc.Add(LoR3Header);
            pdoc.Add(LoR3Row1);
            pdoc.Add(LoR3Row2);
            pdoc.Add(LoR3Row3);
            pdoc.Add(LoR3LoFooter);
            SelectedRptCnt++;
            #endregion
        }

        if (chkrpt4.Checked)
        {
            #region REPORT 4

            //Creating Table for Heading -- Family NAme 
            PdfPTable LoR4Header = new PdfPTable(1);
            //Creating Table for ROW1 -- Chart 
            PdfPTable LoR4Row1 = new PdfPTable(3);
            //Creating Table for ROW1 -- Chart 
            PdfPTable LoR5Row1 = new PdfPTable(3);
            //Creating Table for ROW2 -- Tables 
            PdfPTable LoR4Row2 = new PdfPTable(5);

            PdfPTable LoR4Row2Short = new PdfPTable(3);

            //Creating Table for ROW3 -- Chart 
            PdfPTable LoR4Row3 = new PdfPTable(3);

            PdfPTable LoR4Row2Table1 = new PdfPTable(5);
            PdfPTable LoR4Row2Table2 = new PdfPTable(5);


            PdfPTable LoR4LoFooter = new PdfPTable(1);
            PdfPTable LoR4Row1Cell3Table1 = new PdfPTable(1);

            PdfPTable LoR4Legends = new PdfPTable(7);



            int[] widthR4Header = { 100 };
            LoR4Header.SetWidths(widthR4Header);

            //int[] widthR4Row1 = { 100 };
            //LoR3Row1.SetWidths(widthR3Row1);

            //int[] widthR4Row2 = { 100 };
            //LoR3Row2.SetWidths(widthR3Row2);

            LoR4Row2.SpacingBefore = 20f; //Gap between two tables 
            LoR4Row3.SpacingBefore = 8f; //Gap between two tables 

            int[] widthR4Row1 = { 49, 2, 49 };
            LoR4Row1.SetWidths(widthR4Row1);

            int[] widthR5Row1 = { 49, 2, 49 };
            LoR5Row1.SetWidths(widthR5Row1);


            int[] widthR4Row2 = { 5, 35, 20, 35, 5 };
            LoR4Row2.SetWidths(widthR4Row2);

            int[] widthR4Row2Short = { 30, 40, 30 };
            LoR4Row2Short.SetWidths(widthR4Row2Short);

            int[] widthR4Row3 = { 15, 70, 15 };
            LoR4Row3.SetWidths(widthR4Row3);

            int[] widthR4Row2Table1 = { 3, 40, 25, 25, 10 };
            LoR4Row2Table1.SetWidths(widthR4Row2Table1);

            int[] widthR4Row2Table2 = { 3, 40, 25, 25, 10 };
            LoR4Row2Table2.SetWidths(widthR4Row2Table2);

            int[] widthR4Footer = { 100 };
            LoR4LoFooter.SetWidths(widthR4Footer);

            int[] widthR4Row1Cell3Table1 = { 100 };
            LoR4Row1Cell3Table1.SetWidths(widthR4Row1Cell3Table1);

            int[] widthR4Legends = { 10, 2, 40, 5, 2, 40, 10 };
            LoR4Legends.SetWidths(widthR4Legends);

            LoR4Row2Table1.TotalWidth = 100f;
            LoR4Row2Table1.WidthPercentage = 100f;

            LoR4Row2Table2.TotalWidth = 100f;
            LoR4Row2Table2.WidthPercentage = 100f;


            // LoR4Row1Cell3Table1.TotalWidth = 100f;
            LoR4Row1Cell3Table1.WidthPercentage = 100f;

            string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

            string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

            // LoR4Legends.SpacingBefore = 30f;

            // LoR4Row2.TotalWidth = 100f;
            // LoR4Row2.WidthPercentage = 100f;

            Paragraph PR4FamilyName = new Paragraph(Familyname, setFontsAll(14f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4HeadingRow2 = new Paragraph("GRESHAM ADVISED ASSETS", setFontsAll(10f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4HeadingRow3 = new Paragraph("How Has Gresham Reduced My Portfolio's Risk?", setFontsAll(12f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
            Paragraph PR4HeadingRow4 = new Paragraph(AsOfDate, setFontsAll(10f, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            //Cell - ReportName (Center) <<FAMILY NAME>>
            PdfPCell loR4FamilyName = new PdfPCell();
            //Cell - Heading (Row2) <<GRESHAM ADVISED ASSETS>>
            PdfPCell loR4HeadingRow2 = new PdfPCell();
            //Cell - Heading (Row3) <<How Have My Gresham Advised Assets Performed?>>
            PdfPCell loR4HeadingRow3 = new PdfPCell();
            //Cell - Date (Row4) <<Month dd,yyyy>>
            PdfPCell loR4HeadingRow4 = new PdfPCell();

            PdfPCell CellR4Row1Cell1 = new PdfPCell();
            PdfPCell CellR4Row1Cell2 = new PdfPCell();
            PdfPCell CellR4Row1Cell3 = new PdfPCell();

            PdfPCell CellR5Row1Cell1 = new PdfPCell();

            PdfPCell CellR4Row2Temp1 = new PdfPCell();
            PdfPCell CellR4Row2Temp2 = new PdfPCell();
            PdfPCell CellR4Row2Temp3 = new PdfPCell();
            PdfPCell CellR4Row2Temp4 = new PdfPCell();
            PdfPCell CellR4Row2Temp5 = new PdfPCell();


            PdfPCell CellR4Row1Cell3Table1 = new PdfPCell();

            //PdfPCell CellR4Row1Cell3Table1Row1 = new PdfPCell(); //Portfolio Protection During Worst Market Months
            //PdfPCell CellR4Row1Cell3Table1Row2 = new PdfPCell(); //Return %
            //PdfPCell CellR4Row1Cell3Table1Row3Cell1 = new PdfPCell(); //Blank
            //PdfPCell CellR4Row1Cell3Table1Row3Cell2 = new PdfPCell(); //11/01/2007 - 02/28/2009 Worst Market Period
            //PdfPCell CellR4Row1Cell3Table1Row4Cell1 = new PdfPCell();
            //PdfPCell CellR4Row1Cell3Table1Row4Cell2 = new PdfPCell();
            //PdfPCell CellR4Row1Cell3Table1Row5Cell1 = new PdfPCell(); //Left Chart 
            //PdfPCell CellR4Row1Cell3Table1Row5Cell2 = new PdfPCell(); //Right Chart 
            //PdfPCell CellR4Row1Cell3Table1Row6 = new PdfPCell(); //Legends 

            string Mindate;
            iTextSharp.text.Image chartimg6 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");
            //  string filename6 = GetChart5Left(out Mindate);
            string filename6 = getColumnChartReport4(Familyname);
            chartimg6 = iTextSharp.text.Image.GetInstance(filename6);

            //Paragraph PR4Row1Cell3Table1Row1 = new Paragraph("Portfolio Protection During Worst Market Months", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4Row1Cell3Table1Row2 = new Paragraph("Return %", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4Row1Cell3Table1Row3Cell1 = new Paragraph(" ", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4Row1Cell3Table1Row3Cell2 = new Paragraph("Worst Market Period", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4Row1Cell3Table1Row4Cell1 = new Paragraph(" ", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4Row1Cell3Table1Row4Cell2 = new Paragraph(Mindate, setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            //PR4Row1Cell3Table1Row1.SetAlignment("center");
            //PR4Row1Cell3Table1Row2.SetAlignment("center");

            //PR4Row1Cell3Table1Row1.Leading = 5;
            //PR4Row1Cell3Table1Row2.Leading = 5;
            //PR4Row1Cell3Table1Row3Cell1.Leading = 5;
            //PR4Row1Cell3Table1Row3Cell2.Leading = 5;
            //PR4Row1Cell3Table1Row4Cell1.Leading = 5;
            //PR4Row1Cell3Table1Row4Cell2.Leading = 5;
            //CellR4Row1Cell3Table1Row6.PaddingTop = 30f;

            //CellR4Row1Cell3Table1Row1.FixedHeight = 15f;
            //CellR4Row1Cell3Table1Row2.FixedHeight = 15f;
            //CellR4Row1Cell3Table1Row3Cell1.FixedHeight = 15f;
            //CellR4Row1Cell3Table1Row3Cell2.FixedHeight = 15f;
            //CellR4Row1Cell3Table1Row4Cell1.FixedHeight = 15f;
            //CellR4Row1Cell3Table1Row4Cell2.FixedHeight = 15f;

            //CellR4Row1Cell3Table1Row1.AddElement(PR4Row1Cell3Table1Row1);
            //CellR4Row1Cell3Table1Row2.AddElement(PR4Row1Cell3Table1Row2);
            //CellR4Row1Cell3Table1Row3Cell1.AddElement(PR4Row1Cell3Table1Row3Cell1);
            //CellR4Row1Cell3Table1Row3Cell2.AddElement(PR4Row1Cell3Table1Row3Cell2);
            //CellR4Row1Cell3Table1Row4Cell1.AddElement(PR4Row1Cell3Table1Row4Cell1);
            //CellR4Row1Cell3Table1Row4Cell2.AddElement(PR4Row1Cell3Table1Row4Cell2);

            //iTextSharp.text.Image chartimg7 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");
            //string filename7 = GetChart5Right();
            //chartimg7 = iTextSharp.text.Image.GetInstance(filename7);

            chartimg6.ScalePercent(25, 25);
            // chartimg7.ScalePercent(75, 75);


            CellR4Row1Cell3Table1.AddElement(chartimg6);

            // CellR4Row1Cell3Table1Row5Cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            // CellR4Row1Cell3Table1Row5Cell1.BorderWidth = 1.5f;

            //PdfPCell CellR4Legend1 = new PdfPCell();
            //PdfPCell CellR4Legend2 = new PdfPCell();
            //PdfPCell CellR4Legend3 = new PdfPCell();
            //PdfPCell CellR4Legend4 = new PdfPCell();
            //PdfPCell CellR4Legend5 = new PdfPCell();
            //PdfPCell CellR4Legend6 = new PdfPCell();
            //PdfPCell CellR4Legend7 = new PdfPCell();

            //Paragraph PR4Blank = new Paragraph("", setFontsAll(8f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4Legend3 = new Paragraph(Familyname, setFontsAll(8f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4Legend6 = new Paragraph("MSCI AC World", setFontsAll(8f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            //PR4Blank.Leading = 2; PR4Legend3.Leading = 2;
            //PR4Legend6.Leading = 2;

            //  CellR4Legend1.AddElement(PR4Blank);
            //CellR4Legend2.AddElement(PR4Blank);
            //CellR4Legend3.AddElement(PR4Legend3);
            //CellR4Legend4.AddElement(PR4Blank);
            //CellR4Legend5.AddElement(PR4Blank);
            //CellR4Legend6.AddElement(PR4Legend6);
            //CellR4Legend7.AddElement(PR4Blank);

            //CellR4Legend1.Border = 0;
            //CellR4Legend2.Border = 0;
            //CellR4Legend3.Border = 0;
            //CellR4Legend4.Border = 0;
            //CellR4Legend5.Border = 0;
            //CellR4Legend6.Border = 0;
            //CellR4Legend7.Border = 0;

            //CellR4Legend2.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#558ED5"));
            //CellR4Legend5.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#003399"));

            //LoR4Legends.AddCell(CellR4Legend1);
            //LoR4Legends.AddCell(CellR4Legend2);
            //LoR4Legends.AddCell(CellR4Legend3);
            //LoR4Legends.AddCell(CellR4Legend4);
            //LoR4Legends.AddCell(CellR4Legend5);
            //LoR4Legends.AddCell(CellR4Legend6);
            //LoR4Legends.AddCell(CellR4Legend7);

            ////  CellR4Row1Cell3Table1Row6.PaddingTop = 10f;
            //CellR4Row1Cell3Table1Row6.AddElement(LoR4Legends);

            //CellR4Row1Cell3Table1Row5Cell2.AddElement(chartimg7);
            //  CellR4Row1Cell3Table1Row5Cell2.HorizontalAlignment = Element.ALIGN_LEFT;
            // CellR4Row1Cell3Table1Row5Cell2.BorderWidth = 1.5f;

            //CellR4Row1Cell3Table1Row1.AddElement(PR4Row1Cell3Table1Row1);
            //CellR4Row1Cell3Table1Row2.AddElement(PR4Row1Cell3Table1Row2);
            //CellR4Row1Cell3Table1Row3Cell1.AddElement(PR4Row1Cell3Table1Row3Cell1);
            //CellR4Row1Cell3Table1Row3Cell2.AddElement(PR4Row1Cell3Table1Row3Cell2);
            //CellR4Row1Cell3Table1Row4Cell1.AddElement(PR4Row1Cell3Table1Row4Cell1);
            //CellR4Row1Cell3Table1Row4Cell2.AddElement(PR4Row1Cell3Table1Row4Cell2);

            //CellR4Row1Cell3Table1Row1.Colspan = 2;
            //CellR4Row1Cell3Table1Row2.Colspan = 2;
            //CellR4Row1Cell3Table1Row6.Colspan = 2;

            CellR4Row1Cell3Table1.Border = 0;
            //CellR4Row1Cell3Table1Row1.Border = 0;
            //CellR4Row1Cell3Table1Row2.Border = 0;
            //CellR4Row1Cell3Table1Row3Cell1.Border = 0;
            //CellR4Row1Cell3Table1Row3Cell2.Border = 0;
            //CellR4Row1Cell3Table1Row4Cell1.Border = 0;
            //CellR4Row1Cell3Table1Row4Cell2.Border = 0;
            //CellR4Row1Cell3Table1Row5Cell1.Border = 0;
            //CellR4Row1Cell3Table1Row5Cell2.Border = 0;
            //CellR4Row1Cell3Table1Row6.Border = 0;

            // CellR4Row1Cell3Table1Row5Cell2.Width = 100f;

            // CellR4Row1Cell3Table1Row5Cell1.BorderWidthRight = 1.5f;

            // CellR4Row1Cell3Table1Row5Cell1.UseVariableBorders = 
            LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row2);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row3Cell1);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row3Cell2);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row4Cell1);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row4Cell2);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row5Cell1);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row5Cell2);
            //LoR4Row1Cell3Table1.AddCell(CellR4Row1Cell3Table1Row6);


            //Cell Styles
            loR4FamilyName.Border = 0;
            PR4FamilyName.SetAlignment("center");

            loR4HeadingRow2.Border = 0;
            PR4HeadingRow2.SetAlignment("center");

            loR4HeadingRow3.Border = 0;
            PR4HeadingRow3.SetAlignment("center");

            loR4HeadingRow4.Border = 0;
            PR4HeadingRow4.SetAlignment("center");

            // CellR4Row1Cell1.Border = 0;
            CellR4Row1Cell2.Border = 0;
            //  CellR4Row3Temp3.Border = 0;

            CellR4Row2Temp1.Border = 0;
            CellR4Row2Temp3.Border = 0;
            CellR4Row2Temp5.Border = 0;
            CellR4Row2Temp2.BorderWidth = 1.5f;
            CellR4Row2Temp4.BorderWidth = 1.5f;

            loR4FamilyName.AddElement(PR4FamilyName);
            loR4HeadingRow2.AddElement(PR4HeadingRow2);
            loR4HeadingRow3.AddElement(PR4HeadingRow3);
            loR4HeadingRow4.AddElement(PR4HeadingRow4);

            loR4HeadingRow2.PaddingTop = -4f;
            loR4HeadingRow3.PaddingTop = -5f;
            loR4HeadingRow4.PaddingTop = -5f;

            loR4HeadingRow4.PaddingBottom = 20f;

            LoR4Header.AddCell(loR4FamilyName);
            LoR4Header.AddCell(loR4HeadingRow2);
            LoR4Header.AddCell(loR4HeadingRow3);
            LoR4Header.AddCell(loR4HeadingRow4);

            iTextSharp.text.Image chartimg5 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");
            string filename5 = getShapeChartReport4("4");
            chartimg5 = iTextSharp.text.Image.GetInstance(filename5);

            string IncDate = "";
            strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
            DataSet dsInc = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA  @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate ='" + txtAsofdate.Text + "' , @AnnPerfFlg = 1 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

            IncDate = Convert.ToString(dsInc.Tables[0].Rows[0]["InceptionDate"]);

            DateTime dtInc = DateTime.Parse(IncDate, new CultureInfo("en-US"));
            DateTime dtfixed = DateTime.Parse("01/01/2011", new CultureInfo("en-US"));
            string strR4Row1Cell1Heading = "";
            if (dtInc < dtfixed)
                strR4Row1Cell1Heading = "Performance vs. Volatility (since 01/01/2011)";
            else
                strR4Row1Cell1Heading = "Performance vs. Volatility (since " + dtInc.ToString("MM/dd/yyyy") + ")";

            chartimg5.ScalePercent(25, 25);
            //chartimg5.ScalePercent(75, 75);
            //  chartimg6.ScalePercent(75, 75);
            Paragraph PR4Row1Cell1Heading = new Paragraph(strR4Row1Cell1Heading, setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            PR4Row1Cell1Heading.SetAlignment("center");
            PR4Row1Cell1Heading.SpacingAfter = 2f;
            // PR4Row1Cell1Heading.SpacingBefore = -10f;
            PR4Row1Cell1Heading.Leading = 8f;

            Paragraph PR4Row1Cell3Heading = new Paragraph("Portfolio Protection During Worst Market Months", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            PR4Row1Cell3Heading.SetAlignment("center");
            PR4Row1Cell3Heading.Leading = 8f;
            PR4Row1Cell3Heading.SpacingAfter = 2f;
            //  PR4Row1Cell3Heading.SpacingBefore = -10f;

            CellR4Row1Cell1.AddElement(PR4Row1Cell1Heading);

            CellR4Row1Cell1.AddElement(chartimg5);
            CellR4Row1Cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            CellR4Row1Cell1.VerticalAlignment = Element.ALIGN_TOP;
            CellR4Row1Cell1.BorderWidth = 1.5f;


            CellR4Row1Cell3.AddElement(PR4Row1Cell3Heading);
            CellR4Row1Cell3.AddElement(LoR4Row1Cell3Table1);
            CellR4Row1Cell3.HorizontalAlignment = Element.ALIGN_LEFT;
            CellR4Row1Cell3.VerticalAlignment = Element.ALIGN_TOP;
            CellR4Row1Cell3.BorderWidth = 1.5f;

            LoR4Row1.AddCell(CellR4Row1Cell1);
            LoR4Row1.AddCell(CellR4Row1Cell2);
            LoR4Row1.AddCell(CellR4Row1Cell3);

            iTextSharp.text.Image chartimg8 = iTextSharp.text.Image.GetInstance(Server.MapPath("~") + @"\images\Gresham_Logo.png");
            string filename8 = getShapeChartReport4("5");
            chartimg8 = iTextSharp.text.Image.GetInstance(filename8);


            chartimg8.ScalePercent(25);
            // chartimg8.ScalePercent(75, 75);


            CellR5Row1Cell1.AddElement(chartimg8);
            CellR5Row1Cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            CellR5Row1Cell1.BorderWidth = 1.5f;


            LoR5Row1.AddCell(CellR5Row1Cell1);
            LoR5Row1.AddCell(CellR4Row1Cell2);
            LoR5Row1.AddCell(CellR4Row1Cell3);

            LoR4Row2Table1 = GetReport4Table(GrpName, strAssetClass, "ST");
            CellR4Row2Temp2.AddElement(LoR4Row2Table1);

            //LoR4Row2.AddCell(CellR4Row2Temp1);
            //LoR4Row2.AddCell(CellR4Row2Temp2);
            //LoR4Row2.AddCell(CellR4Row2Temp3);
            //LoR4Row2.AddCell(CellR4Row2Temp4);
            //LoR4Row2.AddCell(CellR4Row2Temp5);

            //Table 2

            PdfPCell CellR4TR2GapLeft = new PdfPCell();
            PdfPCell CellR4TR2GapRight = new PdfPCell();
            PdfPCell CellR4TR2GapLeft5 = new PdfPCell();
            PdfPCell CellR4TR2GapRight5 = new PdfPCell();

            PdfPCell CellR4TR2Row1 = new PdfPCell();

            PdfPCell CellR4TR2Row2Col1 = new PdfPCell();
            PdfPCell CellR4TR2Row2Col2 = new PdfPCell(); //since
            PdfPCell CellR4TR2Row2Col3 = new PdfPCell();

            PdfPCell CellR4TR2Row3Col1 = new PdfPCell();
            PdfPCell CellR4TR2Row3Col2 = new PdfPCell();
            PdfPCell CellR4TR2Row3Col3 = new PdfPCell();

            PdfPCell CellR4TR2Row4Col1 = new PdfPCell();
            PdfPCell CellR4TR2Row4Col2 = new PdfPCell();
            PdfPCell CellR4TR2Row4Col3 = new PdfPCell();

            PdfPCell CellR4TR2Row5Col1 = new PdfPCell();
            PdfPCell CellR4TR2Row5Col2 = new PdfPCell();
            PdfPCell CellR4TR2Row5Col3 = new PdfPCell();

            PdfPCell CellR4TR2Row6Col1 = new PdfPCell();
            PdfPCell CellR4TR2Row6Col2 = new PdfPCell();
            PdfPCell CellR4TR2Row6Col3 = new PdfPCell();

            DataSet dsTableRpt5 = clsDB.getDataSet("EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + "");

            //DataSet dsTableRpt4 = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA  @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate ='" + txtAsofdate.Text + "' , @AnnPerfFlg = 1 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");
            string strHeading = dsTableRpt5.Tables[0].Rows[9]["ReturnName"].ToString();
            string str1AR1 = "";
            string str2AR1 = "";
            string str3AR1 = "";
            string str4AR1 = "";

            string str1VOL1 = "";
            string str2VOL1 = "";
            string str3VOL1 = "";
            string str4VOL11 = "";

            //Annualized Return - GAA
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[0]["Return"])))
            {
                double num1AR = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[0]["Return"].ToString());
                if (num1AR == 0)
                    str1AR1 = "N/A";
                else
                    str1AR1 = num1AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str1AR1 = "N/A";

            //Volatility - GAA
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[1]["Return"])))
            {
                double num1VOL = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[1]["Return"].ToString());
                if (num1VOL == 0)
                    str1VOL1 = "N/A";
                else
                    str1VOL1 = num1VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str1VOL1 = "N/A";

            //Annualized Return -Marketable GAA
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[2]["Return"])))
            {
                double num2AR = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[2]["Return"].ToString());
                if (num2AR == 0)
                    str2AR1 = "N/A";
                else
                    str2AR1 = num2AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str2AR1 = "N/A";

            //Volatility - Marketable GAA
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[3]["Return"])))
            {
                double num2VOL = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[3]["Return"].ToString());
                if (num2VOL == 0)
                    str2VOL1 = "N/A";
                else
                    str2VOL1 = num2VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str2VOL1 = "N/A";

            //Annualized Return -Weighted Benchmark
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[4]["Return"])))
            {
                double num3AR = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[4]["Return"].ToString());
                if (num3AR == 0)
                    str3AR1 = "N/A";
                else
                    str3AR1 = num3AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str3AR1 = "N/A";

            //Volatility - Weighted Benchmark
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[5]["Return"])))
            {
                double num3VOL = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[5]["Return"].ToString());
                if (num3VOL == 0)
                    str3VOL1 = "N/A";
                else
                    str3VOL1 = num3VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str3VOL1 = "N/A";

            //Annualized Return -MSCI AC World
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[6]["Return"])))
            {
                double num4AR = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[6]["Return"].ToString());
                if (num4AR == 0)
                    str4AR1 = "N/A";
                else
                    str4AR1 = num4AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str4AR1 = "N/A";

            //Volatility -MSCI AC World
            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt5.Tables[0].Rows[7]["Return"])))
            {
                double num4VOL = Convert.ToDouble(dsTableRpt5.Tables[0].Rows[7]["Return"].ToString());
                if (num4VOL == 0)
                    str4VOL11 = "N/A";
                else
                    str4VOL11 = num4VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
            }
            else
                str4VOL11 = "N/A";

            Paragraph PR4TR2Gap = new Paragraph(" ", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR4TR2Heading = new Paragraph(strHeading, setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR4TR2Row2Col1 = new Paragraph(" ", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row2Col2 = new Paragraph("Annualized  Return", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row2Col3 = new Paragraph("Volatility", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR4TR2Row3Col1 = new Paragraph("GAA", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row3Col2 = new Paragraph(str1AR1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row3Col3 = new Paragraph(str1VOL1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR4TR2Row4Col1 = new Paragraph("Marketable GAA", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row4Col2 = new Paragraph(str2AR1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row4Col3 = new Paragraph(str2VOL1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            //Paragraph PR4TR2Row5Col1 = new Paragraph("Weighted Benchmark", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4TR2Row5Col2 = new Paragraph("6.0%", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Paragraph PR4TR2Row5Col3 = new Paragraph("6.8%", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            Paragraph PR4TR2Row6Col1 = new Paragraph("MSCI AC World", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row6Col2 = new Paragraph(str4AR1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph PR4TR2Row6Col3 = new Paragraph(str4VOL11, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            PR4TR2Heading.SetAlignment("center");

            PR4TR2Row2Col1.SetAlignment("left");
            PR4TR2Row2Col2.SetAlignment("right");
            PR4TR2Row2Col3.SetAlignment("right");

            PR4TR2Row3Col1.SetAlignment("left");
            PR4TR2Row3Col2.SetAlignment("right");
            PR4TR2Row3Col3.SetAlignment("right");

            PR4TR2Row4Col1.SetAlignment("left");
            PR4TR2Row4Col2.SetAlignment("right");
            PR4TR2Row4Col3.SetAlignment("right");

            //PR4TR2Row5Col1.SetAlignment("left");
            //PR4TR2Row5Col2.SetAlignment("right");
            //PR4TR2Row5Col3.SetAlignment("right");

            PR4TR2Row6Col1.SetAlignment("left");
            PR4TR2Row6Col2.SetAlignment("right");
            PR4TR2Row6Col3.SetAlignment("right");

            CellR4TR2Row1.Border = 0;
            CellR4TR2GapLeft.Border = 0;
            CellR4TR2GapLeft5.Border = 0;
            CellR4TR2GapRight.Border = 0;
            CellR4TR2GapRight5.Border = 0;
            CellR4TR2Row1.Border = 0;

            CellR4TR2Row2Col1.Border = 0;
            CellR4TR2Row2Col2.Border = 0;
            CellR4TR2Row2Col3.Border = 0;

            CellR4TR2Row3Col1.Border = 0;
            CellR4TR2Row3Col2.Border = 0;
            CellR4TR2Row3Col3.Border = 0;

            CellR4TR2Row4Col1.Border = 0;
            CellR4TR2Row4Col2.Border = 0;
            CellR4TR2Row4Col3.Border = 0;

            CellR4TR2Row5Col1.Border = 0;
            CellR4TR2Row5Col2.Border = 0;
            CellR4TR2Row5Col3.Border = 0;
            CellR4TR2Row6Col1.Border = 0;
            CellR4TR2Row6Col2.Border = 0;
            CellR4TR2Row6Col3.Border = 0;

            CellR4TR2Row6Col1.PaddingTop = 11f;
            CellR4TR2Row6Col2.PaddingTop = 11f;
            CellR4TR2Row6Col3.PaddingTop = 11f;

            CellR4TR2Row4Col1.PaddingTop = 11f;
            CellR4TR2Row4Col2.PaddingTop = 11f;
            CellR4TR2Row4Col3.PaddingTop = 11f;

            CellR4TR2GapLeft.AddElement(PR4TR2Gap);
            CellR4TR2Row1.PaddingTop = -5f;
            CellR4TR2Row1.AddElement(PR4TR2Heading);
            CellR4TR2Row1.Colspan = 5;

            // CellR4TR1Row2Col1.AddElement(PR4TR1Row2Col1);
            CellR4TR2Row2Col2.AddElement(PR4TR2Row2Col2);
            CellR4TR2Row2Col2.Colspan = 2;
            CellR4TR2Row2Col3.AddElement(PR4TR2Row2Col3);

            CellR4TR2Row3Col1.AddElement(PR4TR2Row3Col1);
            CellR4TR2Row3Col2.AddElement(PR4TR2Row3Col2);
            CellR4TR2Row3Col3.AddElement(PR4TR2Row3Col3);

            CellR4TR2Row4Col1.AddElement(PR4TR2Row4Col1);
            CellR4TR2Row4Col2.AddElement(PR4TR2Row4Col2);
            CellR4TR2Row4Col3.AddElement(PR4TR2Row4Col3);

            ////CellR4TR2Row5Col1.AddElement(PR4TR2Row5Col1);
            ////CellR4TR2Row5Col2.AddElement(PR4TR2Row5Col2);
            ////CellR4TR2Row5Col3.AddElement(PR4TR2Row5Col3);

            CellR4TR2Row6Col1.AddElement(PR4TR2Row6Col1);
            CellR4TR2Row6Col2.AddElement(PR4TR2Row6Col2);
            CellR4TR2Row6Col3.AddElement(PR4TR2Row6Col3);

            //.AddCell(CellR4TR1GapLeft);
            // LoR4Row2Table1.AddCell(CellR4TR1GapLeft5);
            //LoR4Row2Table1.AddCell(CellR4TR1GapRight);
            // LoR4Row2Table1.AddCell(CellR4TR1GapRight5);

            LoR4Row2Table2.AddCell(CellR4TR2Row1);

            LoR4Row2Table2.AddCell(CellR4TR2GapLeft);
            // LoR4Row2Table1.AddCell(CellR4TR1Row2Col1);
            LoR4Row2Table2.AddCell(CellR4TR2Row2Col2);
            LoR4Row2Table2.AddCell(CellR4TR2Row2Col3);
            LoR4Row2Table2.AddCell(CellR4TR2GapRight);

            LoR4Row2Table2.AddCell(CellR4TR2GapLeft);
            LoR4Row2Table2.AddCell(CellR4TR2Row3Col1);
            LoR4Row2Table2.AddCell(CellR4TR2Row3Col2);
            LoR4Row2Table2.AddCell(CellR4TR2Row3Col3);
            LoR4Row2Table2.AddCell(CellR4TR2GapRight);

            LoR4Row2Table2.AddCell(CellR4TR2GapLeft);
            LoR4Row2Table2.AddCell(CellR4TR2Row4Col1);
            LoR4Row2Table2.AddCell(CellR4TR2Row4Col2);
            LoR4Row2Table2.AddCell(CellR4TR2Row4Col3);
            LoR4Row2Table2.AddCell(CellR4TR2GapRight);

            //LoR4Row2Table2.AddCell(CellR4TR2GapLeft);
            //LoR4Row2Table2.AddCell(CellR4TR2Row5Col1);
            //LoR4Row2Table2.AddCell(CellR4TR2Row5Col2);
            //LoR4Row2Table2.AddCell(CellR4TR2Row5Col3);
            //LoR4Row2Table2.AddCell(CellR4TR2GapRight);

            LoR4Row2Table2.AddCell(CellR4TR2GapLeft);
            LoR4Row2Table2.AddCell(CellR4TR2Row6Col1);
            LoR4Row2Table2.AddCell(CellR4TR2Row6Col2);
            LoR4Row2Table2.AddCell(CellR4TR2Row6Col3);
            LoR4Row2Table2.AddCell(CellR4TR2GapRight);

            CellR4Row2Temp4.AddElement(LoR4Row2Table2);

            LoR4Row2.AddCell(CellR4Row2Temp1);
            LoR4Row2.AddCell(CellR4Row2Temp2);
            LoR4Row2.AddCell(CellR4Row2Temp3);
            LoR4Row2.AddCell(CellR4Row2Temp4);
            LoR4Row2.AddCell(CellR4Row2Temp5);


            PdfPCell CellR4FooterRow1 = new PdfPCell();
            PdfPCell CellR4FooterRow2 = new PdfPCell();
            PdfPCell CellR4FooterRow3 = new PdfPCell();
            PdfPCell CellR4FooterRow4 = new PdfPCell();

            Phrase PR4FooterRow1 = new Phrase();
            Phrase PR4FooterRow2 = new Phrase();
            Phrase PR4FooterRow3 = new Phrase();
            Phrase PR4FooterRow4 = new Phrase();


            Chunk PR4FooterRow1P1 = new Chunk("Gresham Advised Assets (GAA):", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR4FooterRow1P2 = new Chunk(" All Gresham advised investments except cash prior to 1/1/2014 (includes non-marketable strategies).", setFontsAll(7f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR4FooterRow2P1 = new Chunk("Gresham Advised Marketable Assets (Marketable GAA):", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR4FooterRow2P2 = new Chunk(" All Gresham advised investments except cash and non-marketable strategies.", setFontsAll(7f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR4FooterRow3P1 = new Chunk("Weighted Benchmark:", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR4FooterRow3P2 = new Chunk(" The average of the benchmark return for each asset class, weighted by that asset class' percentage of total marketable GAA each month. ", setFontsAll(7f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR4FooterRow4P1 = new Chunk("Performance:", setFontsAll(7f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
            Chunk PR4FooterRow4P2 = new Chunk(" Performance is shown net of all manager fees but gross of Gresham's fee, which covers a wide range of services.", setFontsAll(7f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));


            PR4FooterRow1.Add(PR4FooterRow1P1);
            PR4FooterRow1.Add(PR4FooterRow1P2);
            PR4FooterRow2.Add(PR4FooterRow2P1);
            PR4FooterRow2.Add(PR4FooterRow2P2);
            PR4FooterRow3.Add(PR4FooterRow3P1);
            PR4FooterRow3.Add(PR4FooterRow3P2);
            PR4FooterRow4.Add(PR4FooterRow4P1);
            PR4FooterRow4.Add(PR4FooterRow4P2);


            CellR4FooterRow1.AddElement(PR4FooterRow1);
            CellR4FooterRow2.AddElement(PR4FooterRow2);
            CellR4FooterRow3.AddElement(PR4FooterRow3);
            CellR4FooterRow4.AddElement(PR4FooterRow4);

            CellR4FooterRow1.Border = 0;
            CellR4FooterRow2.Border = 0;
            CellR4FooterRow3.Border = 0;
            CellR4FooterRow4.Border = 0;


            CellR4FooterRow2.PaddingTop = -8f;
            CellR4FooterRow3.PaddingTop = -8f;
            CellR4FooterRow4.PaddingTop = -8f;

            LoR4LoFooter.AddCell(CellR4FooterRow1);
            LoR4LoFooter.AddCell(CellR4FooterRow2);
            LoR4LoFooter.AddCell(CellR4FooterRow3);
            LoR4LoFooter.AddCell(CellR4FooterRow4);


            LoR4LoFooter.SpacingBefore = 25f;
            // LoR3LoFooter.WidthPercentage = 100f;
            // LoR3LoFooter.TotalWidth = 100f;
            LoR4LoFooter.TotalWidth = 700;

            LoR4LoFooter.WriteSelectedRows(0, 3, 55, -100, writer.DirectContent);



            PdfPCell CellR4ShortRow2Cell1 = new PdfPCell();
            PdfPCell CellR4ShortRow2Cell2 = new PdfPCell();
            PdfPCell CellR4ShortRow2Cell3 = new PdfPCell();

            LoR4Row2Table1 = GetReport4Table(GrpName, strAssetClass, "MT");
            CellR4ShortRow2Cell1.Border = 0;
            CellR4ShortRow2Cell3.Border = 0;
            CellR4ShortRow2Cell2.BorderWidth = 1.5f;
            CellR4ShortRow2Cell2.AddElement(LoR4Row2Table1);

            LoR4Row2Short.AddCell(CellR4ShortRow2Cell1);
            LoR4Row2Short.AddCell(CellR4ShortRow2Cell2);
            LoR4Row2Short.AddCell(CellR4ShortRow2Cell3);

            LoR4Row2Short.SpacingBefore = 20f;

            if (SelectedRptCnt > 0)
                pdoc.NewPage();

            //Get Inception date 
            GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";



            if (dtInc < dtfixed)
            {
                pdoc.Add(LoR4Header);
                pdoc.Add(LoR4Row1);
                pdoc.Add(LoR4Row2);
                pdoc.Add(LoR4Row3);
                pdoc.Add(LoR4LoFooter);
            }
            else
            {
                //   pdoc.NewPage();
                pdoc.Add(LoR4Header);
                pdoc.Add(LoR4Row1);
                pdoc.Add(LoR4Row2Short);
                pdoc.Add(LoR4Row3);
                pdoc.Add(LoR4LoFooter);
            }
            #endregion
        }


        pdoc.Close();

        FileInfo loFile = new FileInfo(ls);
        try
        {
            loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

            Response.Write("<script>");
            string lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + strGUID + ".pdf";
            Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
            Response.Write("</script>");

        }
        catch (Exception exc)
        {
            Response.Write(exc.Message);
        }

    }


    #region Get Charts

    #region Report 1

    private string getLineChartReport1()
    {
        Chart LineChartNEW1 = new Chart();
        LineChartNEW1.Height = 400;
        LineChartNEW1.Width = 800;
        LineChartNEW1.BorderlineDashStyle = ChartDashStyle.Solid;
        LineChartNEW1.Visible = false;


        LineChartNEW1.Titles.Add(new System.Web.UI.DataVisualization.Charting.Title("Total Investment Assets vs. Net Invested Capital"));
        LineChartNEW1.Titles[0].Visible = false;
        LineChartNEW1.Font.Name = "Frutiger55";
        LineChartNEW1.Font.Size = 9;
        LineChartNEW1.Font.Bold = true;

        LineChartNEW1.Series.Add(new Series());
        LineChartNEW1.Series.Add(new Series());

        LineChartNEW1.ChartAreas.Add(new ChartArea());
        LineChartNEW1.ChartAreas[0].BorderColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW1.ChartAreas[0].BackSecondaryColor = System.Drawing.Color.Transparent;
        LineChartNEW1.ChartAreas[0].BackColor = System.Drawing.Color.Transparent;
        LineChartNEW1.ChartAreas[0].ShadowColor = System.Drawing.Color.Transparent;

        LineChartNEW1.ChartAreas[0].AxisY.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW1.ChartAreas[0].AxisY.LineWidth = 2;
        LineChartNEW1.ChartAreas[0].AxisY.LabelStyle.Format = "{C0}";
        LineChartNEW1.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW1.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW1.ChartAreas[0].AxisX.LineWidth = 2;

        LineChartNEW1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        LineChartNEW1.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        LineChartNEW1.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        LineChartNEW1.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW1.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        LineChartNEW1.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        LineChartNEW1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW1.Legends.Add(new Legend());
        LineChartNEW1.Legends[0].LegendStyle = LegendStyle.Row;
        LineChartNEW1.Legends[0].Docking = Docking.Bottom;
        LineChartNEW1.Legends[0].Alignment = StringAlignment.Center;
        LineChartNEW1.Legends[0].TextWrapThreshold = 100;
        LineChartNEW1.Legends[0].AutoFitMinFontSize = 7;
        LineChartNEW1.Legends[0].IsTextAutoFit = false;
        LineChartNEW1.Legends[0].MaximumAutoSize = 100;

        /**** Chart 2 *****/

        Chart LineChartNEW2 = new Chart();
        LineChartNEW2.Height = 400;
        LineChartNEW2.Width = 800;
        LineChartNEW2.BorderlineDashStyle = ChartDashStyle.Solid;
        LineChartNEW2.Visible = false;


        LineChartNEW2.Titles.Add(new System.Web.UI.DataVisualization.Charting.Title("Total Investment Assets vs. Inflation Adj. Net Invested Capital"));
        LineChartNEW2.Titles[0].Visible = false;
        LineChartNEW2.Font.Name = "Frutiger55";
        LineChartNEW2.Font.Size = 9;
        LineChartNEW2.Font.Bold = true;

        LineChartNEW2.Series.Add(new Series());
        LineChartNEW2.Series.Add(new Series());

        LineChartNEW2.ChartAreas.Add(new ChartArea());
        LineChartNEW2.ChartAreas[0].BorderColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW2.ChartAreas[0].BackSecondaryColor = System.Drawing.Color.Transparent;
        LineChartNEW2.ChartAreas[0].BackColor = System.Drawing.Color.Transparent;
        LineChartNEW2.ChartAreas[0].ShadowColor = System.Drawing.Color.Transparent;

        LineChartNEW2.ChartAreas[0].AxisY.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW2.ChartAreas[0].AxisY.LineWidth = 2;
        LineChartNEW2.ChartAreas[0].AxisY.LabelStyle.Format = "{C0}";
        LineChartNEW2.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW2.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW2.ChartAreas[0].AxisX.LineWidth = 2;

        LineChartNEW2.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        LineChartNEW2.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        LineChartNEW2.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW2.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        LineChartNEW2.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW2.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        LineChartNEW2.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        LineChartNEW2.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW2.Legends.Add(new Legend());
        LineChartNEW2.Legends[0].LegendStyle = LegendStyle.Row;
        LineChartNEW2.Legends[0].Docking = Docking.Bottom;
        LineChartNEW2.Legends[0].Alignment = StringAlignment.Center;
        LineChartNEW2.Legends[0].TextWrapThreshold = 100;
        LineChartNEW2.Legends[0].AutoFitMinFontSize = 7;
        LineChartNEW2.Legends[0].IsTextAutoFit = false;
        LineChartNEW2.Legends[0].MaximumAutoSize = 100;



        Random rand = new Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmssfff") + rand.Next().ToString();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\OP_" + strGUID + ".xls";
        String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\P_OP_" + strGUID + ".xls";

        DB clsDB = new DB();
        string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        DataSet ds = clsDB.getDataSet("EXEC  SP_R_CLIENT_GOALS_NEW_GA_BASEDATA @GroupName = " + TIAGrp + ", @TrxnGAFlagTxt = 'TIA',@AsOfDate = '" + txtAsofdate.Text + "'");

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

        LineChartNEW1.DataSource = dtatble;
        LineChartNEW1.DataBind();

        LineChart2.DataSource = dtatble;
        LineChart2.DataBind();

        // Set series chart type
        LineChartNEW1.Series[0].ChartType = SeriesChartType.Line;
        LineChartNEW1.Series[1].ChartType = SeriesChartType.Line;

        LineChart2.Series[0].ChartType = SeriesChartType.Line;
        LineChart2.Series[1].ChartType = SeriesChartType.Line;


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

                LineChartNEW1.Series["Series1"].Points.AddXY(dtDate, s1val);
                LineChartNEW1.Series["Series2"].Points.AddXY(dtDate, s2val);

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
                LineChartNEW1.ChartAreas[0].AxisY.Maximum = maxx1;

            //Set Max value of Y-Axis -- Right Line chart 
            if (max1 > max4)
                maxx2 = RoundToMax(max1);

            if (max4 > max1)
                maxx2 = RoundToMax(max4);

            if (maxx2 != 0.0)
                LineChart2.ChartAreas[0].AxisY.Maximum = maxx2;

            if (maxx1 > 5000000 && maxx1 < 60000000)
                LineChartNEW1.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            if (maxx2 > 5000000 && maxx2 < 60000000)
                LineChart2.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            Double S1LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["value"].ToString());
            Double S2LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["NetInvestments"].ToString());
            Double S3LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"].ToString());

            LineChartNEW1.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);
            LineChartNEW1.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);

            LineChartNEW1.Series[0].BorderWidth = 10;
            LineChartNEW1.Series[1].BorderWidth = 10;


            LineChartNEW1.Series[0].Color = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW1.Series[1].Color = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);

            if (ds.Tables[0].Rows.Count < 12) //1years --MONTHLY
            {
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else if (ds.Tables[0].Rows.Count >= 12 && ds.Tables[0].Rows.Count < 36) //Quaterly
            {
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.Interval = 3;
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else
            {
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Years;
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1; //To start with december
                LineChartNEW1.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Days;
            }


            LineChartNEW1.Series[0].Name = "Total Investment Assets (TIA)";
            LineChartNEW1.Series[1].Name = "Net Invested Capital";


            //LineChartNEW1.Series[0].BorderColor = S;
            LineChartNEW1.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW1.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);


            LineChartNEW1.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChartNEW1.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S2LastValue.ToString("C0");

            LineChartNEW1.Series[0].SmartLabelStyle.Enabled = true;
            LineChartNEW1.Series[0].SmartLabelStyle.Enabled = true;

            if (S1LastValue > S2LastValue)
            {
                LineChartNEW1.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                LineChartNEW1.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }
            else
            {
                LineChartNEW1.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                LineChartNEW1.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            }


            LineChartNEW1.ChartAreas[0].AxisX.IsMarginVisible = false;
            LineChartNEW1.ChartAreas[0].AxisX.IsStartedFromZero = true;
            LineChart2.ChartAreas[0].AxisX.IsMarginVisible = false;
            LineChart2.ChartAreas[0].AxisX.IsStartedFromZero = true;

            //Remove Extra for label to display 
            LineChartNEW1.Series[0].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW1.Series[0].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChartNEW1.Series[0].SmartLabelStyle.CalloutLineWidth = 0;

            LineChartNEW1.Series[1].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW1.Series[1].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChartNEW1.Series[1].SmartLabelStyle.CalloutLineWidth = 0;



            LineChartNEW1.Series[0].IsVisibleInLegend = true;
            LineChartNEW1.Series[1].IsVisibleInLegend = true;


            LineChartNEW1.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChartNEW1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;


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

            //Remove Extra for label to display 
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
                LineChart2.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1; //To start with december
                LineChart2.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Days;
            }


            LineChart2.Series[0].IsVisibleInLegend = true;
            LineChart2.Series[1].IsVisibleInLegend = true;


            LineChart2.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
        }


        //LineChartNEW1.ChartAreas[0].Position.X = 5;
        //LineChartNEW1.ChartAreas[0].Position.Y = 8;
        //LineChartNEW1.ChartAreas[0].Position.Height = 82;
        //LineChartNEW1.ChartAreas[0].Position.Width = 97;

        //LineChart2.ChartAreas[0].Position.X = 5;
        //LineChart2.ChartAreas[0].Position.Y = 8;
        //LineChart2.ChartAreas[0].Position.Height = 82;
        //LineChart2.ChartAreas[0].Position.Width = 97;


        Random rnd = new Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\OP_" + RNum + ".bmp";
        string filename1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\A_OP_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(1300, 800);
        Bitmap bm1 = new Bitmap(1300, 800);

        bm.SetResolution(300, 300);
        bm1.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);
        System.Drawing.Graphics gGraphics1 = System.Drawing.Graphics.FromImage(bm1);


        LineChartNEW1.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 1300, 800));
        LineChart2.Paint(gGraphics1, new System.Drawing.Rectangle(0, 0, 1300, 800));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);
        bm1.Save(filename1, System.Drawing.Imaging.ImageFormat.Bmp);


        //  Chart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in Chart1.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;



    }

    #endregion

    #region Report 3

    private string getLineChartReport3()
    {

        Chart LineChartNEW = new Chart();
        LineChartNEW.Height = 195;
        LineChartNEW.Width = 1200;
        LineChartNEW.BorderlineDashStyle = ChartDashStyle.Solid;
        LineChartNEW.Visible = false;


        LineChartNEW.Titles.Add(new System.Web.UI.DataVisualization.Charting.Title("Growth of My Gresham Advised Assets (GAA)"));
        LineChartNEW.Titles[0].Visible = false;
        LineChartNEW.Font.Name = "Frutiger55";
        LineChartNEW.Font.Size = 9;
        LineChartNEW.Font.Bold = true;

        LineChartNEW.Series.Add(new Series());
        LineChartNEW.Series.Add(new Series());
        LineChartNEW.Series.Add(new Series());


        LineChartNEW.ChartAreas.Add(new ChartArea());
        LineChartNEW.ChartAreas[0].BorderColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW.ChartAreas[0].BackSecondaryColor = System.Drawing.Color.Transparent;
        LineChartNEW.ChartAreas[0].BackColor = System.Drawing.Color.Transparent;
        LineChartNEW.ChartAreas[0].ShadowColor = System.Drawing.Color.Transparent;

        LineChartNEW.ChartAreas[0].AxisY.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW.ChartAreas[0].AxisY.LineWidth = 2;
        LineChartNEW.ChartAreas[0].AxisY.LabelStyle.Format = "{C0}";
        LineChartNEW.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW.ChartAreas[0].AxisX.LineWidth = 2;

        LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        LineChartNEW.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        LineChartNEW.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        LineChartNEW.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        LineChartNEW.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        LineChartNEW.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW.Legends.Add(new Legend());
        LineChartNEW.Legends[0].LegendStyle = LegendStyle.Row;
        LineChartNEW.Legends[0].Docking = Docking.Bottom;
        LineChartNEW.Legends[0].Alignment = StringAlignment.Center;
        LineChartNEW.Legends[0].TextWrapThreshold = 100;
        LineChartNEW.Legends[0].AutoFitMinFontSize = 7;
        LineChartNEW.Legends[0].IsTextAutoFit = false;
        LineChartNEW.Legends[0].MaximumAutoSize = 100;

        Random rand = new Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmssfff") + rand.Next().ToString();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\2OP_" + strGUID + ".xls";
        String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\P_2OP_" + strGUID + ".xls";

        DB clsDB = new DB();
        string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
        DataSet ds = clsDB.getDataSet("exec SP_R_WEALTH_CHART_NEW_GA_BASEDATA @GroupName = " + GrpName + " , @PositionGAFlagTxt = 'GA', @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

        string Dt1;
        string sDay;
        string sMonth;
        string sYear;

        //chart1
        //  TimeSeries s1 = new TimeSeries("Gresham Advised Assets");
        //  TimeSeries s2 = new TimeSeries("Net Invested Capital");
        // TimeSeries s3 = new TimeSeries("Inflation Adj. Net Invested Capital");

        DataTable dtatble = ds.Tables[0];

        LineChartNEW.DataSource = dtatble;
        LineChartNEW.DataBind();

        // Set series chart type
        LineChartNEW.Series[0].ChartType = SeriesChartType.Line;
        LineChartNEW.Series[1].ChartType = SeriesChartType.Line;
        LineChartNEW.Series[2].ChartType = SeriesChartType.Line;

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

                LineChartNEW.Series["Series1"].Points.AddXY(dtDate, s1val);
                LineChartNEW.Series["Series2"].Points.AddXY(dtDate, s2val);
                LineChartNEW.Series["Series3"].Points.AddXY(dtDate, s3val);

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
                LineChartNEW.ChartAreas[0].AxisY.Maximum = maxx1;

            if (maxx1 > 5000000 && maxx1 < 50000000)
                LineChartNEW.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            Double S1LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Gresham Advised Assets"].ToString());
            Double S2LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Net Invested Capital"].ToString());
            Double S3LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"].ToString());

            LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);
            LineChartNEW.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);

            LineChartNEW.Series[0].BorderWidth = 10;
            LineChartNEW.Series[1].BorderWidth = 10;
            LineChartNEW.Series[2].BorderWidth = 10;

            LineChartNEW.Series[0].Color = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW.Series[1].Color = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);
            LineChartNEW.Series[2].Color = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);

            if (ds.Tables[0].Rows.Count < 12) //1years --MONTHLY
            {
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else if (ds.Tables[0].Rows.Count >= 12 && ds.Tables[0].Rows.Count < 36) //Quaterly
            {
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.Interval = 3;
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else
            {
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Years;
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1; //To start with december
                LineChartNEW.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Days;
            }



            LineChartNEW.ChartAreas[0].AxisX.IsMarginVisible = false;
            LineChartNEW.ChartAreas[0].AxisX.IsStartedFromZero = true;

            LineChartNEW.Series[0].Name = "Gresham Advised Assets";
            LineChartNEW.Series[1].Name = "Net Invested Capital";
            LineChartNEW.Series[2].Name = "Inflation Adj. Net Invested Capital";

            //LineChartNEW.Series[0].BorderColor = S;
            LineChartNEW.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);
            LineChartNEW.Series[2].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);

            LineChartNEW.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChartNEW.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S2LastValue.ToString("C0");
            LineChartNEW.Series[2].Points[ds.Tables[0].Rows.Count - 1].Label = S3LastValue.ToString("C0");

            int S1 = 0, S2 = 0, S3 = 0;
            double MaxPoint = 0.0, MinPoint = 0.0;
            double[] values = { S1LastValue, S2LastValue, S3LastValue };
            double maxval = values.Max();
            double minval = values.Min();

            if (S1LastValue == maxval)
            {
                LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                S1 = 1;
            }
            else if (S2LastValue == maxval)
            {
                LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                S2 = 1;
            }
            else
            {
                LineChartNEW.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                S3 = 1;
            }

            if (S1LastValue == minval)
            {
                LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                S1 = 1;
            }
            else if (S2LastValue == minval)
            {
                LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                S2 = 1;
            }
            else
            {
                LineChartNEW.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                S3 = 1;
            }

            if (S1 == 0)
            {
                double diff1 = maxval - S1LastValue;
                double diff2 = S1LastValue - minval;

                if (diff1 > diff2)
                    LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                else
                    LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }

            if (S2 == 0)
            {
                double diff1 = maxval - S2LastValue;
                double diff2 = S2LastValue - minval;

                if (diff1 > diff2)
                    LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                else
                    LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }

            if (S3 == 0)
            {
                double diff1 = maxval - S3LastValue;
                double diff2 = S3LastValue - minval;

                if (diff1 > diff2)
                    LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                else
                    LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }

            LineChartNEW.Series[2].SmartLabelStyle.IsMarkerOverlappingAllowed = true;
            LineChartNEW.Series[1].SmartLabelStyle.IsMarkerOverlappingAllowed = true;
            LineChartNEW.Series[0].SmartLabelStyle.IsMarkerOverlappingAllowed = true;


            //if (S1LastValue > S2LastValue && S1LastValue > S3LastValue)
            //{
            //    LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //    LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            //}

            //if (S2LastValue > S1LastValue && S2LastValue > S3LastValue)
            //{
            //    LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            //    LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //}

            //if (S3LastValue > S2LastValue && S3LastValue > S1LastValue)
            //{
            //    LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            //    LineChartNEW.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;

            //}



            //Remove Extra for label to display 
            LineChartNEW.Series[0].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW.Series[0].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChartNEW.Series[0].SmartLabelStyle.CalloutLineWidth = 0;

            LineChartNEW.Series[1].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW.Series[1].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChartNEW.Series[1].SmartLabelStyle.CalloutLineWidth = 0;

            LineChartNEW.Series[2].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW.Series[2].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChartNEW.Series[2].SmartLabelStyle.CalloutLineWidth = 0;


            LineChartNEW.Series[2].SmartLabelStyle.IsOverlappedHidden = false;
            LineChartNEW.Series[1].SmartLabelStyle.IsOverlappedHidden = false;
            LineChartNEW.Series[0].SmartLabelStyle.IsOverlappedHidden = false;

            LineChartNEW.Series[0].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;
            LineChartNEW.Series[1].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;
            LineChartNEW.Series[2].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;

            LineChartNEW.Series[0].IsVisibleInLegend = true;
            LineChartNEW.Series[1].IsVisibleInLegend = true;
            LineChartNEW.Series[2].IsVisibleInLegend = true;

            LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
        }



        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\A_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(2600, 600);

        bm.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);

        LineChartNEW.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 2600, 600));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        //  Chart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in Chart1.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;
    }

    private string getBarChartReport3()
    {


        Chart BarChart1 = new Chart();
        BarChart1.Height = 195;
        BarChart1.Width = 1200;
        BarChart1.BorderlineDashStyle = ChartDashStyle.Solid;
        BarChart1.Visible = false;


        BarChart1.Titles.Add(new System.Web.UI.DataVisualization.Charting.Title("Annual Performance of Gresham Advised Assets (GAA)"));
        BarChart1.Titles[0].Visible = false;
        BarChart1.Font.Name = "Frutiger55";
        BarChart1.Font.Size = 9;
        BarChart1.Font.Bold = true;

        BarChart1.Series.Add(new Series());
        BarChart1.Series[0].ChartType = SeriesChartType.Column;
        BarChart1.Series[0].IsXValueIndexed = false;
        BarChart1.Series[0].Color = System.Drawing.ColorTranslator.FromHtml("#2A6FB6");
        BarChart1.Series[0].BorderColor = System.Drawing.Color.Transparent;


        BarChart1.ChartAreas.Add(new ChartArea());
        BarChart1.ChartAreas[0].BorderColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        BarChart1.ChartAreas[0].BackSecondaryColor = System.Drawing.Color.Transparent;
        BarChart1.ChartAreas[0].BackColor = System.Drawing.Color.Transparent;
        BarChart1.ChartAreas[0].ShadowColor = System.Drawing.Color.Transparent;

        BarChart1.ChartAreas[0].AxisY.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        BarChart1.ChartAreas[0].AxisY.LineWidth = 2;
        BarChart1.ChartAreas[0].AxisY.LabelAutoFitMaxFontSize = 8;
        BarChart1.ChartAreas[0].AxisY.LabelStyle.Format = "{0.0}%";
        BarChart1.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        BarChart1.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        BarChart1.ChartAreas[0].AxisX.LineWidth = 2;
        BarChart1.ChartAreas[0].AxisX.Interval = 1;
        BarChart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        BarChart1.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        BarChart1.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        BarChart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;



        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\OP_11122.xls";

        // JFreeChart chart = ChartFactory.createBarChart(
        DB clsDB = new DB();

        string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

        DataSet ds = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA @GroupName = " + GrpName + ", @PositionGAFlagTxt = 'GA' , @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "', @AnnPerfFlg = 0 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

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

                string d = Convert.ToInt16(sDay) + "/" + Convert.ToInt16(sMonth) + "/" + Convert.ToInt32(sYear);
                Chart1.Series[0].Points.AddXY(sYear, s1val * 100);
            }
        }

        BarChart1.Series[0].IsValueShownAsLabel = true;
        BarChart1.Series[0].LabelFormat = "{0.0}%";
        BarChart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true; //disabled inner gridlines
        BarChart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false; //disabled inner gridlines
        BarChart1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 2;

        BarChart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false; //disabled inner gridlines
        BarChart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false; //disabled inner gridlines

        BarChart1.ChartAreas[0].AxisX.IsStartedFromZero = true;
        BarChart1.ChartAreas[0].AxisX.IsMarginVisible = true;

        //BarChart1.ChartAreas[0].AxisX.IsReversed = true;
        //clsDB.getConfiguration();

        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\A_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(2600, 600);

        bm.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);

        BarChart1.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 2600, 600));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        //  BarChart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in BarChart1.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;
    }

    #endregion

    #region Report 4
    private PdfPTable GetReport4Table(string GrpName, string strAssetClass, string ReportType)
    {




        DB clsDB = new DB();

        PdfPTable LoR4Row2Table1 = new PdfPTable(5);

        int[] widthR4Row2Table1 = { 3, 40, 25, 25, 10 };
        LoR4Row2Table1.SetWidths(widthR4Row2Table1);

        LoR4Row2Table1.TotalWidth = 100f;
        LoR4Row2Table1.WidthPercentage = 100f;
        string qry = "";
        if (ReportType == "LT") //long term 
            qry = "EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + "";

        else
            qry = "EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + ",@StartDate = '01-JAN-2011'";
        DataSet dsTableRpt4 = clsDB.getDataSet(qry);
        //DataSet dsTableRpt4 = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA  @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate ='" + txtAsofdate.Text + "' , @AnnPerfFlg = 1 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");
        string Heading = dsTableRpt4.Tables[0].Rows[9]["ReturnName"].ToString();

        string str1AR = "";
        string str2AR = "";
        string str3AR = "";
        string str4AR = "";

        string str1VOL = "";
        string str2VOL = "";
        string str3VOL = "";
        string str4VOL = "";

        //Annualized Return - GAA
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[0]["Return"])))
        {
            double num1AR = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[0]["Return"].ToString());
            if (num1AR == 0)
                str1AR = "N/A";
            else
                str1AR = num1AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str1AR = "N/A";

        //Volatility - GAA
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[1]["Return"])))
        {
            double num1VOL = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[1]["Return"].ToString());
            if (num1VOL == 0)
                str1VOL = "N/A";
            else
                str1VOL = num1VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str1VOL = "N/A";

        //Annualized Return -Marketable GAA
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[2]["Return"])))
        {
            double num2AR = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[2]["Return"].ToString());
            if (num2AR == 0)
                str2AR = "N/A";
            else
                str2AR = num2AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str2AR = "N/A";

        //Volatility - Marketable GAA
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[3]["Return"])))
        {
            double num2VOL = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[3]["Return"].ToString());
            if (num2VOL == 0)
                str2VOL = "N/A";
            else
                str2VOL = num2VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str2VOL = "N/A";

        //Annualized Return -Weighted Benchmark
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[4]["Return"])))
        {
            double num3AR = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[4]["Return"].ToString());
            if (num3AR == 0)
                str3AR = "N/A";
            else
                str3AR = num3AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str3AR = "N/A";

        //Volatility - Weighted Benchmark
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[5]["Return"])))
        {
            double num3VOL = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[5]["Return"].ToString());
            if (num3VOL == 0)
                str3VOL = "N/A";
            else
                str3VOL = num3VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str3VOL = "N/A";

        //Annualized Return -MSCI AC World
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[6]["Return"])))
        {
            double num4AR = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[6]["Return"].ToString());
            if (num4AR == 0)
                str4AR = "N/A";
            else
                str4AR = num4AR.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str4AR = "N/A";

        //Volatility -MSCI AC World
        if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt4.Tables[0].Rows[7]["Return"])))
        {
            double num4VOL = Convert.ToDouble(dsTableRpt4.Tables[0].Rows[7]["Return"].ToString());
            if (num4VOL == 0)
                str4VOL = "N/A";
            else
                str4VOL = num4VOL.ToString("F1", CultureInfo.InvariantCulture) + "%";
        }
        else
            str4VOL = "N/A";


        PdfPCell CellR4TR1GapLeft = new PdfPCell();
        PdfPCell CellR4TR1GapRight = new PdfPCell();
        PdfPCell CellR4TR1GapLeft5 = new PdfPCell();
        PdfPCell CellR4TR1GapRight5 = new PdfPCell();

        PdfPCell CellR4TR1Row1 = new PdfPCell();

        PdfPCell CellR4TR1Row2Col1 = new PdfPCell();
        PdfPCell CellR4TR1Row2Col2 = new PdfPCell(); //since
        PdfPCell CellR4TR1Row2Col3 = new PdfPCell();

        PdfPCell CellR4TR1Row3Col1 = new PdfPCell();
        PdfPCell CellR4TR1Row3Col2 = new PdfPCell();
        PdfPCell CellR4TR1Row3Col3 = new PdfPCell();

        PdfPCell CellR4TR1Row4Col1 = new PdfPCell();
        PdfPCell CellR4TR1Row4Col2 = new PdfPCell();
        PdfPCell CellR4TR1Row4Col3 = new PdfPCell();

        PdfPCell CellR4TR1Row5Col1 = new PdfPCell();
        PdfPCell CellR4TR1Row5Col2 = new PdfPCell();
        PdfPCell CellR4TR1Row5Col3 = new PdfPCell();

        PdfPCell CellR4TR1Row6Col1 = new PdfPCell();
        PdfPCell CellR4TR1Row6Col2 = new PdfPCell();
        PdfPCell CellR4TR1Row6Col3 = new PdfPCell();

        Paragraph PR4TR1Gap = new Paragraph(" ", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Heading;
        if (ReportType == "LT") //long term 
            PR4TR1Heading = new Paragraph(Heading, setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        else
            PR4TR1Heading = new Paragraph(Heading, setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        Paragraph PR4TR1Row2Col1 = new Paragraph(" ", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row2Col2 = new Paragraph("Annualized  Return", setFontsAll(8f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row2Col3 = new Paragraph("Volatility", setFontsAll(8f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        Paragraph PR4TR1Row3Col1 = new Paragraph("GAA", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row3Col2 = new Paragraph(str1AR, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row3Col3 = new Paragraph(str1VOL, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        Paragraph PR4TR1Row4Col1 = new Paragraph("Marketable GAA", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row4Col2 = new Paragraph(str2AR, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row4Col3 = new Paragraph(str2VOL, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        Paragraph PR4TR1Row5Col1 = new Paragraph("Weighted Benchmark", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row5Col2 = new Paragraph(str3AR, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row5Col3 = new Paragraph(str3VOL, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        Paragraph PR4TR1Row6Col1 = new Paragraph("MSCI AC World", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row6Col2 = new Paragraph(str4AR, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4TR1Row6Col3 = new Paragraph(str4VOL, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        PR4TR1Heading.SetAlignment("center");

        PR4TR1Row2Col1.SetAlignment("left");
        PR4TR1Row2Col2.SetAlignment("right");
        PR4TR1Row2Col3.SetAlignment("right");

        PR4TR1Row3Col1.SetAlignment("left");
        PR4TR1Row3Col2.SetAlignment("right");
        PR4TR1Row3Col3.SetAlignment("right");

        PR4TR1Row4Col1.SetAlignment("left");
        PR4TR1Row4Col2.SetAlignment("right");
        PR4TR1Row4Col3.SetAlignment("right");

        PR4TR1Row5Col1.SetAlignment("left");
        PR4TR1Row5Col2.SetAlignment("right");
        PR4TR1Row5Col3.SetAlignment("right");

        PR4TR1Row6Col1.SetAlignment("left");
        PR4TR1Row6Col2.SetAlignment("right");
        PR4TR1Row6Col3.SetAlignment("right");

        CellR4TR1Row1.Border = 0;
        CellR4TR1GapLeft.Border = 0;
        CellR4TR1GapLeft5.Border = 0;
        CellR4TR1GapRight.Border = 0;
        CellR4TR1GapRight5.Border = 0;
        CellR4TR1Row1.Border = 0;

        CellR4TR1Row2Col1.Border = 0;
        CellR4TR1Row2Col2.Border = 0;
        CellR4TR1Row2Col3.Border = 0;

        CellR4TR1Row3Col1.Border = 0;
        CellR4TR1Row3Col2.Border = 0;
        CellR4TR1Row3Col3.Border = 0;

        CellR4TR1Row4Col1.Border = 0;
        CellR4TR1Row4Col2.Border = 0;
        CellR4TR1Row4Col3.Border = 0;

        CellR4TR1Row5Col1.Border = 0;
        CellR4TR1Row5Col2.Border = 0;
        CellR4TR1Row5Col3.Border = 0;
        CellR4TR1Row6Col1.Border = 0;
        CellR4TR1Row6Col2.Border = 0;
        CellR4TR1Row6Col3.Border = 0;

        CellR4TR1GapLeft.AddElement(PR4TR1Gap);

        CellR4TR1Row1.AddElement(PR4TR1Heading);
        CellR4TR1Row1.PaddingTop = -5f;
        CellR4TR1Row1.Colspan = 5;

        // CellR4TR1Row2Col1.AddElement(PR4TR1Row2Col1);
        CellR4TR1Row2Col2.AddElement(PR4TR1Row2Col2);
        CellR4TR1Row2Col2.Colspan = 2;
        CellR4TR1Row2Col3.AddElement(PR4TR1Row2Col3);

        CellR4TR1Row3Col1.AddElement(PR4TR1Row3Col1);
        CellR4TR1Row3Col2.AddElement(PR4TR1Row3Col2);
        CellR4TR1Row3Col3.AddElement(PR4TR1Row3Col3);

        CellR4TR1Row4Col1.AddElement(PR4TR1Row4Col1);
        CellR4TR1Row4Col2.AddElement(PR4TR1Row4Col2);
        CellR4TR1Row4Col3.AddElement(PR4TR1Row4Col3);

        CellR4TR1Row5Col1.AddElement(PR4TR1Row5Col1);
        CellR4TR1Row5Col2.AddElement(PR4TR1Row5Col2);
        CellR4TR1Row5Col3.AddElement(PR4TR1Row5Col3);
        CellR4TR1Row5Col1.PaddingTop = -5f;
        CellR4TR1Row5Col2.PaddingTop = -5f;
        CellR4TR1Row5Col3.PaddingTop = -5f;

        CellR4TR1Row6Col1.AddElement(PR4TR1Row6Col1);
        CellR4TR1Row6Col2.AddElement(PR4TR1Row6Col2);
        CellR4TR1Row6Col3.AddElement(PR4TR1Row6Col3);

        //.AddCell(CellR4TR1GapLeft);
        // LoR4Row2Table1.AddCell(CellR4TR1GapLeft5);
        //LoR4Row2Table1.AddCell(CellR4TR1GapRight);
        // LoR4Row2Table1.AddCell(CellR4TR1GapRight5);

        LoR4Row2Table1.AddCell(CellR4TR1Row1);

        LoR4Row2Table1.AddCell(CellR4TR1GapLeft);
        // LoR4Row2Table1.AddCell(CellR4TR1Row2Col1);
        LoR4Row2Table1.AddCell(CellR4TR1Row2Col2);
        LoR4Row2Table1.AddCell(CellR4TR1Row2Col3);
        LoR4Row2Table1.AddCell(CellR4TR1GapRight);

        LoR4Row2Table1.AddCell(CellR4TR1GapLeft);
        LoR4Row2Table1.AddCell(CellR4TR1Row3Col1);
        LoR4Row2Table1.AddCell(CellR4TR1Row3Col2);
        LoR4Row2Table1.AddCell(CellR4TR1Row3Col3);
        LoR4Row2Table1.AddCell(CellR4TR1GapRight);

        LoR4Row2Table1.AddCell(CellR4TR1GapLeft);
        LoR4Row2Table1.AddCell(CellR4TR1Row4Col1);
        LoR4Row2Table1.AddCell(CellR4TR1Row4Col2);
        LoR4Row2Table1.AddCell(CellR4TR1Row4Col3);
        LoR4Row2Table1.AddCell(CellR4TR1GapRight);

        LoR4Row2Table1.AddCell(CellR4TR1GapLeft);
        LoR4Row2Table1.AddCell(CellR4TR1Row5Col1);
        LoR4Row2Table1.AddCell(CellR4TR1Row5Col2);
        LoR4Row2Table1.AddCell(CellR4TR1Row5Col3);
        LoR4Row2Table1.AddCell(CellR4TR1GapRight);

        LoR4Row2Table1.AddCell(CellR4TR1GapLeft);
        LoR4Row2Table1.AddCell(CellR4TR1Row6Col1);
        LoR4Row2Table1.AddCell(CellR4TR1Row6Col2);
        LoR4Row2Table1.AddCell(CellR4TR1Row6Col3);
        LoR4Row2Table1.AddCell(CellR4TR1GapRight);


        return LoR4Row2Table1;
    }

    private string getShapeChartReport4(string rptno)
    {
        Chart ShapeChartRptNEW4 = new Chart();
        ShapeChartRptNEW4.Height = 750;
        ShapeChartRptNEW4.Width = 750;
        ShapeChartRptNEW4.BorderlineDashStyle = ChartDashStyle.Solid;
        ShapeChartRptNEW4.Visible = false;


        ShapeChartRptNEW4.Titles.Add(new System.Web.UI.DataVisualization.Charting.Title("Performance vs. Volatility (since 01/01/2011)"));
        ShapeChartRptNEW4.Titles[0].Visible = false;
        ShapeChartRptNEW4.Font.Name = "Frutiger55";
        ShapeChartRptNEW4.Font.Size = 9;
        ShapeChartRptNEW4.Font.Bold = true;

        ShapeChartRptNEW4.Series.Add(new Series());
        ShapeChartRptNEW4.Series.Add(new Series());
        ShapeChartRptNEW4.Series.Add(new Series());
        ShapeChartRptNEW4.Series.Add(new Series());




        ShapeChartRptNEW4.ChartAreas.Add(new ChartArea());
        ShapeChartRptNEW4.ChartAreas[0].BorderColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].BackSecondaryColor = System.Drawing.Color.Transparent;
        ShapeChartRptNEW4.ChartAreas[0].BackColor = System.Drawing.Color.Transparent;
        ShapeChartRptNEW4.ChartAreas[0].ShadowColor = System.Drawing.Color.Transparent;

        ShapeChartRptNEW4.ChartAreas[0].AxisY.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        ShapeChartRptNEW4.ChartAreas[0].AxisY.LineWidth = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.Title = "Annualized Return %";

        //ShapeChartRptNEW4.ChartAreas[0].AxisY.TitleFont.Name = "Frutiger55";
        ShapeChartRptNEW4.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Frutiger55", 9, FontStyle.Bold);



        ShapeChartRptNEW4.ChartAreas[0].AxisY.LabelAutoFitMaxFontSize = 8;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.LabelStyle.Format = "{N0}%";
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        ShapeChartRptNEW4.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        ShapeChartRptNEW4.ChartAreas[0].AxisX.LineWidth = 2;

        ShapeChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.Format = "{N0}%";
        ShapeChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        ShapeChartRptNEW4.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;



        System.Random rand = new System.Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmssfff") + rand.Next().ToString();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\2OP_" + strGUID + ".xls";

        double Xmax = 0.0;
        double Ymax = 0.0;
        double axismax = 0.0;
        DB clsDB = new DB();
        string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

        DataSet ds = clsDB.getDataSet("EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + ",@StartDate = '01-JAN-2011'");

        DataTable dt = GetFormatedDatatable(ds);

        // Populate series data with random data
        Random random = new Random();
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

        ShapeChartRpt4.Series[0].Points.AddXY(s1X, s1Y);
        ShapeChartRpt4.Series[1].Points.AddXY(s2X, s2Y);
        ShapeChartRpt4.Series[2].Points.AddXY(s3X, s3Y);
        ShapeChartRpt4.Series[3].Points.AddXY(s4X, s4Y);

        ShapeChartRpt4.Series[0].Points[0].Label = dt.Rows[0]["Name"].ToString();
        ShapeChartRpt4.Series[1].Points[0].Label = dt.Rows[1]["Name"].ToString();
        ShapeChartRpt4.Series[2].Points[0].Label = dt.Rows[2]["Name"].ToString();
        ShapeChartRpt4.Series[3].Points[0].Label = dt.Rows[3]["Name"].ToString();

        // Set point chart type
        ShapeChartRpt4.Series[0].ChartType = SeriesChartType.Point;
        ShapeChartRpt4.Series[1].ChartType = SeriesChartType.Point;
        ShapeChartRpt4.Series[2].ChartType = SeriesChartType.Point;
        ShapeChartRpt4.Series[3].ChartType = SeriesChartType.Point;


        // Enable data points labels
        // ShapeChartRpt4.Series["Series1"].IsValueShownAsLabel = true;
        //  ShapeChartRpt4.Series["Series1"]["LabelStyle"] = "Center";

        // Set marker size
        ShapeChartRpt4.Series[0].MarkerSize = 10;
        ShapeChartRpt4.Series[1].MarkerSize = 10;
        ShapeChartRpt4.Series[2].MarkerSize = 10;
        ShapeChartRpt4.Series[3].MarkerSize = 10;

        ShapeChartRpt4.ChartAreas[0].Position.X = 0;
        ShapeChartRpt4.ChartAreas[0].Position.Y = 2;
        ShapeChartRpt4.ChartAreas[0].Position.Height = 95;
        ShapeChartRpt4.ChartAreas[0].Position.Width = 95;

        ShapeChartRpt4.ChartAreas[0].AxisX.Maximum = maxval;
        ShapeChartRpt4.ChartAreas[0].AxisY.Maximum = maxval;

        // ShapeChartRpt4.ChartAreas[0].InnerPlotPosition.Height = 80;
        // ShapeChartRpt4.ChartAreas[0].InnerPlotPosition.Width = 80;

        // Set marker shape
        ShapeChartRpt4.Series[0].MarkerStyle = MarkerStyle.Square; //Total GAA
        ShapeChartRpt4.Series[1].MarkerStyle = MarkerStyle.Circle; //Marketable GAA
        ShapeChartRpt4.Series[2].MarkerStyle = MarkerStyle.Diamond; //Strategic Benchmark
        ShapeChartRpt4.Series[3].MarkerStyle = MarkerStyle.Triangle; //MSCI

        // Set marker color 
        ShapeChartRpt4.Series[0].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#548ACF");
        ShapeChartRpt4.Series[1].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#8064A2");
        ShapeChartRpt4.Series[2].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");
        ShapeChartRpt4.Series[3].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#17375E");

        // Set marker border -  Strategic Benchmark
        ShapeChartRpt4.Series[2].MarkerBorderWidth = 1;
        ShapeChartRpt4.Series[2].MarkerBorderColor = System.Drawing.ColorTranslator.FromHtml("#215968");

        ShapeChartRpt4.ChartAreas[0].AxisX.Interval = 2;
        ShapeChartRpt4.ChartAreas[0].AxisY.Interval = 2;
        ShapeChartRpt4.ChartAreas[0].AxisX.Minimum = 0.0;

        ShapeChartRpt4.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Regular);
        ShapeChartRpt4.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Regular);

        ShapeChartRpt4.Titles[0].Font = new System.Drawing.Font("Frutiger65", 9, FontStyle.Bold);
        ShapeChartRpt4.Titles[0].Docking = Docking.Top;
        ShapeChartRpt4.Titles[0].DockingOffset = -2;

        // Enable 3D
        // ShapeChartRpt4.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;

        Random rnd = new Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\OP_" + RNum + ".bmp";

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

    private string getColumnChartReport4(string Fname)
    {

        Chart ColumnChartRptNEW4 = new Chart();
        ColumnChartRptNEW4.Height = 750;
        ColumnChartRptNEW4.Width = 750;
        ColumnChartRptNEW4.BorderlineDashStyle = ChartDashStyle.Solid;
        ColumnChartRptNEW4.Visible = false;


        ColumnChartRptNEW4.Titles.Add(new System.Web.UI.DataVisualization.Charting.Title("Portfolio Protection During Worst Market Months"));
        ColumnChartRptNEW4.Titles[0].Visible = false;
        ColumnChartRptNEW4.Titles[0].Alignment = ContentAlignment.TopCenter;
        ColumnChartRptNEW4.Font.Name = "Frutiger55";
        ColumnChartRptNEW4.Font.Size = 9;
        ColumnChartRptNEW4.Font.Bold = true;

        ColumnChartRptNEW4.Series.Add(new Series());
        ColumnChartRptNEW4.Series.Add(new Series());
        ColumnChartRptNEW4.Series.Add(new Series());

        ColumnChartRptNEW4.Series[0].Color = System.Drawing.ColorTranslator.FromHtml("#558ED5");
        ColumnChartRptNEW4.Series[1].Color = System.Drawing.ColorTranslator.FromHtml("#ffffff");
        ColumnChartRptNEW4.Series[2].Color = System.Drawing.ColorTranslator.FromHtml("#003399");



        ColumnChartRptNEW4.ChartAreas.Add(new ChartArea());
        ColumnChartRptNEW4.ChartAreas[0].BorderColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ColumnChartRptNEW4.ChartAreas[0].BackSecondaryColor = System.Drawing.Color.Transparent;
        ColumnChartRptNEW4.ChartAreas[0].BackColor = System.Drawing.Color.Transparent;
        ColumnChartRptNEW4.ChartAreas[0].ShadowColor = System.Drawing.Color.Transparent;

        ColumnChartRptNEW4.ChartAreas[0].AxisY.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        ColumnChartRptNEW4.ChartAreas[0].AxisY.LineWidth = 2;
        //ColumnChartRptNEW4.ChartAreas[0].AxisY.TitleFont.Name = "Frutiger55";
        ColumnChartRptNEW4.ChartAreas[0].AxisY.TitleFont = new System.Drawing.Font("Frutiger55", 9, FontStyle.Bold);


        ColumnChartRptNEW4.ChartAreas[0].AxisY.LabelStyle.Format = "{N0}%";
        ColumnChartRptNEW4.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 6, FontStyle.Regular);
        ColumnChartRptNEW4.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ColumnChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ColumnChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        ColumnChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        ColumnChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        ColumnChartRptNEW4.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        ColumnChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 9, FontStyle.Bold);
        ColumnChartRptNEW4.ChartAreas[0].AxisX.LineWidth = 2;
        ColumnChartRptNEW4.ChartAreas[0].AxisX.Title = "Return %";

        ColumnChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        ColumnChartRptNEW4.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ColumnChartRptNEW4.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        ColumnChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ColumnChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        ColumnChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        ColumnChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        ColumnChartRptNEW4.Legends.Add(new Legend());
        ColumnChartRptNEW4.Legends[0].LegendStyle = LegendStyle.Row;
        ColumnChartRptNEW4.Legends[0].Docking = Docking.Bottom;
        ColumnChartRptNEW4.Legends[0].TitleFont = new System.Drawing.Font("Frutiger55", 8, FontStyle.Regular);
        ColumnChartRptNEW4.Legends[0].TextWrapThreshold = 100;
        ColumnChartRptNEW4.Legends[0].AutoFitMinFontSize = 7;
        ColumnChartRptNEW4.Legends[0].IsTextAutoFit = false;
        ColumnChartRptNEW4.Legends[0].MaximumAutoSize = 100;


        Random rand = new Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmssfff") + rand.Next().ToString();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\CC_" + strGUID + ".xls";

        // JFreeChart chart = ChartFactory.createBarChart(
        DB clsDB = new DB();

        string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

        DataSet ds = clsDB.getDataSet("Exec SP_R_WORST_MONTH_MAXDD_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + "");

        string Dt1;
        string sDay;
        string sMonth;
        string sYear;

        DataTable dtatble = ds.Tables[0];

        ColumnChartRpt4.DataSource = dtatble;
        ColumnChartRpt4.DataBind();

        // Set series chart type
        ColumnChartRptNEW4.Series[0].ChartType = SeriesChartType.Column;
        ColumnChartRptNEW4.Series[1].ChartType = SeriesChartType.RangeColumn;
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


                ColumnChartRptNEW4.Series[0].Points.AddXY(sDate, s1val);
                ColumnChartRptNEW4.Series[1].Points.AddXY(sDate, s2val);
                ColumnChartRptNEW4.Series[2].Points.AddXY(sDate, s2val);

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

            ColumnChartRptNEW4.Series[0].Points.AddXY(" ", _s3val);
            ColumnChartRptNEW4.Series[1].Points.AddXY(" ", _s4val);
            ColumnChartRptNEW4.Series[2].Points.AddXY(" ", _s4val);

        }


        ColumnChartRptNEW4.Series[0].IsValueShownAsLabel = true;
        ColumnChartRptNEW4.Series[0].LabelFormat = "{0}%";
        ColumnChartRptNEW4.Series[0].Font = new System.Drawing.Font("Frutiger55", 6F, System.Drawing.FontStyle.Regular);
        ColumnChartRptNEW4.Series[2].Font = new System.Drawing.Font("Frutiger55", 6F, System.Drawing.FontStyle.Regular);

        ColumnChartRptNEW4.Series[1].IsValueShownAsLabel = false;


        ColumnChartRptNEW4.Series[2].IsValueShownAsLabel = true;
        ColumnChartRptNEW4.Series[2].LabelFormat = "{0}%";


        ColumnChartRptNEW4.ChartAreas[0].AxisY.MajorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRptNEW4.ChartAreas[0].AxisX.MajorGrid.Enabled = false; //disabled inner gridlines
                                                                          // ColumnChartRptNEW4.ChartAreas[0].AxisY.MajorGrid.LineWidth = 2;

        ColumnChartRptNEW4.ChartAreas[0].AxisY.MinorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRptNEW4.ChartAreas[0].AxisX.MinorGrid.Enabled = false; //disabled inner gridlines

        ColumnChartRptNEW4.ChartAreas[0].AxisY2.MajorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRptNEW4.ChartAreas[0].AxisX2.MajorGrid.Enabled = false; //disabled inner gridlines
                                                                           // ColumnChartRptNEW4.ChartAreas[0].AxisY.MajorGrid.LineWidth = 2;

        ColumnChartRptNEW4.ChartAreas[0].AxisY2.MinorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRptNEW4.ChartAreas[0].AxisX2.MinorGrid.Enabled = false; //disabled inner gridlines

        ColumnChartRptNEW4.ChartAreas[0].AxisX.Enabled = AxisEnabled.False;

        ColumnChartRptNEW4.ChartAreas[0].AxisX.IsStartedFromZero = true;
        ColumnChartRptNEW4.ChartAreas[0].AxisX2.IsStartedFromZero = true;
        ColumnChartRptNEW4.ChartAreas[0].AxisX2.Title = "Return %";
        ColumnChartRptNEW4.ChartAreas[0].AxisX2.TitleFont = new System.Drawing.Font("Frutiger55", 7F, System.Drawing.FontStyle.Bold);
        // ColumnChartRptNEW4.ChartAreas[0].AxisX2.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Bold);


        ColumnChartRptNEW4.ChartAreas[0].AxisY2.LabelStyle.Format = "{N0}%";
        ColumnChartRptNEW4.ChartAreas[0].AxisY.LabelStyle.Format = "{N0}%";

        ColumnChartRptNEW4.ChartAreas[0].AxisX.IsMarginVisible = false;
        ColumnChartRptNEW4.ChartAreas[0].AxisX2.IsMarginVisible = false;

        // ColumnChartRptNEW4.Series[0]["PointWidth"] = "0.5";
        //// ColumnChartRptNEW4.Series[1]["PointWidth"] = "1";
        // ColumnChartRptNEW4.Series[1]["PointWidth"] = "0.5";
        // ColumnChartRptNEW4.Series[2]["PointWidth"] = "0.5";

        ColumnChartRptNEW4.Series[0]["PixelPointWidth"] = "190";
        ColumnChartRptNEW4.Series[1]["PixelPointWidth"] = "25";

        //  ColumnChartRptNEW4.Series[1]["PixelPointWidth"].PadRight(0);
        ColumnChartRptNEW4.Series[2]["PixelPointWidth"] = "190";

        ColumnChartRptNEW4.Series[0].Name = Fname;
        ColumnChartRptNEW4.Series[2].Name = Convert.ToString(ds.Tables[0].Rows[0]["BenchMarkName"]);
        ColumnChartRptNEW4.Series[1].IsVisibleInLegend = false;

        //  ColumnChartRptNEW4.Series[0].BorderWidth = 5;
        // ColumnChartRptNEW4.Series[0].BorderColor = System.Drawing.Color.Aqua;


        ColumnChartRptNEW4.Series[0].XAxisType = AxisType.Primary;
        ColumnChartRptNEW4.Series[0].YAxisType = AxisType.Primary;

        //vertical line above x axis 
        VerticalLineAnnotation annotation = new VerticalLineAnnotation();
        annotation.AnchorDataPoint = ColumnChartRptNEW4.Series[0].Points[3];
        annotation.AxisX = ColumnChartRptNEW4.ChartAreas[0].AxisX;
        annotation.AxisY = ColumnChartRptNEW4.ChartAreas[0].AxisY;
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
        annotation1.AnchorDataPoint = ColumnChartRptNEW4.Series[0].Points[3];
        annotation1.AxisX = ColumnChartRptNEW4.ChartAreas[0].AxisX;
        annotation1.AxisY = ColumnChartRptNEW4.ChartAreas[0].AxisY;
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
        ColumnChartRptNEW4.Annotations.Add(annotation2);

        ColumnChartRptNEW4.Annotations.Add(annotation);
        ColumnChartRptNEW4.Annotations.Add(annotation1);

        //  ColumnChartRptNEW4.Series[0].YAxisType = AxisType.Secondary;

        //   ColumnChartRptNEW4.Series[1].XAxisType = AxisType.Primary;
        //  ColumnChartRptNEW4.Series[1].YAxisType = AxisType.Primary;
        ColumnChartRptNEW4.Series[1].XAxisType = AxisType.Secondary;
        ColumnChartRptNEW4.Series[1].YAxisType = AxisType.Secondary;


        ColumnChartRptNEW4.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 6, FontStyle.Regular);
        ColumnChartRptNEW4.ChartAreas[0].AxisY2.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 6, FontStyle.Regular);
        ColumnChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Bold);

        if (min1 <= 0.0 && min1 >= -10.00)
        {
            ColumnChartRptNEW4.ChartAreas[0].AxisY.Interval = 5;
            ColumnChartRptNEW4.ChartAreas[0].AxisY2.Interval = 5;
            min1 = min1 - 5.0;
        }
        else if (max1 > 0.0 && max1 <= 10.00)
        {
            ColumnChartRptNEW4.ChartAreas[0].AxisY.Interval = 5;
            ColumnChartRptNEW4.ChartAreas[0].AxisY2.Interval = 5;
        }
        else
        {
            ColumnChartRptNEW4.ChartAreas[0].AxisY.Interval = 10;
            ColumnChartRptNEW4.ChartAreas[0].AxisY2.Interval = 10;

        }

        if (min1 < -10.0)
            min1 = min1 - 10.0;

        if (max1 <= 0.0)
        {
            max1 = 0.0;
            ColumnChartRptNEW4.ChartAreas[0].AxisY.IsStartedFromZero = true;
            ColumnChartRptNEW4.ChartAreas[0].AxisY2.IsStartedFromZero = true;
        }


        //   ColumnChartRptNEW4.ChartAreas[0].AxisY.Minimum = min1;
        //   ColumnChartRptNEW4.ChartAreas[0].AxisY2.Minimum = min1;

        ColumnChartRptNEW4.ChartAreas[0].AxisY.Maximum = max1;
        ColumnChartRptNEW4.ChartAreas[0].AxisY2.Maximum = max1;

        ColumnChartRptNEW4.ChartAreas[0].Position.X = 5;
        ColumnChartRptNEW4.ChartAreas[0].Position.Y = 0;
        ColumnChartRptNEW4.ChartAreas[0].Position.Height = 93;
        ColumnChartRptNEW4.ChartAreas[0].Position.Width = 100;

        ColumnChartRptNEW4.ChartAreas[0].AxisX2.LabelStyle.Angle = 0;
        ColumnChartRptNEW4.ChartAreas[0].AxisX2.IsLabelAutoFit = false;

        ColumnChartRptNEW4.Legends[0].Position.Auto = false;
        ColumnChartRptNEW4.Legends[0].Position = new ElementPosition(2, 92, 100, 8);

        ColumnChartRptNEW4.Titles[0].Font = new System.Drawing.Font("Frutiger65", 9, FontStyle.Bold);
        ColumnChartRptNEW4.Titles[0].Docking = Docking.Top;
        ColumnChartRptNEW4.Titles[0].DockingOffset = -2;

        //ColumnChartRptNEW4.Series[1].YAxisType = AxisType.Secondary;

        //   ColumnChartRptNEW4.ChartAreas[0].AxisY.IsReversed = true;
        //    ColumnChartRptNEW4.ChartAreas[0].AxisY2.IsReversed = true;
        // ColumnChartRptNEW4.ChartAreas[0].AxisY.IsReversed = true;
        // ColumnChartRptNEW4.ChartAreas[0].AxisY2.IsReversed = true;
        // ColumnChartRptNEW4.ChartAreas[0].AxisX.IsReversed = true;
        // ColumnChartRptNEW4.ChartAreas[0].AxisX2.IsReversed = true;
        //clsDB.getConfiguration();

        Random rnd = new Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\A_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(1280, 1000);

        bm.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);

        ColumnChartRptNEW4.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 1280, 1000));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        //  ColumnChartRptNEW4.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in ColumnChartRptNEW4.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;
    }

    #endregion
    #endregion

    public string GetAllItemsTextFromListBox(ListBox lstBox, bool IsOnlySelected)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 0; i < lstBox.Items.Count; i++)
            {
                if (IsOnlySelected)
                {
                    if (lstBox.Items[i].Selected)
                    {
                        lstselecteditems = lstselecteditems + "," + lstBox.Items[i].Text;
                        //insert command
                    }
                }
                else
                {
                    lstselecteditems = lstselecteditems + "," + lstBox.Items[i].Text;
                }
            }
            if (lstselecteditems != "")
            {
                lstselecteditems = lstselecteditems.Substring(1);
            }
        }


        return lstselecteditems;
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

    #region Old Java Code (NOT IN USE)

    //public String generateLabelString(XYDataset dataset, int series, int item)
    //{
    //    String result = null;
    //    Object[] items = createItemArray(dataset, series, item);
    //    result = MessageFormat.format("{1}", items);
    //    return result;
    //}

    //protected Object[] createItemArray(XYDataset dataset, int series,
    //                                     int item)
    //{


    //    Object[] result = new Object[3];
    //    result[0] = dataset.getSeriesKey(series).ToString();

    //    double x = dataset.getXValue(series, item);
    //    if (Double.IsNaN(x) && dataset.getX(series, item) == null)
    //    {
    //        result[1] = this.nullXString;
    //    }
    //    else
    //    {
    //        if (this.xDateFormat != null)
    //        {
    //            result[1] = this.xDateFormat.format(new Date((long)x));
    //        }
    //        else
    //        {
    //            result[1] = x;
    //        }
    //    }

    //    double y = dataset.getYValue(series, item);
    //    if (Double.IsNaN(y) && dataset.getY(series, item) == null)
    //    {
    //        result[2] = this.nullYString;
    //    }
    //    else
    //    {
    //        if (this.yDateFormat != null)
    //        {
    //            result[2] = this.yDateFormat.format(new Date((long)y));
    //        }
    //        else
    //        {
    //            result[2] = y;
    //        }
    //    }
    //    return result;
    //}
    //public String generateLabelString(XYDataset dataset, int series, int item)
    //{
    //    String result = null;
    //    Object[] items = createItemArray(dataset, series, item);
    //    result = MessageFormat.format(this.format, items);
    //    return result;
    //}


    //private DateFormat yDateFormat;
    //private String nullXString = "null";

    //private String formatString;
    //private NumberFormat xFormat;
    //private DateFormat xDateFormat;
    //private NumberFormat yFormat;
    private String nullYString = "null";

    //public NumberFormat getXFormat()
    //{
    //    return this.xFormat;
    //}

    //public String getFormatString()
    //{
    //    return this.formatString;
    //}

    //public DateFormat getXDateFormat()
    //{
    //    return this.xDateFormat;
    //}

    //public NumberFormat getYFormat()
    //{
    //    return this.yFormat;
    //}

    //public DateFormat getYDateFormat()
    //{
    //    return this.yDateFormat;
    //}
    //public class LabelGenerator : XYItemLabelGenerator
    //{
    //    //public String generateLabel(XYDataset dataset, int series, int item)
    //    //{
    //    //    LabeledXYDataset labelSource = (LabeledXYDataset)dataset;
    //    //    return labelSource.getLabel(series, item);
    //    //}
    //    private int itemcount;
    //    public LabelGenerator(int itemcount)
    //    {
    //        this.itemcount = itemcount;
    //    }

    //    public String generateLabel(XYDataset dataset, int series, int item)
    //    {
    //        String result = "";

    //        if (item == itemcount)
    //            result = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(dataset.getYValue(series, item)));
    //        return result;
    //    }

    //}


    //private string GetShapeChart4()
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";
    //    String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\P_2OP_" + strGUID + ".xls";

    //    DB clsDB = new DB();
    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
    //    DataSet ds = clsDB.getDataSet("EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = '" + GrpName + "',@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = '" + strAssetClass + "',@StartDate = '01-JAN-2011'");

    //    DataTable dt = GetFormatedDatatable(ds);



    //    DefaultCategoryDataset dataset = new DefaultCategoryDataset();


    //    if (dt.Rows.Count > 0)
    //    {
    //        for (int i = 0; i < dt.Rows.Count; i++)
    //        {
    //            Double s1val = Convert.ToDouble(dt.Rows[i]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //            dataset.setValue(s1val, dt.Rows[i]["Name"].ToString(), dt.Rows[i]["X"].ToString());
    //        }
    //    }



    //    JFreeChart chart = ChartFactory.createLineChart(
    //                                    "Performance vs. Volatility (since 01/01/2011)", // chart title
    //                                    "Annualized Volatility (Standard Deviation)", // domain axis label
    //                                    "Annualized Return %", // range axis label
    //                                    dataset, // data
    //                                    PlotOrientation.VERTICAL, // orientation
    //                                    false, // include legend
    //                                    true, // tooltips
    //                                    false // urls
    //                                    );


    //    CategoryPlot plot = (CategoryPlot)chart.getPlot();
    //    plot.setBackgroundPaint(java.awt.Color.lightGray);
    //    plot.setRangeGridlinePaint(java.awt.Color.white);


    //    // customise the range axis...
    //    NumberAxis rangeAxis = (NumberAxis)plot.getRangeAxis();
    //    rangeAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());



    //    // customise the renderer...
    //    LineAndShapeRenderer renderer = (LineAndShapeRenderer)plot.getRenderer();

    //    //  Shape shape = new Rectangle2D.Double(-3.0, -3.0, 6.0, 6.0);
    //    Shape shape = new Ellipse2D.Double(-3.0, -3.0, 6.0, 6.0);


    //    renderer.setShapesVisible(true);
    //    renderer.setDrawOutlines(true);
    //    renderer.setUseFillPaint(true);
    //    renderer.setFillPaint(java.awt.Color.white);
    //    renderer.setSeriesShape(0, shape);
    //    renderer.setSeriesPaint(0, java.awt.Color.BLACK);

    //    //CategoryItemLabelGenerator generator1 = new CategoryItemLabelGenerator("{2}", new DecimalFormat("0.00"), new DecimalFormat("0.00"));
    //    // renderer.setBaseItemLabelGenerator(generator1);
    //    renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(
    //               ItemLabelAnchor.OUTSIDE12, TextAnchor.BOTTOM_CENTER));
    //    renderer.setBaseItemLabelsVisible(true);



    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, chart, 430, 300);

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();

    //}

    //private string getChartsTAB1()
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";
    //    String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\P_OP_" + strGUID + ".xls";

    //    DB clsDB = new DB();
    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    DataSet ds = clsDB.getDataSet("EXEC  SP_R_CLIENT_GOALS_NEW_GA_BASEDATA @GroupName = " + TIAGrp + ", @TrxnGAFlagTxt = 'TIA',@AsOfDate = '" + txtAsofdate.Text + "'");

    //    string Dt1;
    //    string sDay;
    //    string sMonth;
    //    string sYear;

    //    //chart1
    //    TimeSeries s1 = new TimeSeries("Total Investment Assets (TIA)");
    //    TimeSeries s2 = new TimeSeries("Net Invested Capital");

    //    //chart2
    //    TimeSeries s3 = new TimeSeries("Total Investment Assets (TIA)");
    //    TimeSeries s4 = new TimeSeries("Inflation Adj. Net Invested Capital");

    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
    //        {
    //            Double s1val = Convert.ToDouble(ds.Tables[0].Rows[i]["value"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //            Double s2val = Convert.ToDouble(ds.Tables[0].Rows[i]["NetInvestments"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //            Double s4val = Convert.ToDouble(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //            Dt1 = ds.Tables[0].Rows[i]["Date"].ToString();
    //            sDay = DateTime.Parse(Dt1).ToString("dd");
    //            sMonth = DateTime.Parse(Dt1).ToString("MM");
    //            sYear = DateTime.Parse(Dt1).ToString("yyyy");

    //            //chart 1
    //            s1.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s1val);
    //            s2.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s2val);

    //            //chart 2
    //            s3.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s1val);
    //            s4.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s4val);
    //        }
    //    }

    //    #region chart1
    //    TimeSeriesCollection dataset1 = new TimeSeriesCollection();
    //    dataset1.addSeries(s1);
    //    dataset1.addSeries(s2);
    //    dataset1.setDomainIsPointsInTime(true);

    //    AppDomain.CurrentDomain.Load("JCommon");

    //    JFreeChart Chart1 = ChartFactory.createTimeSeriesChart("Total Investment Assets vs. Net Invested Capital", "", "", dataset1, true, true, false);

    //    Chart1.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    //LineAndShapeRenderer renderer = (LineAndShapeRenderer)plot.getRenderer(); 
    //    //renderer.setShapesVisible(true);
    //    //renderer.setDrawOutlines(true); 
    //    //renderer.setUseFillPaint(true);

    //    // NOW DO SOME OPTIONAL CUSTOMISATION OF THE CHART...
    //    Chart1.setBackgroundPaint(java.awt.Color.white);
    //    Chart1.setBorderVisible(false);


    //    // get a reference to the plot for further customisation...
    //    XYPlot plot = (XYPlot)Chart1.getPlot();
    //    plot.setBackgroundPaint(java.awt.Color.white);
    //    plot.setAxisOffset(new RectangleInsets(0.0, 0.0, 0, 0.0));
    //    // plot.getRangeAxis().setLowerMargin(0);
    //    // plot.getRangeAxis().setUpperMargin(0);

    //    // plot.setInsets(new RectangleInsets(0, 5.0, 0, 0));
    //    //  plot.setRangeCrosshairStroke(new BasicStroke(13));
    //    //plot.setDomainGridlinePaint(java.awt.Color.decode("#FFFFFF"));
    //    //plot.setDomainGridlinesVisible(false);

    //    //plot.setRangeGridlinePaint(Color.white);
    //    //   plot.setDomainCrosshairVisible(false);
    //    //  plot.setRangeCrosshairVisible(true);
    //    plot.setDomainGridlinesVisible(false);
    //    plot.setRangeGridlinesVisible(true);
    //    plot.setOutlineStroke(new BasicStroke(0));
    //    plot.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));

    //    BasicStroke gridstroke = new BasicStroke(0.05f);

    //    // plot.setDomainGridlinePaint(java.awt.Color.decode("#BFBFBF"));
    //    // plot.setRangeGridlinePaint(java.awt.Color.decode("#BFBFBF"));

    //    // plot.setDomainGridlineStroke(gridstroke);
    //    plot.setRangeGridlineStroke(gridstroke);


    //    XYPlot plot1 = Chart1.getXYPlot();

    //    //ValueAxis axis = plot1.getRangeAxis();
    //    //plot1.setOutlinePaint(java.awt.Color.decode("#000000"));
    //    //plot1.setOutlineStroke(new BasicStroke(2));
    //    // plot1.setInsets(new RectangleInsets(0, 0, 0, 0));
    //    //plot1.setOutlinePaint(java.awt.Color.decode("#000000"));
    //    //plot1.setOutlineStroke(new BasicStroke(2));
    //    //axis.setAxisLineStroke(new BasicStroke(2));

    //    // java.awt.Font font = new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 10);
    //    //axis.setTickLabelFont(font);
    //    //axis.setAxisLinePaint(java.awt.Color.decode("#868686"));


    //    //NumberAxis domain = (NumberAxis)plot1.getDomainAxis();
    //    DateAxis dateaxis = (DateAxis)plot1.getDomainAxis();
    //    //axis.setLabelAngle(Math.PI / 2.0);
    //    //axis.setVerticalTickLabels(true);
    //    dateaxis.setVerticalTickLabels(true);

    //    dateaxis.setDateFormatOverride(new SimpleDateFormat("MMM-yy"));
    //    //dateaxis.setAxisLinePaint(java.awt.Color.decode("#000000"));
    //    //dateaxis.setAxisLineStroke(new BasicStroke(2));
    //    dateaxis.setLowerMargin(0d);
    //    //dateaxis.setUpperMargin(0d);
    //    dateaxis.setAxisLinePaint(java.awt.Color.decode("#868686"));

    //    java.awt.Font font1 = new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 10);
    //    dateaxis.setTickLabelFont(font1);
    //    dateaxis.setAxisLineStroke(new BasicStroke(2));


    //    ///DateTickUnit unit = new DateTickUnit(DateTickUnit.MONTH, 3, new SimpleDateFormat("MMM-yy"));
    //    /////dateaxis.setTickUnit(unit);
    //    //  dateaxis.setLowerBound(0);

    //    DecimalFormat dfKey = new DecimalFormat("###,###");

    //    XYLineAndShapeRenderer xylineandshaperenderer = (XYLineAndShapeRenderer)plot1.getRenderer();
    //    xylineandshaperenderer.setBaseShapesVisible(false);
    //    xylineandshaperenderer.setSeriesFillPaint(0, java.awt.Color.decode(ColorTIA1));
    //    xylineandshaperenderer.setSeriesFillPaint(1, java.awt.Color.decode(ColorNetInvestedCap));
    //    xylineandshaperenderer.setSeriesPaint(0, java.awt.Color.decode(ColorTIA1));
    //    xylineandshaperenderer.setSeriesPaint(1, java.awt.Color.decode(ColorNetInvestedCap));

    //    xylineandshaperenderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.CENTER, TextAnchor.CENTER));
    //    //xylineandshaperenderer.setSeriesShape(1, new Ellipse2D.Double(-4.0, -4.0, 8.0, 8.0));
    //    //BasicStroke stroke = new BasicStroke(1.0f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_MITER, 50, new float[] { 1f, 2f }, 0);

    //    //xylineandshaperenderer.setSeriesOutlineStroke(1, stroke);
    //    xylineandshaperenderer.setBaseStroke(new BasicStroke(0));
    //    xylineandshaperenderer.setItemLabelsVisible(true);
    //    xylineandshaperenderer.setUseFillPaint(true);

    //    //xylineandshaperenderer.setStroke(new BasicStroke(3f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_BEVEL)); 
    //    //xylineandshaperenderer.setBaseShapesVisible(true);
    //    //xylineandshaperenderer.setBaseShapesFilled(true);



    //    plot1.setRenderer(xylineandshaperenderer);

    //    XYItemRenderer renderer = plot1.getRenderer();
    //    //XYItemLabelGenerator generator1 = new StandardXYItemLabelGenerator("{2}", new DecimalFormat("0.00"), new DecimalFormat("0.00"));
    //    //renderer.setItemLabelGenerator(generator1);

    //    //*********code for label at end point on line ******************//

    //    XYItemLabelGenerator generator1 =
    //new StandardXYItemLabelGenerator("{2}", new DecimalFormat("0.00"), new DecimalFormat("0.00"));
    //    renderer.setBaseItemLabelGenerator(new LabelGenerator(ds.Tables[0].Rows.Count - 1)); //because it starts from zero.
    //    renderer.setBaseItemLabelsVisible(true);


    //    renderer.setSeriesPositiveItemLabelPosition(0, new ItemLabelPosition(ItemLabelAnchor.OUTSIDE1, TextAnchor.BOTTOM_RIGHT, TextAnchor.BOTTOM_RIGHT, 0));
    //    renderer.setSeriesPositiveItemLabelPosition(1, new ItemLabelPosition(ItemLabelAnchor.OUTSIDE7, TextAnchor.TOP_RIGHT, TextAnchor.TOP_RIGHT, 0));

    //    //*********end code for label at end point on line ******************//

    //    NumberFormat currency = NumberFormat.getCurrencyInstance(Locale.US);
    //    currency.setMaximumFractionDigits(0);


    //    for (int i = 0; i < 2; i++) // for each time series 
    //        plot.getRenderer().setSeriesStroke(i, new BasicStroke(3f));


    //    NumberAxis rangeAxis = (NumberAxis)plot.getRangeAxis();
    //    rangeAxis.setNumberFormatOverride(currency);
    //    rangeAxis.setAutoRangeIncludesZero(true);
    //    rangeAxis.setAxisLineStroke(new BasicStroke(0f));

    //    rangeAxis.setAxisLineStroke(new BasicStroke(2));

    //    // to draw the top line parallel to x axis
    //    double upperrange = rangeAxis.getUpperBound() * extendedrangePerc;
    //    rangeAxis.setRange(rangeAxis.getLowerBound(), upperrange);

    //    java.awt.Font font = new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 10);
    //    rangeAxis.setTickLabelFont(font);
    //    rangeAxis.setAxisLinePaint(java.awt.Color.decode("#868686"));


    //    LegendTitle legendTitle = ((JFreeChart)Chart1).getLegend();
    //    LegendTitle legendTitleNew = new LegendTitle(plot, new ColumnArrangement(), new ColumnArrangement());
    //    legendTitleNew.setPosition(legendTitle.getPosition());
    //    legendTitleNew.setBackgroundPaint(legendTitle.getBackgroundPaint());
    //    legendTitleNew.setBorder(0, 0, 0, 0);

    //    //Remove old Legend 
    //    ((JFreeChart)Chart1).removeLegend();
    //    //Add new Legend 
    //    ((JFreeChart)Chart1).addLegend(legendTitleNew);

    //    //ChartRenderingInfo info = new ChartRenderingInfo(new StandardEntityCollection());
    //    //BufferedImage image = Chart1.createBufferedImage(2000, 1600, 500, 400, null);

    //    //byte[] encoded = null;
    //    //EncoderUtil encoder = new EncoderUtil();

    //    //ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
    //    //java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
    //    //thisImageMapInfo.setChartArea(new Rectangle2D.Double(1.0, 2.0, 3.0, 4.0));
    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, Chart1, 430, 300);

    //    #endregion

    //    #region Chart2

    //    TimeSeriesCollection dataset2 = new TimeSeriesCollection();
    //    dataset2.addSeries(s3);
    //    dataset2.addSeries(s4);
    //    dataset2.setDomainIsPointsInTime(true);

    //    AppDomain.CurrentDomain.Load("JCommon");

    //    JFreeChart Chart2 = ChartFactory.createTimeSeriesChart("Total Investment Assets vs. Inflation Adj. Net Invested Capital", "", "", dataset2, true, true, false);
    //    Chart2.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));
    //    //LineAndShapeRenderer renderer = (LineAndShapeRenderer)plot.getRenderer(); 
    //    //renderer.setShapesVisible(true);
    //    //renderer.setDrawOutlines(true); 
    //    //renderer.setUseFillPaint(true);

    //    // NOW DO SOME OPTIONAL CUSTOMISATION OF THE CHART...
    //    Chart2.setBackgroundPaint(java.awt.Color.white);
    //    Chart2.setBorderVisible(false);
    //    // get a reference to the plot for further customisation...
    //    XYPlot plot3 = (XYPlot)Chart2.getPlot();
    //    plot3.setBackgroundPaint(java.awt.Color.WHITE);
    //    //plot.setAxisOffset(new RectangleInsets(0, 5.0, 0, 5.0));
    //    //plot3.setDomainGridlinePaint(java.awt.Color.decode("#FFFFFF"));
    //    //plot.setRangeGridlinePaint(Color.white);
    //    // plot.setDomainCrosshairVisible(false);
    //    // plot.setRangeCrosshairVisible(true);
    //    plot3.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));

    //    plot3.setDomainGridlinesVisible(false);
    //    plot3.setRangeGridlinesVisible(true);
    //    plot3.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));


    //    // plot.setDomainGridlinePaint(java.awt.Color.decode("#BFBFBF"));
    //    // plot.setRangeGridlinePaint(java.awt.Color.decode("#BFBFBF"));

    //    // plot.setDomainGridlineStroke(gridstroke);
    //    plot3.setRangeGridlineStroke(gridstroke);


    //    //plot3.setRangeAxisLocation(0, AxisLocation.BOTTOM_OR_RIGHT); 
    //    XYPlot plot4 = Chart2.getXYPlot();

    //    ValueAxis axis2 = plot4.getRangeAxis();


    //    //NumberAxis domain = (NumberAxis)plot1.getDomainAxis();
    //    DateAxis dateaxis2 = (DateAxis)plot4.getDomainAxis();
    //    //axis.setLabelAngle(Math.PI / 2.0);
    //    //axis.setVerticalTickLabels(true);
    //    dateaxis2.setVerticalTickLabels(true);
    //    dateaxis2.setDateFormatOverride(new SimpleDateFormat("MMM-yy"));
    //    dateaxis2.setLowerMargin(0d);

    //    dateaxis2.setAxisLinePaint(java.awt.Color.decode("#868686"));
    //    dateaxis2.setTickLabelFont(font1);
    //    dateaxis2.setAxisLineStroke(new BasicStroke(2));


    //    /////DateTickUnit unit = new DateTickUnit(DateTickUnit.MONTH, 3, new SimpleDateFormat("MMM-yy"));
    //    /////dateaxis.setTickUnit(unit);
    //    //  dateaxis.setLowerBound(0);
    //    XYLineAndShapeRenderer xylineandshaperenderer2 = (XYLineAndShapeRenderer)plot4.getRenderer();
    //    xylineandshaperenderer2.setBaseShapesVisible(false);
    //    xylineandshaperenderer2.setSeriesFillPaint(0, java.awt.Color.decode(ColorTIA2));
    //    xylineandshaperenderer2.setSeriesFillPaint(1, java.awt.Color.decode(ColorInflationAdjInvCap));
    //    xylineandshaperenderer2.setSeriesPaint(0, java.awt.Color.decode(ColorTIA2));
    //    xylineandshaperenderer2.setSeriesPaint(1, java.awt.Color.decode(ColorInflationAdjInvCap));
    //    //xylineandshaperenderer.setSeriesShape(1, new Ellipse2D.Double(-4.0, -4.0, 8.0, 8.0));
    //    //BasicStroke stroke = new BasicStroke(1.0f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_MITER, 50, new float[] { 1f, 2f }, 0);

    //    //xylineandshaperenderer.setSeriesOutlineStroke(1, stroke);

    //    xylineandshaperenderer2.setBaseStroke(new BasicStroke(13));
    //    xylineandshaperenderer2.setItemLabelsVisible(true);
    //    xylineandshaperenderer2.setUseFillPaint(true);
    //    //xylineandshaperenderer.setBaseShapesVisible(true);
    //    //xylineandshaperenderer.setBaseShapesFilled(true);
    //    plot4.setRenderer(xylineandshaperenderer2);


    //    XYItemRenderer renderer2 = plot4.getRenderer();
    //    XYItemLabelGenerator generator2 =
    //new StandardXYItemLabelGenerator("{2}", new DecimalFormat("0"), new DecimalFormat("0"));
    //    renderer2.setBaseItemLabelGenerator(new LabelGenerator(ds.Tables[0].Rows.Count - 1)); //because it starts from zero.
    //    renderer2.setBaseItemLabelsVisible(true);

    //    renderer2.setSeriesPositiveItemLabelPosition(0, new ItemLabelPosition(ItemLabelAnchor.OUTSIDE6, TextAnchor.BOTTOM_RIGHT, TextAnchor.BOTTOM_RIGHT, 0));
    //    renderer2.setSeriesPositiveItemLabelPosition(1, new ItemLabelPosition(ItemLabelAnchor.OUTSIDE1, TextAnchor.BOTTOM_RIGHT, TextAnchor.TOP_RIGHT, 0));


    //    NumberFormat currency2 = NumberFormat.getCurrencyInstance(Locale.US);

    //    currency2.setMaximumFractionDigits(0);


    //    for (int i = 0; i < 2; i++) // for each time series 
    //        plot3.getRenderer().setSeriesStroke(i, new BasicStroke(3f));


    //    NumberAxis rangeAxis2 = (NumberAxis)plot3.getRangeAxis();
    //    rangeAxis2.setNumberFormatOverride(currency2);
    //    rangeAxis2.setAutoRangeIncludesZero(true);


    //    rangeAxis2.setAxisLineStroke(new BasicStroke(2));

    //    double upperrange2 = rangeAxis2.getUpperBound() * extendedrangePerc;
    //    rangeAxis2.setRange(rangeAxis2.getLowerBound(), upperrange2);

    //    rangeAxis2.setTickLabelFont(font);
    //    rangeAxis2.setAxisLinePaint(java.awt.Color.decode("#868686"));
    //    //LegendItemSource 

    //    LegendTitle legendTitle2 = ((JFreeChart)Chart2).getLegend();
    //    LegendTitle legendTitleNew2 = new LegendTitle(plot3, new ColumnArrangement(), new ColumnArrangement());
    //    legendTitleNew2.setPosition(legendTitle2.getPosition());
    //    legendTitleNew2.setBackgroundPaint(legendTitle2.getBackgroundPaint());
    //    legendTitleNew2.setBorder(0, 0, 0, 0);

    //    //Remove old Legend 
    //    ((JFreeChart)Chart2).removeLegend();
    //    //Add new Legend 
    //    ((JFreeChart)Chart2).addLegend(legendTitleNew2);

    //    //ChartRenderingInfo info2 = new ChartRenderingInfo(new StandardEntityCollection());
    //    //BufferedImage image2 = Chart2.createBufferedImage(2000, 1600, 500, 400, null);

    //    //byte[] encoded2= null;
    //    //EncoderUtil encoder2 = new EncoderUtil();

    //    //ChartRenderingInfo thisImageMapInfo2 = new ChartRenderingInfo();
    //    //thisImageMapInfo2.setChartArea(new Rectangle2D.Double(1.0, 2.0, 3.0, 4.0));
    //    //java.io.OutputStream jos2 = new java.io.FileOutputStream(fsFinalLocation1.Replace(".xls", ".png"));
    //    java.io.File file2 = new java.io.File(fsFinalLocation1.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file2, Chart2, 430, 300);


    //    #endregion


    //    return fsFinalLocation.Replace(".xls", ".png").ToString();


    //}
    //private string GetShapeChart4TEST()
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";


    //    DB clsDB = new DB();
    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
    //    DataSet ds = clsDB.getDataSet("EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + ",@StartDate = '01-JAN-2011'");

    //    DataTable dt = GetFormatedDatatable(ds);

    //    XYSeriesCollection dataset = new XYSeriesCollection();

    //    XYSeries series1 = new XYSeries(dt.Rows[0]["Name"].ToString());
    //    XYSeries series2 = new XYSeries(dt.Rows[1]["Name"].ToString());
    //    XYSeries series3 = new XYSeries(dt.Rows[2]["Name"].ToString());
    //    XYSeries series4 = new XYSeries(dt.Rows[3]["Name"].ToString());

    //    Double s1X = Convert.ToDouble(dt.Rows[0]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s1Y = Convert.ToDouble(dt.Rows[0]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    Double s2X = Convert.ToDouble(dt.Rows[1]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s2Y = Convert.ToDouble(dt.Rows[1]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    Double s3X = Convert.ToDouble(dt.Rows[2]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s3Y = Convert.ToDouble(dt.Rows[2]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    Double s4X = Convert.ToDouble(dt.Rows[3]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s4Y = Convert.ToDouble(dt.Rows[3]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    series1.add(s1X, s1Y);
    //    series2.add(s2X, s2Y);
    //    series3.add(s3X, s3Y);
    //    series4.add(s4X, s4Y);

    //    dataset.addSeries(series1);
    //    dataset.addSeries(series2);
    //    dataset.addSeries(series3);
    //    dataset.addSeries(series4);



    //    JFreeChart chart = ChartFactory.createXYLineChart("Performance vs. Volatility (since 01/01/2011)", // chart title
    //                                                    "Annualized Volatility (Standard Deviation)", // domain axis label
    //                                                    "Annualized Return %", // range axis label
    //                                                    dataset, // data
    //                                                    PlotOrientation.VERTICAL, // orientation
    //                                                    false, // include legend
    //                                                    false, // tooltips
    //                                                    false // urls
    //                                                    );

    //    chart.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));


    //    chart.setBackgroundPaint(java.awt.Color.white);
    //    chart.setBorderVisible(false);

    //    XYPlot plot = (XYPlot)chart.getPlot();
    //    plot.setBackgroundPaint(java.awt.Color.white);
    //    plot.setRangeGridlinePaint(java.awt.Color.white);
    //    plot.setAxisOffset(new RectangleInsets(0.0, 0.0, 0, 0.0));
    //    plot.setDomainGridlinesVisible(false);
    //    plot.setRangeGridlinesVisible(true);
    //    plot.setOutlineStroke(new BasicStroke(0));
    //    plot.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));


    //    // customise the range axis...
    //    NumberAxis rangeAxis = (NumberAxis)plot.getRangeAxis();
    //    rangeAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());

    //    rangeAxis.setUpperMargin(0.2);
    //    rangeAxis.setLowerMargin(0.2);

    //    DecimalFormat pctFormat = new DecimalFormat("0'%'");
    //    rangeAxis.setNumberFormatOverride(pctFormat);

    //    ValueAxis valAxis1 = (ValueAxis)plot.getRangeAxis();
    //    DecimalFormat pctFormat1 = new DecimalFormat("0.00");
    //    //  CategoryAxis categoryaxis = plot.getDomainAxis();

    //    NumberAxis DomAxis = (NumberAxis)plot.getDomainAxis();
    //    DomAxis.setNumberFormatOverride(pctFormat);
    //    DomAxis.setLowerBound(0);
    //    DomAxis.setLowerMargin(0.5);

    //    java.awt.Font font3 = new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 11);
    //    plot.getDomainAxis().setLabelFont(font3);
    //    plot.getRangeAxis().setLabelFont(font3);



    //    // customise the renderer...
    //    XYLineAndShapeRenderer renderer = (XYLineAndShapeRenderer)plot.getRenderer();

    //    // Shape shape = new Rectangle2D.Double(-3.0, -3.0, 6.0, 6.0);
    //    Shape shape = new Ellipse2D.Double(-3.0, -3.0, 6.0, 6.0);


    //    renderer.setShapesVisible(true);
    //    renderer.setDrawOutlines(true);
    //    renderer.setUseFillPaint(true);
    //    renderer.setFillPaint(java.awt.Color.white);
    //    renderer.setSeriesShape(0, shape);
    //    renderer.setSeriesPaint(0, java.awt.Color.BLACK);

    //    XYItemLabelGenerator generator1 =
    //    new StandardXYItemLabelGenerator("{0}", new DecimalFormat("0.00"), new DecimalFormat("0.00"));
    //    renderer.setBaseItemLabelGenerator(generator1);

    //    renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(
    //               ItemLabelAnchor.OUTSIDE4, TextAnchor.TOP_CENTER));
    //    renderer.setBaseItemLabelsVisible(true);


    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, chart, 420, 300);

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();

    //}

    //private string getChartsTAB3()
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";
    //    String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\P_2OP_" + strGUID + ".xls";

    //    DB clsDB = new DB();
    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
    //    DataSet ds = clsDB.getDataSet("exec SP_R_WEALTH_CHART_NEW_GA_BASEDATA @GroupName = " + GrpName + " , @PositionGAFlagTxt = 'GA', @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

    //    string Dt1;
    //    string sDay;
    //    string sMonth;
    //    string sYear;

    //    //chart1
    //    TimeSeries s1 = new TimeSeries("Gresham Advised Assets");
    //    TimeSeries s2 = new TimeSeries("Net Invested Capital");
    //    TimeSeries s3 = new TimeSeries("Inflation Adj. Net Invested Capital");


    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
    //        {
    //            Double s1val = Convert.ToDouble(ds.Tables[0].Rows[i]["Gresham Advised Assets"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //            Double s2val = Convert.ToDouble(ds.Tables[0].Rows[i]["Net Invested Capital"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //            Double s3val = Convert.ToDouble(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //            Dt1 = ds.Tables[0].Rows[i]["Year"].ToString();
    //            sDay = DateTime.Parse(Dt1).ToString("dd");
    //            sMonth = DateTime.Parse(Dt1).ToString("MM");
    //            sYear = DateTime.Parse(Dt1).ToString("yyyy");

    //            //chart 1
    //            s1.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s1val);
    //            s2.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s2val);
    //            s3.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s3val);
    //        }
    //    }

    //    #region chart1
    //    TimeSeriesCollection dataset1 = new TimeSeriesCollection();
    //    dataset1.addSeries(s1);
    //    dataset1.addSeries(s2);
    //    dataset1.addSeries(s3);
    //    dataset1.setDomainIsPointsInTime(true);

    //    AppDomain.CurrentDomain.Load("JCommon");

    //    JFreeChart Chart1 = ChartFactory.createTimeSeriesChart("Growth of My Gresham Advised Assets (GAA)", "", "", dataset1, true, true, false);

    //    Chart1.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    //LineAndShapeRenderer renderer = (LineAndShapeRenderer)plot.getRenderer(); 
    //    //renderer.setShapesVisible(true);
    //    //renderer.setDrawOutlines(true); 
    //    //renderer.setUseFillPaint(true);

    //    // NOW DO SOME OPTIONAL CUSTOMISATION OF THE CHART...
    //    Chart1.setBackgroundPaint(java.awt.Color.white);
    //    Chart1.setBorderVisible(false);


    //    // get a reference to the plot for further customisation...
    //    XYPlot plot = (XYPlot)Chart1.getPlot();
    //    plot.setBackgroundPaint(java.awt.Color.white);
    //    plot.setAxisOffset(new RectangleInsets(0.0, 0.0, 0, 0.0));
    //    // plot.getRangeAxis().setLowerMargin(0);
    //    // plot.getRangeAxis().setUpperMargin(0);

    //    // plot.setInsets(new RectangleInsets(0, 5.0, 0, 0));
    //    //  plot.setRangeCrosshairStroke(new BasicStroke(13));
    //    //plot.setDomainGridlinePaint(java.awt.Color.decode("#FFFFFF"));
    //    //plot.setDomainGridlinesVisible(false);

    //    //plot.setRangeGridlinePaint(Color.white);
    //    //   plot.setDomainCrosshairVisible(false);
    //    //  plot.setRangeCrosshairVisible(true);
    //    plot.setDomainGridlinesVisible(false);
    //    plot.setRangeGridlinesVisible(true);
    //    plot.setOutlineStroke(new BasicStroke(0));
    //    plot.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));
    //    BasicStroke gridstroke = new BasicStroke(0.05f);

    //    // plot.setDomainGridlinePaint(java.awt.Color.decode("#BFBFBF"));
    //    // plot.setRangeGridlinePaint(java.awt.Color.decode("#BFBFBF"));

    //    // plot.setDomainGridlineStroke(gridstroke);
    //    plot.setRangeGridlineStroke(gridstroke);

    //    //CategoryPlot Cplot = Chart1.getCategoryPlot();
    //    //Cplot.setBackgroundPaint(Color.lightGray);
    //    //Cplot.setDomainGridlinePaint(Color.white);
    //    //Cplot.setRangeGridlinePaint(Color.white);

    //    //CategoryItemRenderer renderer1 = Cplot.getRenderer();
    //    //renderer1.setItemLabelGenerator(new PerfAnalytics1(50.0));
    //    //// renderer.setItemLabelFont(new java.awt.Font("Serif", java.awtFont.PLAIN, 20));
    //    //renderer1.setItemLabelsVisible(true);



    //    XYPlot plot1 = Chart1.getXYPlot();

    //    //ValueAxis axis = plot1.getRangeAxis();
    //    //plot1.setOutlinePaint(java.awt.Color.decode("#000000"));
    //    //plot1.setOutlineStroke(new BasicStroke(2));
    //    // plot1.setInsets(new RectangleInsets(0, 0, 0, 0));
    //    //plot1.setOutlinePaint(java.awt.Color.decode("#000000"));
    //    //plot1.setOutlineStroke(new BasicStroke(2));
    //    //axis.setAxisLineStroke(new BasicStroke(2));

    //    // java.awt.Font font = new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 10);
    //    //axis.setTickLabelFont(font);
    //    //axis.setAxisLinePaint(java.awt.Color.decode("#868686"));


    //    //NumberAxis domain = (NumberAxis)plot1.getDomainAxis();
    //    DateAxis dateaxis = (DateAxis)plot1.getDomainAxis();
    //    //axis.setLabelAngle(Math.PI / 2.0);
    //    //axis.setVerticalTickLabels(true);
    //    dateaxis.setVerticalTickLabels(true);

    //    dateaxis.setDateFormatOverride(new SimpleDateFormat("MMM-yy"));
    //    //dateaxis.setAxisLinePaint(java.awt.Color.decode("#000000"));
    //    //dateaxis.setAxisLineStroke(new BasicStroke(2));
    //    dateaxis.setLowerMargin(0d);
    //    //dateaxis.setUpperMargin(0d);
    //    dateaxis.setAxisLinePaint(java.awt.Color.decode("#868686"));

    //    java.awt.Font font1 = new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 10);
    //    dateaxis.setTickLabelFont(font1);
    //    dateaxis.setAxisLineStroke(new BasicStroke(2));


    //    ///DateTickUnit unit = new DateTickUnit(DateTickUnit.MONTH, 3, new SimpleDateFormat("MMM-yy"));
    //    /////dateaxis.setTickUnit(unit);
    //    //  dateaxis.setLowerBound(0);

    //    DecimalFormat dfKey = new DecimalFormat("###,###");

    //    XYLineAndShapeRenderer xylineandshaperenderer = (XYLineAndShapeRenderer)plot1.getRenderer();
    //    xylineandshaperenderer.setBaseShapesVisible(false);
    //    xylineandshaperenderer.setSeriesFillPaint(0, java.awt.Color.decode(ColorTIA1));
    //    xylineandshaperenderer.setSeriesFillPaint(1, java.awt.Color.decode(ColorNetInvestedCap));
    //    xylineandshaperenderer.setSeriesPaint(0, java.awt.Color.decode(ColorTIA1));
    //    xylineandshaperenderer.setSeriesPaint(1, java.awt.Color.decode(ColorNetInvestedCap));

    //    xylineandshaperenderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.CENTER, TextAnchor.CENTER));
    //    //xylineandshaperenderer.setSeriesShape(1, new Ellipse2D.Double(-4.0, -4.0, 8.0, 8.0));
    //    //BasicStroke stroke = new BasicStroke(1.0f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_MITER, 50, new float[] { 1f, 2f }, 0);

    //    //xylineandshaperenderer.setSeriesOutlineStroke(1, stroke);
    //    xylineandshaperenderer.setBaseStroke(new BasicStroke(0));
    //    xylineandshaperenderer.setItemLabelsVisible(true);
    //    xylineandshaperenderer.setUseFillPaint(true);

    //    NumberFormat format = NumberFormat.getNumberInstance();
    //    format.setMaximumFractionDigits(2); // etc.
    //    //   string aaa = StandardXYItemLabelGenerator.DEFAULT_ITEM_LABEL_FORMAT;
    //    //  XYItemLabelGenerator generator =
    //    //    new StandardXYItemLabelGenerator(generateLabel(dataset1, 1, 1, format), format, format);
    //    //   xylineandshaperenderer.setBaseItemLabelGenerator(generator);
    //    //XYItemLabelGenerator generator =
    //    //    new StandardXYItemLabelGenerator(
    //    //        StandardXYItemLabelGenerator.DEFAULT_ITEM_LABEL_FORMAT,
    //    //        format, format);

    //    // xylineandshaperenderer.setBaseItemLabelsVisible(true);
    //    //  xylineandshaperenderer.setBaseItemLabelGenerator(generator);
    //    // xylineandshaperenderer.setSeriesItemLabelGenerator(0, generator);

    //    //xylineandshaperenderer.setBaseShapesVisible(true);
    //    //xylineandshaperenderer.setBaseShapesFilled(true);
    //    plot1.setRenderer(xylineandshaperenderer);

    //    //XYItemRenderer renderer = plot1.getRenderer();
    //    //XYItemLabelGenerator generator1 = new StandardXYItemLabelGenerator("{2}", new DecimalFormat("0.00"), new DecimalFormat("0.00"));
    //    //renderer.setItemLabelGenerator(generator1);
    //    // plot.setRangeAxisLocation(0, AxisLocation.BOTTOM_OR_RIGHT);

    //    //*********code for label at end point on line ******************//

    //    XYItemRenderer renderer = plot1.getRenderer();

    //    XYItemLabelGenerator generator =
    //new StandardXYItemLabelGenerator("{2}", new DecimalFormat("0"), new DecimalFormat("0"));
    //    renderer.setBaseItemLabelGenerator(new LabelGenerator(ds.Tables[0].Rows.Count - 1)); //because it starts from zero.
    //    renderer.setBaseItemLabelsVisible(true);

    //    renderer.setSeriesPositiveItemLabelPosition(0, new ItemLabelPosition(ItemLabelAnchor.OUTSIDE1, TextAnchor.BOTTOM_RIGHT, TextAnchor.BOTTOM_RIGHT, 0));
    //    renderer.setSeriesPositiveItemLabelPosition(1, new ItemLabelPosition(ItemLabelAnchor.OUTSIDE7, TextAnchor.TOP_RIGHT, TextAnchor.TOP_RIGHT, 0));
    //    renderer.setSeriesPositiveItemLabelPosition(2, new ItemLabelPosition(ItemLabelAnchor.OUTSIDE7, TextAnchor.TOP_RIGHT, TextAnchor.TOP_RIGHT, 0));

    //    //*********end code for label at end point on line ******************//


    //    NumberFormat currency = NumberFormat.getCurrencyInstance(Locale.US);
    //    currency.setMaximumFractionDigits(0);


    //    for (int i = 0; i < 3; i++) // for each time series 
    //        plot.getRenderer().setSeriesStroke(i, new BasicStroke(3f));


    //    NumberAxis rangeAxis = (NumberAxis)plot.getRangeAxis();
    //    rangeAxis.setNumberFormatOverride(currency);
    //    rangeAxis.setAutoRangeIncludesZero(true);
    //    rangeAxis.setAxisLineStroke(new BasicStroke(0f));

    //    rangeAxis.setAxisLineStroke(new BasicStroke(2));

    //    // to draw the top line parallel to x axis
    //    extendedrangePerc = extendedrangePerc + 0.05;
    //    double upperrange = rangeAxis.getUpperBound() * extendedrangePerc;
    //    rangeAxis.setRange(rangeAxis.getLowerBound(), upperrange);

    //    java.awt.Font font = new java.awt.Font("Frutiger55", java.awt.Font.PLAIN, 10);
    //    rangeAxis.setTickLabelFont(font);
    //    rangeAxis.setAxisLinePaint(java.awt.Color.decode("#868686"));




    //    LegendTitle legendTitle = ((JFreeChart)Chart1).getLegend();
    //    LegendTitle legendTitleNew = new LegendTitle(plot);
    //    legendTitleNew.setPosition(legendTitle.getPosition());
    //    legendTitleNew.setBackgroundPaint(legendTitle.getBackgroundPaint());
    //    legendTitleNew.setBorder(0, 0, 0, 0);

    //    //Remove old Legend 
    //    ((JFreeChart)Chart1).removeLegend();
    //    //Add new Legend 
    //    ((JFreeChart)Chart1).addLegend(legendTitleNew);

    //    //ChartRenderingInfo info = new ChartRenderingInfo(new StandardEntityCollection());
    //    //BufferedImage image = Chart1.createBufferedImage(2000, 1600, 500, 400, null);

    //    //byte[] encoded = null;
    //    //EncoderUtil encoder = new EncoderUtil();

    //    //ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
    //    //java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
    //    //thisImageMapInfo.setChartArea(new Rectangle2D.Double(1.0, 2.0, 3.0, 4.0));
    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, Chart1, 860, 230);

    //    #endregion

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();


    //}

    //private string GetChart5()
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";

    //    // JFreeChart chart = ChartFactory.createBarChart(
    //    DB clsDB = new DB();

    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

    //    //DataSet ds = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA @GroupName = " + GrpName + ", @PositionGAFlagTxt = 'GA' , @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "', @AnnPerfFlg = 0 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

    //    DataSet ds = clsDB.getDataSet("Exec SP_R_WORST_MONTH_MAXDD_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + "");

    //    string Dt1;

    //    string sDate;

    //    DefaultCategoryDataset bardataset = new DefaultCategoryDataset();
    //    DefaultCategoryDataset bardataset1 = new DefaultCategoryDataset();


    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
    //        {
    //            string strVal1 = Convert.ToString(ds.Tables[0].Rows[i]["Honore"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Honore"]);
    //            string strVal2 = Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]);
    //            Double s1val = Convert.ToDouble(strVal1, System.Globalization.CultureInfo.InvariantCulture);
    //            Double s2val = Convert.ToDouble(strVal2, System.Globalization.CultureInfo.InvariantCulture);

    //            sDate = ds.Tables[0].Rows[i]["Month"].ToString();

    //            bardataset.setValue(s1val * 100, "Marks", sDate);
    //            bardataset.setValue(s2val * 100, "Marks1", sDate);
    //        }
    //    }


    //    string strVal3 = Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["Honore"]);
    //    string strVal4 = Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index"]);
    //    Double s3val = Convert.ToDouble(strVal3, System.Globalization.CultureInfo.InvariantCulture);
    //    Double s4val = Convert.ToDouble(strVal4, System.Globalization.CultureInfo.InvariantCulture);

    //    sDate = ds.Tables[0].Rows[0]["MinDate"].ToString();

    //    bardataset1.setValue(s3val * 100, "Marks", sDate);
    //    bardataset1.setValue(s4val * 100, "Marks1", sDate);


    //    JFreeChart barchart = ChartFactory.createBarChart(
    //     "",      //Title  
    //     "",             // X-axis Label  
    //     "",               // Y-axis Label  
    //     bardataset,             // Dataset  
    //     PlotOrientation.VERTICAL,      //Plot orientation  
    //     false,                // Show legend  
    //     false,                // Use tooltips  
    //     false                // Generate URLs  
    //  );

    //    barchart.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    barchart.getTitle().setPaint(java.awt.Color.BLACK);    // Set the colour of the title  
    //    barchart.setBackgroundPaint(java.awt.Color.white);    // Set the background colour of the chart  

    //    CategoryPlot cp = barchart.getCategoryPlot();  // Get the Plot object for a bar graph  
    //    cp.setBackgroundPaint(java.awt.Color.white);       // Set the plot background colour  
    //    cp.setRangeGridlinePaint(java.awt.Color.gray);      // Set the colour of the plot gridlines  
    //    cp.setDomainAxisLocation(AxisLocation.TOP_OR_LEFT);
    //    //cp.setRangeZeroBaselineVisible(true);
    //    //cp.setDomainZeroBaselineStroke(new java.awt.BasicStroke(10));
    //    //cp.setDomainZeroBaselinePaint(Color.GREEN);
    //    //cp.setRangeZeroBaselineStroke(new java.awt.BasicStroke(10));
    //    //cp.setRangeZeroBaselinePaint(Color.GREEN);

    //    cp.getRenderer().setSeriesStroke(0, new BasicStroke(2.0f, BasicStroke.JOIN_MITER, BasicStroke.CAP_SQUARE, 10.0f, new float[] { 1.0f, -1.0f }, 0.0f));

    //    CategoryMarker categorymarker = new CategoryMarker("Sep-02", java.awt.Color.ORANGE, new BasicStroke(1.0F));
    //    categorymarker.setDrawAsLine(true);
    //    categorymarker.setLabel("Marker Label");
    //    //   categorymarker.setLabelFont(new Font("Dialog", 0, 11));
    //    categorymarker.setLabelTextAnchor(TextAnchor.CENTER);
    //    categorymarker.setLabelOffset(new RectangleInsets(-2D, -5D, -2D, -5D));

    //    cp.addDomainMarker(categorymarker, Layer.FOREGROUND);


    //    CategoryAxis categoryaxis = cp.getDomainAxis();
    //    // categoryaxis.setMaximumCategoryLabelWidthRatio(100);


    //    NumberAxis rangeAxis = (NumberAxis)cp.getRangeAxis();
    //    rangeAxis.setUpperMargin(0.2);
    //    rangeAxis.setLowerMargin(0.2);
    //    DecimalFormat pctFormat = new DecimalFormat("##0'%'");
    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    rangeAxis.setNumberFormatOverride(pctFormat);

    //    NumberAxis axis2 = (NumberAxis)cp.getRangeAxis();
    //    cp.setRangeAxis(1, axis2);
    //    //axis2.setUpperMargin(0.2);
    //    //axis2.setLowerMargin(0.2);
    //    //cp.setRangeAxisLocation(1, AxisLocation.BOTTOM_OR_RIGHT);
    //    //cp.mapDatasetToRangeAxis(1, 1);

    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    //  axis2.setNumberFormatOverride(pctFormat);

    //    //rangeAxis.setLowerBound(0);



    //    //CategoryMarker catmarker = new CategoryMarker("Sep-02");
    //    //catmarker.setOutlinePaint(java.awt.Color.red);
    //    //catmarker.setPaint(java.awt.Color.red);
    //    //catmarker.setStroke(new BasicStroke(1.0f));
    //    //catmarker.setDrawAsLine(true);
    //    //catmarker.setLabel("Marker Label");
    //    ////catmarker.setLabelFont(new Font("Dialog", Font.PLAIN, 11));
    //    //catmarker.setLabelTextAnchor(TextAnchor.TOP_RIGHT);
    //    //catmarker.setLabelOffset(new RectangleInsets(2, 5, 2, 5));   



    //    Marker marker = new ValueMarker(-3);
    //    marker.setOutlinePaint(java.awt.Color.red);
    //    marker.setPaint(java.awt.Color.red);
    //    marker.setStroke(new BasicStroke(5.0f));


    //    // cp.addDomainMarker(catmarker, Layer.FOREGROUND);

    //    //cp.addAnnotation(new CategoryLineAnnotation("Category 2", -5.0,
    //    //       "Category 4",  -2.0, java.awt.Color.red, new BasicStroke(2.0f)));


    //    BarRenderer barrenderer = (BarRenderer)cp.getRenderer();
    //    barrenderer.setDrawBarOutline(false);
    //    barrenderer.setSeriesPaint(0, java.awt.Color.decode("#558ED5"));
    //    barrenderer.setSeriesPaint(1, java.awt.Color.decode("#003399"));
    //    //barrenderer.setSeriesPaint(0, gradientpaint);




    //    // LineAndShapeRenderer renderer = (LineAndShapeRenderer)cp.getRenderer();
    //    barrenderer.setStroke(new BasicStroke(4f, 2, 2));


    //    StandardCategoryItemLabelGenerator labelGen = new StandardCategoryItemLabelGenerator("{2}%", new DecimalFormat("0.0"));
    //    barrenderer.setBaseItemLabelGenerator(labelGen);
    //    barrenderer.setBaseItemLabelsVisible(true);


    //    //   DateAxis dateaxis = (DateAxis)cp.getDomainAxis();
    //    //axis.setLabelAngle(Math.PI / 2.0);
    //    //axis.setVerticalTickLabels(true);
    //    // dateaxis.setVerticalTickLabels(true);
    //    //  dateaxis.setDateFormatOverride(new SimpleDateFormat("yyyy"));

    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, barchart, 420, 300);

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();
    //}

    //private string GetChart5TEST()
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";

    //    // JFreeChart chart = ChartFactory.createBarChart(
    //    DB clsDB = new DB();

    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

    //    //DataSet ds = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA @GroupName = " + GrpName + ", @PositionGAFlagTxt = 'GA' , @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "', @AnnPerfFlg = 0 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

    //    DataSet ds = clsDB.getDataSet("Exec SP_R_WORST_MONTH_MAXDD_NEW_GA_BASEDATA @GroupName = 'Linderman, L - Linda Linderman GA',@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '04/30/2014',@BenchMarkName = null,@AssetNameTxt = 'Domestic Equity, International Equity, Low Volatility Hedged Strategies, Opportunistic Growth, Fixed Income, Liquid Real Assets, Illiquid Real Assets, Private Equity, Cash and Equivalents'");

    //    string Dt1;

    //    string sDate;

    //    DefaultCategoryDataset bardataset = new DefaultCategoryDataset();
    //    DefaultCategoryDataset bardataset1 = new DefaultCategoryDataset();


    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
    //        {
    //            string strVal1 = Convert.ToString(ds.Tables[0].Rows[i]["Honore"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Honore"]);
    //            string strVal2 = Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]);
    //            Double s1val = Convert.ToDouble(strVal1, System.Globalization.CultureInfo.InvariantCulture);
    //            Double s2val = Convert.ToDouble(strVal2, System.Globalization.CultureInfo.InvariantCulture);

    //            sDate = ds.Tables[0].Rows[i]["Month"].ToString();

    //            bardataset.setValue(s1val * 100, "Marks", sDate);
    //            bardataset.setValue(s2val * 100, "Marks1", sDate);
    //        }
    //    }


    //    string strVal3 = Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["Honore"]);
    //    string strVal4 = Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index"]);
    //    Double s3val = Convert.ToDouble(strVal3, System.Globalization.CultureInfo.InvariantCulture);
    //    Double s4val = Convert.ToDouble(strVal4, System.Globalization.CultureInfo.InvariantCulture);

    //    sDate = ds.Tables[0].Rows[0]["MinDate"].ToString();

    //    bardataset1.setValue(s3val * 100, "Marks", sDate);
    //    bardataset1.setValue(s4val * 100, "Marks1", sDate);


    //    CategoryAxis rangeAxis1 = new CategoryAxis(" ");
    //    // rangeAxis1.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
    //    BarRenderer renderer1 = new BarRenderer();
    //    renderer1.setMaximumBarWidth(.10);

    //    renderer1.setDrawBarOutline(false);
    //    renderer1.setSeriesPaint(0, java.awt.Color.decode("#558ED5"));
    //    renderer1.setSeriesPaint(1, java.awt.Color.decode("#003399"));

    //    renderer1.setBaseToolTipGenerator(new StandardCategoryToolTipGenerator());
    //    CategoryPlot subplot1 = new CategoryPlot(bardataset, rangeAxis1, null, renderer1);
    //    subplot1.setDomainGridlinesVisible(true);

    //    subplot1.setDomainAxisLocation(AxisLocation.TOP_OR_LEFT);

    //    NumberAxis axis2 = (NumberAxis)subplot1.getRangeAxis();
    //    subplot1.setRangeAxis(1, axis2);

    //    subplot1.setBackgroundPaint(java.awt.Color.white);
    //    subplot1.setAxisOffset(new RectangleInsets(1.0, 1.0, 0.0, 0.0));
    //    subplot1.setDomainGridlinesVisible(false);
    //    subplot1.setRangeGridlinesVisible(false);
    //    subplot1.setOutlineStroke(new BasicStroke(0));
    //    subplot1.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));


    //    CategoryAxis rangeAxis2 = new CategoryAxis(" ");
    //    // rangeAxis2.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
    //    BarRenderer renderer2 = new BarRenderer();
    //    renderer2.setMaximumBarWidth(.35);
    //    renderer2.setBaseToolTipGenerator(new StandardCategoryToolTipGenerator());
    //    renderer2.setSeriesPaint(0, java.awt.Color.decode("#558ED5"));
    //    renderer2.setSeriesPaint(1, java.awt.Color.decode("#003399"));
    //    CategoryPlot subplot2 = new CategoryPlot(bardataset1, rangeAxis2, null, renderer2);

    //    subplot2.setBackgroundPaint(java.awt.Color.white);
    //    subplot2.setAxisOffset(new RectangleInsets(1.0, 0.0, 0.0, 1.0));
    //    subplot2.setDomainGridlinesVisible(false);
    //    subplot2.setRangeGridlinesVisible(false);
    //    subplot2.setOutlineStroke(new BasicStroke(0));
    //    // subplot2.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));


    //    subplot2.setOutlineStroke(new BasicStroke(0f));
    //    subplot2.setDomainGridlinesVisible(false);
    //    // subplot2.setDomainGridlinesVisible(true);
    //    subplot2.setDomainAxisLocation(AxisLocation.TOP_OR_LEFT);


    //    NumberAxis axis = (NumberAxis)subplot2.getRangeAxis();
    //    subplot2.setRangeAxis(1, axis);
    //    subplot2.setRangeAxisLocation(AxisLocation.BOTTOM_OR_RIGHT);

    //    //   CombinedDomainCategoryPlot plot = new CombinedDomainCategoryPlot();
    //    CombinedRangeCategoryPlot plot = new CombinedRangeCategoryPlot();



    //    // plot.setDataset(bardataset1);
    //    //  plot.setDataset(bardataset);
    //    plot.add(subplot1, 3);
    //    plot.add(subplot2, 1);
    //    plot.setOrientation(PlotOrientation.VERTICAL);


    //    NumberAxis axis1 = (NumberAxis)plot.getRangeAxis();
    //    plot.setRangeAxis(1, axis1);
    //    plot.setRangeAxisLocation(0, AxisLocation.BOTTOM_OR_LEFT);
    //    plot.setRangeAxisLocation(1, AxisLocation.BOTTOM_OR_RIGHT);
    //    plot.setGap(0.0);

    //    JFreeChart combinedchart = new JFreeChart("Combined Plot", JFreeChart.DEFAULT_TITLE_FONT, plot, false);



    //    combinedchart.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    combinedchart.getTitle().setPaint(java.awt.Color.BLACK);    // Set the colour of the title  
    //    combinedchart.setBackgroundPaint(java.awt.Color.white);    // Set the background colour of the chart  

    //    //CategoryPlot cp = barchart.getCategoryPlot();  // Get the Plot object for a bar graph  
    //    //cp.setBackgroundPaint(java.awt.Color.white);       // Set the plot background colour  
    //    //cp.setRangeGridlinePaint(java.awt.Color.gray);      // Set the colour of the plot gridlines  
    //    //cp.setDomainAxisLocation(AxisLocation.TOP_OR_LEFT);
    //    //cp.setRangeZeroBaselineVisible(true);
    //    //cp.setDomainZeroBaselineStroke(new java.awt.BasicStroke(10));
    //    //cp.setDomainZeroBaselinePaint(Color.GREEN);
    //    //cp.setRangeZeroBaselineStroke(new java.awt.BasicStroke(10));
    //    //cp.setRangeZeroBaselinePaint(Color.GREEN);

    //    //CategoryAxis categoryaxis = cp.getDomainAxis();
    //    // categoryaxis.setMaximumCategoryLabelWidthRatio(100);


    //    //NumberAxis rangeAxis = (NumberAxis)cp.getRangeAxis();
    //    //rangeAxis.setUpperMargin(0.2);
    //    //rangeAxis.setLowerMargin(0.2);
    //    //DecimalFormat pctFormat = new DecimalFormat("##0'%'");
    //    ////  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    //rangeAxis.setNumberFormatOverride(pctFormat);

    //    //NumberAxis axis2 = (NumberAxis)cp.getRangeAxis();
    //    // cp.setRangeAxis(1, axis2);
    //    //axis2.setUpperMargin(0.2);
    //    //axis2.setLowerMargin(0.2);
    //    //cp.setRangeAxisLocation(1, AxisLocation.BOTTOM_OR_RIGHT);
    //    //cp.mapDatasetToRangeAxis(1, 1);

    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    //  axis2.setNumberFormatOverride(pctFormat);

    //    //rangeAxis.setLowerBound(0);

    //    //Marker marker = new ValueMarker(-3);
    //    //marker.setOutlinePaint(java.awt.Color.red);
    //    //marker.setPaint(java.awt.Color.red);
    //    //marker.setStroke(new BasicStroke(5.0f));
    //    //cp.addRangeMarker(marker, Layer.FOREGROUND);

    //    //cp.addAnnotation(new CategoryLineAnnotation("Category 2", -5.0,
    //    //       "Category 4",  -2.0, java.awt.Color.red, new BasicStroke(2.0f)));


    //    //BarRenderer barrenderer = (BarRenderer)cp.getRenderer();
    //    //barrenderer.setDrawBarOutline(false);
    //    //barrenderer.setSeriesPaint(0, java.awt.Color.decode("#558ED5"));
    //    //barrenderer.setSeriesPaint(1, java.awt.Color.decode("#003399"));
    //    //barrenderer.setSeriesPaint(0, gradientpaint);

    //    //StandardCategoryItemLabelGenerator labelGen = new StandardCategoryItemLabelGenerator("{2}%", new DecimalFormat("0.0"));
    //    //barrenderer.setBaseItemLabelGenerator(labelGen);
    //    //barrenderer.setBaseItemLabelsVisible(true);


    //    //   DateAxis dateaxis = (DateAxis)cp.getDomainAxis();
    //    //axis.setLabelAngle(Math.PI / 2.0);
    //    //axis.setVerticalTickLabels(true);
    //    // dateaxis.setVerticalTickLabels(true);
    //    //  dateaxis.setDateFormatOverride(new SimpleDateFormat("yyyy"));

    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, combinedchart, 420, 300);

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();
    //}

    //private string GetChart5Left(out string mindate)
    //{

    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();
    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";

    //    // JFreeChart chart = ChartFactory.createBarChart(
    //    DB clsDB = new DB();

    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

    //    //DataSet ds = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA @GroupName = " + GrpName + ", @PositionGAFlagTxt = 'GA' , @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "', @AnnPerfFlg = 0 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

    //    DataSet ds = clsDB.getDataSet("Exec SP_R_WORST_MONTH_MAXDD_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + "");

    //    string Dt1;

    //    string sDate;

    //    DefaultCategoryDataset bardataset = new DefaultCategoryDataset();
    //    //DefaultCategoryDataset bardataset1 = new DefaultCategoryDataset();


    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
    //        {
    //            string strVal1 = Convert.ToString(ds.Tables[0].Rows[i]["Honore"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Honore"]);
    //            string strVal2 = Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]);
    //            Double s1val = Convert.ToDouble(strVal1, System.Globalization.CultureInfo.InvariantCulture);
    //            Double s2val = Convert.ToDouble(strVal2, System.Globalization.CultureInfo.InvariantCulture);

    //            string strVal3 = Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]);
    //            string strVal4 = Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]);
    //            Double s3val = Convert.ToDouble(strVal3, System.Globalization.CultureInfo.InvariantCulture);
    //            Double s4val = Convert.ToDouble(strVal4, System.Globalization.CultureInfo.InvariantCulture);

    //            //decimal ss = 12.5m;
    //            //ss = Math.Round(ss, 0, MidpointRounding.AwayFromZero);
    //            s1val = Math.Round(s1val * 100, 0, MidpointRounding.AwayFromZero);
    //            s2val = Math.Round(s2val * 100, 0, MidpointRounding.AwayFromZero);

    //            s3val = Math.Round(s3val * 100, 0, MidpointRounding.AwayFromZero);
    //            s4val = Math.Round(s4val * 100, 0, MidpointRounding.AwayFromZero);

    //            sDate = ds.Tables[0].Rows[i]["Month"].ToString();

    //            bardataset.setValue(s1val, "1", sDate);
    //            bardataset.setValue(s2val, "2", sDate);

    //            //To Set axis max-min value
    //            if (s1val <= min1)
    //                min1 = s1val;
    //            if (s2val <= min1)
    //                min1 = s2val;
    //            if (s3val <= min1)
    //                min1 = s3val;
    //            if (s4val <= min1)
    //                min1 = s4val;

    //            if (s1val >= max1)
    //                max1 = s1val;
    //            if (s2val >= max1)
    //                max1 = s2val;
    //            if (s3val >= max1)
    //                max1 = s1val;
    //            if (s4val >= max1)
    //                max1 = s4val;
    //        }
    //    }


    //    //double Max1 = 
    //    sDate = ds.Tables[0].Rows[0]["MinDate"].ToString();
    //    mindate = sDate;
    //    JFreeChart barchart = ChartFactory.createBarChart(
    //     "",      //Title  
    //     "",             // X-axis Label  
    //     "",               // Y-axis Label  
    //     bardataset,             // Dataset  
    //     PlotOrientation.VERTICAL,      //Plot orientation  
    //     false,                // Show legend  
    //     false,                // Use tooltips  
    //     false                // Generate URLs  
    //  );

    //    barchart.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    barchart.getTitle().setPaint(java.awt.Color.BLACK);    // Set the colour of the title  
    //    barchart.setBackgroundPaint(java.awt.Color.white);    // Set the background colour of the chart  

    //    CategoryPlot cp = barchart.getCategoryPlot();  // Get the Plot object for a bar graph  
    //    cp.setBackgroundPaint(java.awt.Color.white);       // Set the plot background colour  
    //    cp.setRangeGridlinePaint(java.awt.Color.white);      // Set the colour of the plot gridlines  
    //    cp.setDomainAxisLocation(AxisLocation.TOP_OR_LEFT);
    //    cp.setAxisOffset(new RectangleInsets(0, 0, 0, 0));
    //    cp.setDomainGridlinesVisible(false);
    //    cp.setRangeGridlinesVisible(false);
    //    cp.setOutlinePaint(java.awt.Color.white);

    //    CategoryAxis categoryaxis = cp.getDomainAxis();
    //    // categoryaxis.setMaximumCategoryLabelWidthRatio(100);
    //    //categoryaxis.setLowerMargin(0.0);
    //    //categoryaxis.setCategoryMargin(0.0);
    //    //categoryaxis.setUpperMargin(0.0);
    //    categoryaxis.setLabelFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 10));
    //    categoryaxis.setTickLabelFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 10));
    //    categoryaxis.setAxisLinePaint(java.awt.Color.decode("#868686")); //Axis color
    //    categoryaxis.setAxisLineStroke(new BasicStroke(1f));


    //    NumberAxis rangeAxis = (NumberAxis)cp.getRangeAxis();
    //    rangeAxis.setUpperMargin(0.2);
    //    rangeAxis.setLowerMargin(0.2);
    //    DecimalFormat pctFormat = new DecimalFormat("##0'%'");
    //    rangeAxis.setAxisLinePaint(java.awt.Color.decode("#868686"));//Axis color


    //    rangeAxis.setTickUnit(new NumberTickUnit(10));
    //    if (min1 < 0.0)
    //        min1 = min1 - 10.0;
    //    rangeAxis.setRange(min1, max1);

    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    rangeAxis.setNumberFormatOverride(pctFormat);

    //    //NumberAxis axis2 = (NumberAxis)cp.getRangeAxis();
    //    //cp.setRangeAxis(1, axis2);
    //    //axis2.setUpperMargin(0.2);
    //    //axis2.setLowerMargin(0.2);
    //    //cp.setRangeAxisLocation(1, AxisLocation.BOTTOM_OR_RIGHT);
    //    //cp.mapDatasetToRangeAxis(1, 1);

    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    //  axis2.setNumberFormatOverride(pctFormat);

    //    //rangeAxis.setLowerBound(0);

    //    BarRenderer barrenderer = (BarRenderer)cp.getRenderer();
    //    barrenderer.setDrawBarOutline(false);
    //    barrenderer.setSeriesPaint(0, java.awt.Color.decode("#558ED5"));
    //    barrenderer.setSeriesPaint(1, java.awt.Color.decode("#003399"));
    //    //barrenderer.setSeriesPaint(0, gradientpaint);
    //    barrenderer.setItemMargin(.1);
    //    // barrenderer.setItemMargin(0.0);

    //    // LineAndShapeRenderer renderer = (LineAndShapeRenderer)cp.getRenderer();
    //    barrenderer.setStroke(new BasicStroke(4f, 2, 2));

    //    DecimalFormat df = new DecimalFormat("0'%'");

    //    StandardCategoryItemLabelGenerator labelGen = new StandardCategoryItemLabelGenerator("{2}", df);
    //    barrenderer.setBaseItemLabelGenerator(labelGen);
    //    barrenderer.setBaseItemLabelsVisible(true);
    //    barchart.setPadding(new RectangleInsets(0, -10, 0, -10));

    //    barchart.setBorderVisible(false);
    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, barchart, 310, 250);




    //    return fsFinalLocation.Replace(".xls", ".png").ToString();
    //}

    //private string GetChart5Right()
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";

    //    // JFreeChart chart = ChartFactory.createBarChart(
    //    DB clsDB = new DB();

    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

    //    //DataSet ds = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA @GroupName = " + GrpName + ", @PositionGAFlagTxt = 'GA' , @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "', @AnnPerfFlg = 0 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

    //    DataSet ds = clsDB.getDataSet("Exec SP_R_WORST_MONTH_MAXDD_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + "");

    //    string Dt1;

    //    string sDate;

    //    DefaultCategoryDataset bardataset = new DefaultCategoryDataset();
    //    DefaultCategoryDataset bardataset1 = new DefaultCategoryDataset();


    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
    //        {
    //            string strVal1 = Convert.ToString(ds.Tables[0].Rows[i]["Honore"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Honore"]);
    //            string strVal2 = Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]);
    //            Double s1val = Convert.ToDouble(strVal1, System.Globalization.CultureInfo.InvariantCulture);
    //            Double s2val = Convert.ToDouble(strVal2, System.Globalization.CultureInfo.InvariantCulture);

    //            sDate = ds.Tables[0].Rows[i]["Month"].ToString();

    //            bardataset.setValue(s1val * 100, "Marks", sDate);
    //            bardataset.setValue(s2val * 100, "Marks1", sDate);
    //        }
    //    }


    //    string strVal3 = Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]);
    //    string strVal4 = Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]);
    //    Double s3val = Convert.ToDouble(strVal3, System.Globalization.CultureInfo.InvariantCulture);
    //    Double s4val = Convert.ToDouble(strVal4, System.Globalization.CultureInfo.InvariantCulture);

    //    s3val = Math.Round(s3val * 100, 0, MidpointRounding.AwayFromZero);
    //    s4val = Math.Round(s4val * 100, 0, MidpointRounding.AwayFromZero);

    //    sDate = ds.Tables[0].Rows[0]["MinDate"].ToString();

    //    bardataset1.setValue(s3val, "Marks", ".");
    //    bardataset1.setValue(s4val, "Marks1", ".");


    //    JFreeChart barchart = ChartFactory.createBarChart(
    //     "",      //Title  
    //     "",             // X-axis Label  
    //     "",               // Y-axis Label  
    //     bardataset1,             // Dataset  
    //     PlotOrientation.VERTICAL,      //Plot orientation  
    //     false,                // Show legend  
    //     false,                // Use tooltips  
    //     false                // Generate URLs  
    //  );

    //    barchart.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    barchart.getTitle().setPaint(java.awt.Color.BLACK);    // Set the colour of the title  
    //    barchart.setBackgroundPaint(java.awt.Color.white);    // Set the background colour of the chart  

    //    CategoryPlot cp = barchart.getCategoryPlot();  // Get the Plot object for a bar graph  
    //    cp.setBackgroundPaint(java.awt.Color.white);       // Set the plot background colour  
    //    cp.setRangeGridlinePaint(java.awt.Color.white);      // Set the colour of the plot gridlines  
    //    cp.setDomainGridlinesVisible(false);
    //    cp.setRangeGridlinesVisible(false);
    //    cp.setDomainAxisLocation(AxisLocation.TOP_OR_LEFT);
    //    cp.setRangeAxisLocation(AxisLocation.TOP_OR_RIGHT);
    //    cp.setAxisOffset(new RectangleInsets(0, 0, 0, 0));
    //    cp.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));

    //    CategoryAxis categoryaxis = cp.getDomainAxis();
    //    // categoryaxis.setMaximumCategoryLabelWidthRatio(100);
    //    // categoryaxis.setLabelPaint(java.awt.Color.white);
    //    //categoryaxis.setTickLabelsVisible(false);
    //    categoryaxis.setTickMarkPaint(java.awt.Color.white);
    //    categoryaxis.setTickMarkStroke(new BasicStroke(0.1f));
    //    categoryaxis.setAxisLinePaint(java.awt.Color.decode("#868686")); //Axis color
    //    categoryaxis.setAxisLineStroke(new BasicStroke(1f));
    //    categoryaxis.setCategoryMargin(0.1);

    //    NumberAxis rangeAxis = (NumberAxis)cp.getRangeAxis();
    //    rangeAxis.setUpperMargin(1);
    //    rangeAxis.setLowerMargin(1);
    //    DecimalFormat pctFormat = new DecimalFormat("##0'%'");
    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    rangeAxis.setNumberFormatOverride(pctFormat);
    //    rangeAxis.setTickUnit(new NumberTickUnit(10));
    //    rangeAxis.setRange(min1, max1);
    //    rangeAxis.setAxisLinePaint(java.awt.Color.decode("#868686")); //Axis color

    //    //NumberAxis axis2 = (NumberAxis)cp.getRangeAxis();
    //    //cp.setRangeAxis(1, axis2);
    //    //axis2.setUpperMargin(0.2);
    //    //axis2.setLowerMargin(0.2);
    //    //cp.setRangeAxisLocation(1, AxisLocation.BOTTOM_OR_RIGHT);
    //    //cp.mapDatasetToRangeAxis(1, 1);

    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    //  axis2.setNumberFormatOverride(pctFormat);

    //    //rangeAxis.setLowerBound(0);

    //    BarRenderer barrenderer = (BarRenderer)cp.getRenderer();
    //    barrenderer.setDrawBarOutline(false);
    //    barrenderer.setSeriesPaint(0, java.awt.Color.decode("#558ED5"));
    //    barrenderer.setSeriesPaint(1, java.awt.Color.decode("#003399"));
    //    //barrenderer.setSeriesPaint(0, gradientpaint);
    //    barrenderer.setItemMargin(.1);


    //    // LineAndShapeRenderer renderer = (LineAndShapeRenderer)cp.getRenderer();
    //    barrenderer.setStroke(new BasicStroke(4f, 2, 2));

    //    StandardCategoryItemLabelGenerator labelGen = new StandardCategoryItemLabelGenerator("{2}%", new DecimalFormat("0"));
    //    barrenderer.setBaseItemLabelGenerator(labelGen);
    //    barrenderer.setBaseItemLabelsVisible(true);
    //    barchart.setPadding(new RectangleInsets(0, -10, 0, -10));

    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, barchart, 100, 250);

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();
    //}

    //private string GetShapeChart4New(string ReportNum)
    //{
    //    System.Random rand = new System.Random();
    //    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";

    //    double Xmax = 0.0;
    //    double Ymax = 0.0;
    //    double axismax = 0.0;
    //    DB clsDB = new DB();
    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

    //    DataSet ds = clsDB.getDataSet("EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + ",@StartDate = '01-JAN-2011'");

    //    DataTable dt = GetFormatedDatatable(ds);


    //    XYSeriesCollection dataset = new XYSeriesCollection();

    //    XYSeries series1 = new XYSeries(dt.Rows[0]["Name"].ToString());
    //    XYSeries series2 = new XYSeries(dt.Rows[1]["Name"].ToString());
    //    XYSeries series3 = new XYSeries(dt.Rows[2]["Name"].ToString());
    //    XYSeries series4 = new XYSeries(dt.Rows[3]["Name"].ToString());

    //    Double s1X = Convert.ToDouble(dt.Rows[0]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s1Y = Convert.ToDouble(dt.Rows[0]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    Double s2X = Convert.ToDouble(dt.Rows[1]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s2Y = Convert.ToDouble(dt.Rows[1]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    Double s3X = Convert.ToDouble(dt.Rows[2]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s3Y = Convert.ToDouble(dt.Rows[2]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    Double s4X = Convert.ToDouble(dt.Rows[3]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
    //    Double s4Y = Convert.ToDouble(dt.Rows[3]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

    //    series1.add(s1X, s1Y);
    //    series2.add(s2X, s2Y);
    //    series3.add(s3X, s3Y);
    //    series4.add(s4X, s4Y);

    //    dataset.addSeries(series1);
    //    dataset.addSeries(series2);
    //    dataset.addSeries(series3);
    //    dataset.addSeries(series4);

    //    if (s1Y >= Ymax)
    //        Ymax = s1Y;
    //    if (s2Y >= Ymax)
    //        Ymax = s2Y;
    //    if (s3Y >= Ymax)
    //        Ymax = s3Y;
    //    if (s4Y >= Ymax)
    //        Ymax = s4Y;

    //    if (s1X >= Xmax)
    //        Xmax = s1X;
    //    if (s2X >= Xmax)
    //        Xmax = s2X;
    //    if (s3X >= Xmax)
    //        Xmax = s3X;
    //    if (s4X >= Xmax)
    //        Xmax = s4X;

    //    if (Xmax > Ymax)
    //        axismax = Xmax;
    //    else
    //        axismax = Ymax;

    //    JFreeChart chart = ChartFactory.createXYLineChart("Performance vs. Volatility (since 01/01/2011)", // chart title
    //                                                    "Annualized Volatility (Standard Deviation)", // domain axis label
    //                                                    "Annualized Return %", // range axis label
    //                                                    dataset, // data
    //                                                    PlotOrientation.VERTICAL, // orientation
    //                                                    false, // include legend
    //                                                    false, // tooltips
    //                                                    false // urls
    //                                                    );

    //    chart.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    if (ReportNum == "5")
    //    {
    //        DataSet ds1 = clsDB.getDataSet("EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + GrpName + ",@PositionGAFlagTxt = 'GA',@TrxnGAFlagTxt = 'GA',@AsOfDate = '" + txtAsofdate.Text + "',@BenchMarkName = null,@AssetNameTxt = " + strAssetClass + "");
    //        string strTitle = ds1.Tables[0].Rows[9]["ReturnName"].ToString();
    //        chart.setTitle(strTitle);
    //    }

    //    chart.setBackgroundPaint(java.awt.Color.white);
    //    chart.setBorderVisible(false);

    //    XYPlot plot = (XYPlot)chart.getPlot();
    //    plot.setBackgroundPaint(java.awt.Color.white);
    //    plot.setRangeGridlinePaint(java.awt.Color.white);
    //    plot.setAxisOffset(new RectangleInsets(0.0, 0.0, 0, 0.0));
    //    plot.setDomainGridlinesVisible(false);
    //    plot.setRangeGridlinesVisible(true);
    //    plot.setOutlineStroke(new BasicStroke(0));
    //    plot.setOutlinePaint(java.awt.Color.decode("#FFFFFF"));


    //    // customise the range axis...
    //    NumberAxis rangeAxis = (NumberAxis)plot.getRangeAxis();
    //    rangeAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());

    //    rangeAxis.setUpperMargin(0.2);
    //    rangeAxis.setLowerMargin(0.2);
    //    rangeAxis.setTickUnit(new NumberTickUnit(2));
    //    rangeAxis.setRange(0.0, axismax + 2);

    //    DecimalFormat pctFormat = new DecimalFormat("0'%'");
    //    rangeAxis.setNumberFormatOverride(pctFormat);

    //    ValueAxis valAxis1 = (ValueAxis)plot.getRangeAxis();
    //    DecimalFormat pctFormat1 = new DecimalFormat("0.00");
    //    //  CategoryAxis categoryaxis = plot.getDomainAxis();



    //    NumberAxis DomAxis = (NumberAxis)plot.getDomainAxis();
    //    DomAxis.setNumberFormatOverride(pctFormat);
    //    DomAxis.setLowerBound(0);
    //    DomAxis.setLowerMargin(0.5);
    //    DomAxis.setTickUnit(new NumberTickUnit(2));
    //    DomAxis.setRange(0.0, axismax + 2);
    //    java.awt.Font font3 = new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 11);
    //    plot.getDomainAxis().setLabelFont(font3);
    //    plot.getRangeAxis().setLabelFont(font3);



    //    // customise the renderer...
    //    XYLineAndShapeRenderer renderer = (XYLineAndShapeRenderer)plot.getRenderer();

    //    // Shape shape = new Rectangle2D.Double(-3.0, -3.0, 6.0, 6.0);
    //    Shape shape = new Ellipse2D.Double(-3.0, -3.0, 6.0, 6.0);


    //    renderer.setShapesVisible(true);
    //    renderer.setDrawOutlines(true);
    //    renderer.setUseFillPaint(true);
    //    renderer.setFillPaint(java.awt.Color.white);
    //    renderer.setSeriesShape(0, shape);
    //    renderer.setSeriesPaint(0, java.awt.Color.BLACK);

    //    XYItemLabelGenerator generator1 =
    //    new StandardXYItemLabelGenerator("{0}", new DecimalFormat("0.00"), new DecimalFormat("0.00"));
    //    renderer.setBaseItemLabelGenerator(generator1);

    //    renderer.setBasePositiveItemLabelPosition(new ItemLabelPosition(
    //               ItemLabelAnchor.OUTSIDE4, TextAnchor.TOP_CENTER));
    //    renderer.setBaseItemLabelsVisible(true);


    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, chart, 420, 350);

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();

    //}

    //private string getBarChartTAB3()
    //{
    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_11122.xls";

    //    // JFreeChart chart = ChartFactory.createBarChart(
    //    DB clsDB = new DB();

    //    string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
    //    string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

    //    string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

    //    DataSet ds = clsDB.getDataSet("exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA @GroupName = " + GrpName + ", @PositionGAFlagTxt = 'GA' , @TrxnGAFlagTxt = 'GA' ,@AsOfDate = '" + txtAsofdate.Text + "', @AnnPerfFlg = 0 , @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@AssetNameTxt = " + strAssetClass + ",@InclFixedIncome = 1");

    //    string Dt1;
    //    string sDay;
    //    string sMonth;
    //    string sYear;

    //    DefaultCategoryDataset bardataset = new DefaultCategoryDataset();


    //    if (ds.Tables[0].Rows.Count > 0)
    //    {
    //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
    //        {
    //            string strReturn = Convert.ToString(ds.Tables[0].Rows[i]["Return"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Return"]);
    //            Double s1val = Convert.ToDouble(strReturn, System.Globalization.CultureInfo.InvariantCulture);

    //            Dt1 = ds.Tables[0].Rows[i]["Dates"].ToString();
    //            sDay = "01";
    //            sMonth = "01";
    //            sYear = ds.Tables[0].Rows[i]["Dates"].ToString();

    //            //chart 1
    //            //   s1.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s1val);
    //            Day day = new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear));
    //            string d = Convert.ToInt16(sDay) + "/" + Convert.ToInt16(sMonth) + "/" + Convert.ToInt32(sYear);
    //            bardataset.setValue(s1val * 100, "Marks", sYear);
    //        }
    //    }


    //    //bardataset.setValue(6, "Marks", "Aditi");
    //    //bardataset.setValue(3, "Marks", "Pooja");
    //    //bardataset.setValue(10, "Marks", "Ria");
    //    //bardataset.setValue(5, "Marks", "Twinkle");
    //    //bardataset.setValue(20, "Marks", "Rutvi");

    //    JFreeChart barchart = ChartFactory.createBarChart(
    //     "Annual Performance of Gresham Advised Assets (GAA)",      //Title  
    //     "",             // X-axis Label  
    //     "",               // Y-axis Label  
    //     bardataset,             // Dataset  
    //     PlotOrientation.VERTICAL,      //Plot orientation  
    //     false,                // Show legend  
    //     false,                // Use tooltips  
    //     false                // Generate URLs  
    //  );

    //    barchart.getTitle().setFont(new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12));

    //    barchart.getTitle().setPaint(java.awt.Color.BLACK);    // Set the colour of the title  
    //    barchart.setBackgroundPaint(java.awt.Color.white);    // Set the background colour of the chart  
    //    CategoryPlot cp = barchart.getCategoryPlot();  // Get the Plot object for a bar graph  
    //    cp.setBackgroundPaint(java.awt.Color.white);       // Set the plot background colour  
    //    cp.setRangeGridlinePaint(java.awt.Color.gray);      // Set the colour of the plot gridlines  



    //    //cp.setRangeZeroBaselineVisible(true);
    //    //cp.setDomainZeroBaselineStroke(new java.awt.BasicStroke(10));
    //    //cp.setDomainZeroBaselinePaint(Color.GREEN);
    //    //cp.setRangeZeroBaselineStroke(new java.awt.BasicStroke(10));
    //    //cp.setRangeZeroBaselinePaint(Color.GREEN);

    //    CategoryAxis categoryaxis = cp.getDomainAxis();
    //    categoryaxis.setMaximumCategoryLabelWidthRatio(100);

    //    NumberAxis rangeAxis = (NumberAxis)cp.getRangeAxis();
    //    rangeAxis.setUpperMargin(0.2);
    //    rangeAxis.setLowerMargin(0.2);
    //    DecimalFormat pctFormat = new DecimalFormat("##0'%'");
    //    //  rangeAxis.setTickUnit(new NumberTickUnit(.1, new DecimalFormat("##0%")));
    //    rangeAxis.setNumberFormatOverride(pctFormat);

    //    //rangeAxis.setLowerBound(0);

    //    BarRenderer barrenderer = (BarRenderer)cp.getRenderer();
    //    barrenderer.setDrawBarOutline(false);
    //    barrenderer.setSeriesPaint(0, java.awt.Color.decode("#558ED5"));
    //    //barrenderer.setSeriesPaint(0, gradientpaint);

    //    StandardCategoryItemLabelGenerator labelGen = new StandardCategoryItemLabelGenerator("{2}%", new DecimalFormat("0.0"));
    //    barrenderer.setBaseItemLabelGenerator(labelGen);
    //    barrenderer.setBaseItemLabelsVisible(true);

    //    //   DateAxis dateaxis = (DateAxis)cp.getDomainAxis();
    //    //axis.setLabelAngle(Math.PI / 2.0);
    //    //axis.setVerticalTickLabels(true);
    //    // dateaxis.setVerticalTickLabels(true);
    //    //  dateaxis.setDateFormatOverride(new SimpleDateFormat("yyyy"));

    //    java.io.File file = new java.io.File(fsFinalLocation.Replace(".xls", ".png"));
    //    ChartUtilities.saveChartAsPNG(file, barchart, 860, 210);

    //    return fsFinalLocation.Replace(".xls", ".png").ToString();

    //}


    #endregion

    #region Fonts
    public iTextSharp.text.Font setFontsAll(float size, int bold, int italic, iTextSharp.text.Color foColor)
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
        string fontpath = HttpContext.Current.Server.MapPath(".");

        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanai.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //}
        //if (bold == 1 && italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\Fverdanaz.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        //}

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
    public iTextSharp.text.Font setFontsAll(float size, int bold, int italic)
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


    public iTextSharp.text.Font setFontsAllFrutiger(float size, int bold, int italic)
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
        string fontpath = HttpContext.Current.Server.MapPath(".");

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


    public iTextSharp.text.Font setFontsverdana()
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
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, 9, iTextSharp.text.Font.NORMAL);



        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    //FTI_____.PFM
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //}

        return font;
        #endregion
    }


    public iTextSharp.text.Font setFontsAll1(float size, int bold)
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
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);

            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }


        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);

        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}


        return font;
        #endregion
    }

    #endregion


    public string RoundUp(string lsFormatedString)
    {
        lsFormatedString = String.Format("{0:#,###0.00;(#,###0.00)}", Convert.ToDecimal(lsFormatedString));
        return lsFormatedString;
    }

    public string MoneyFormat(string MnyString)
    {
        if (MnyString != "")
        {
            if (MnyString.Contains("-"))
            {
                MnyString = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(MnyString));
                MnyString = MnyString.Replace("(", "($");
            }
            else
                MnyString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(MnyString));
        }
        else
        {
            MnyString = "$0";
        }

        return MnyString;
    }


}




