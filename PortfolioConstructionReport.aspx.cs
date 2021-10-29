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

using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.IO;
using System.Text;

using org.jfree.chart;
using org.jfree.chart.block;
using org.jfree.chart.plot;
using org.jfree.chart.title;
using Microsoft.IdentityModel.Claims;
using System.Threading;

public partial class PortfolioConstructionReport : System.Web.UI.Page
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
    protected void btnSubmit_Click(object sender, EventArgs e)
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

            //  String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";
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

            string filepdfname = objCombinedReports.MergeReports(fsFinalLocation, "Portfolio Construction Chart");

            FileInfo loFile = new FileInfo(filepdfname);

            loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

            //Response.Write("<script>");
            //  string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";         
            //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
            //Response.Write("</script>");
            
            Response.Write("<script>");           
            Response.Write("window.open('ViewReport.aspx?" + strGUID+".pdf" + "', 'mywindow')");
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

    private decimal RoundPercent(decimal Value)
    {
        Value = Convert.ToDecimal(String.Format("{0:#,###0.0;(#,###0.0)}", Value));
        return Value;
    }
    private decimal RoundValue(decimal Value)
    {
        Value = Convert.ToDecimal(String.Format("{0:#,###0;(#,###0)}", Value));
        return Value;
    }

    public string generatePDFFinal_New(DataSet newdataset)
    {
        liPageSize = 29;

        DB clsDB = new DB();

        String lsFooterTxt = String.Empty;

        DataTable table = newdataset.Tables[1];
        Random rnd = new Random();
        string strRndNumber = Convert.ToString(rnd.Next(5));
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + strRndNumber;

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        int topmargin = 25;//30
        if (Convert.ToBoolean(newdataset.Tables[3].Rows[0]["ShowTargetFlg"]) == true && FooterText != "")
            topmargin = 25;

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 50, 50, topmargin, 6);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        AddFooter(document, FooterText);

        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
                                         //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        //png.SetAbsolutePosition(40, 810);//540
        //png.ScalePercent(10);
        //document.Add(png);

        lsTotalNumberofColumns = 15 + "";

        iTextSharp.text.Table Table = new iTextSharp.text.Table(3, 2);
        iTextSharp.text.Cell cell = new Cell();

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(15, 8);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        setTableProperty(loTable);
        //setTableProperty1(Table);

        int liTotalPage = 1;// (newdataset.Tables[0].Rows.Count / liPageSize);
        int liCurrentPage = 0;
        liPageSize = 38;

        iTextSharp.text.Chunk lochunk = new Chunk();
        iTextSharp.text.Chunk chunk = new Chunk();

        if (table.Rows.Count > 0)
        {
            if (AsOfDate != "")
                lsDateName = Convert.ToDateTime(AsOfDate).ToString("MMMM dd, yyyy") + "";

            string Cash_UUid = "9776259D-0392-4DE0-8A12-0399724ABF8D";
            string Fixed_Income_UUid = "028B5EFB-D604-DE11-A38C-001D09665E8F";
            string Hedged_Strategies_UUid = "2287692A-D704-DE11-A38C-001D09665E8F";
            string Domestic_Equity_UUid = "E2A78BEB-D604-DE11-A38C-001D09665E8F";
            string International_Equity_UUid = "42B39247-D704-DE11-A38C-001D09665E8F";
            string Global_Opportunistic_UUid = "8413896B-4925-DF11-B686-001D09665E8F";
            string Liquid_Real_Assets_UUid = "0332530A-1AD3-DF11-9789-0019B9E7EE05";
            /***********************************************************************************/
            string PrivEqty_UUID = "02FFE912-D704-DE11-A38C-001D09665E8F";
            string ConHold_UUID = "E2465B5C-40A7-4A35-B5EA-50A2C74CF6F5";
            string Illiquid_Real_Assets_UUID = "C2A2D71C-D704-DE11-A38C-001D09665E8F";

            string ColName = "sas_assetClassID";

            string MarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 1))));
            string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2)))));
            string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 1))));
            string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal((GetFilteredValue(table, ConHold_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2)))));

            double MarketPerc = Convert.ToDouble(MarketablePercent);
            double nonMarketPerc = Convert.ToDouble(NonMarketablePercent);
            if (MarketPerc > 100.0)
                MarketablePercent = "100.0";
            if (nonMarketPerc > 100.0)
                NonMarketablePercent = "100.0";

            #region Data Values
            double ConcentratedHolding = 0;
            //string DashLines = "|\n|\n|";
            string DashLines = "";
            //string UpperDashLines = "|\n|";
            string UpperDashLines = "";
            for (int i = 0; i < 8; i++)
            {
                int colsize = 15;
                for (int j = 0; j < colsize; j++)
                {
                    string Text = "";
                    string lsfamilyName = "";

                    if (i == 0 && j == 0)
                    {
                        lochunk = new Chunk("\n" + lsFamiliesName.Replace("''", "'") + Text + "", setFontsAll(14, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.Leading = 10F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 1 && j == 0)
                    {
                        string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";
                        if (GreshamAdvisedFlag == "GA")
                        {
                            lochunk = new Chunk("GRESHAM ADVISED ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);

                            lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                            loCell.Add(lochunk);

                            lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                            loCell.Add(lochunk);
                        }
                        else
                        {
                            lochunk = new Chunk("TOTAL INVESTMENT ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);

                            lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                            loCell.Add(lochunk);

                            lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                            loCell.Add(lochunk);
                        }

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
                        loCell.Leading = 12F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 1)
                    {
                        string UUID = "9776259D-0392-4DE0-8A12-0399724ABF8D";
                        string Hdr1 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        lochunk = new Chunk(Hdr1 + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();

                        loCell.Colspan = 7;
                        j = j + 6;
                        loCell.Border = 0;
                        loCell.Add(lochunk);
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 7)
                    {
                        lochunk = new Chunk(UpperDashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 8)
                    {
                        string UUID = "E2465B5C-40A7-4A35-B5EA-50A2C74CF6F5";
                        string Hdr2 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        lochunk = new Chunk(Hdr2 + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Colspan = 5;
                        j = j + 4;
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 13)
                    {
                        lochunk = new Chunk(UpperDashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 14)
                    {
                        string UUID = "C2A2D71C-D704-DE11-A38C-001D09665E8F";
                        string Hdr3 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        lochunk = new Chunk(Hdr3 + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 0)
                    {
                        lochunk = new Chunk("Marketable Strategies\n\n" + MarketablePercent + Text + "%\n$" + MarketableValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 2)
                    {
                        string UUID = "9776259D-0392-4DE0-8A12-0399724ABF8D";
                        string Cash = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Cash + GetSeprator(Cash) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        //loCell.BorderWidthBottom = 2F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 4)
                    {
                        string UUID = "028B5EFB-D604-DE11-A38C-001D09665E8F";
                        string Fixed_Income = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Fixed_Income + GetSeprator(Fixed_Income) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 6) //Hedged Strategies Heading
                    {
                        string UUID = "2287692A-D704-DE11-A38C-001D09665E8F";
                        string Hedged_Strategies = GetFilteredValue(table, UUID, ColName);
                        string HedgedStr = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Hedged_Strategies + GetSeprator(Hedged_Strategies) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + HedgedStr + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 7)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 8)
                    {
                        string UUID = "E2A78BEB-D604-DE11-A38C-001D09665E8F";
                        string Domestic_Equity = GetFilteredValue(table, UUID, ColName);
                        string Domequity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Domestic_Equity + GetSeprator(Domestic_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Domequity + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 10) //International Equity Heading
                    {
                        string UUID = "42B39247-D704-DE11-A38C-001D09665E8F";
                        string International_Equity = GetFilteredValue(table, UUID, ColName);
                        string Intquity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + International_Equity + GetSeprator(International_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Intquity + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 12)
                    {
                        string UUID = "8413896B-4925-DF11-B686-001D09665E8F";
                        string Global_Opportunistic = GetFilteredValue(table, UUID, ColName);
                        string GlobOpp = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Global_Opportunistic + GetSeprator(Global_Opportunistic) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + GlobOpp + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        //SetBorder(loCell, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 13)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 14)
                    {
                        string UUID = "0332530A-1AD3-DF11-9789-0019B9E7EE05";
                        string Liquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Liquid_Real_Assets + GetSeprator(Liquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 0)
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.BorderWidthBottom = 1F;
                        loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 0)
                    {
                        lochunk = new Chunk("Non-Marketable Strategies\n" + NonMarketablePercent + Text + "%\n$" + NonMarketableValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 6 && j == 7)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 8)
                    {
                        string ConHoldUUID = "E2465B5C-40A7-4A35-B5EA-50A2C74CF6F5";
                        string PrivEqtyUUID = "02FFE912-D704-DE11-A38C-001D09665E8F";
                        ConcentratedHolding = Convert.ToDouble(GetFilteredValue(table, ConHoldUUID, ColName, 1));

                        if (ConcentratedHolding == 0)
                        {
                            //string Private_Equity = Convert.ToString(table.Rows[10][0]);
                            string Private_Equity = GetFilteredValue(table, PrivEqtyUUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, PrivEqtyUUID, ColName, 1)));
                            lochunk = new Chunk("" + Private_Equity + GetSeprator(Private_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, PrivEqtyUUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }
                        else
                        {
                            //string Concentrated_Holdings = Convert.ToString(table.Rows[1][0]);
                            string Concentrated_Holdings = GetFilteredValue(table, ConHoldUUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, ConHoldUUID, ColName, 1)));
                            lochunk = new Chunk("" + Concentrated_Holdings + GetSeprator(Concentrated_Holdings) + RoundUp(Convert.ToString(GetFilteredValue(table, ConHoldUUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }

                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 10)
                    {
                        if (ConcentratedHolding != 0)
                        {
                            string UUID = "02FFE912-D704-DE11-A38C-001D09665E8F";
                            string Private_Equity = GetFilteredValue(table, UUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                            lochunk = new Chunk(Private_Equity + GetSeprator(Private_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                            SetBorder(loCell, true, true, true, true);

                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loTable.AddCell(loCell);
                        }
                        else
                        {
                            lochunk = new Chunk("", setFontsAll(7, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);
                            loCell.Border = 0;
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loTable.AddCell(loCell);
                        }

                    }
                    else if (i == 6 && j == 13)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 14)
                    {
                        string UUID = "C2A2D71C-D704-DE11-A38C-001D09665E8F";
                        string Illiquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk(Illiquid_Real_Assets + GetSeprator(Illiquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else
                    {
                        lochunk = new Chunk(Text, setFontsAll(10, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);
                    }
                }
            }
            #endregion

            lsTotalNumberofColumns = 3 + "";
            setTableProperty(Table);

            string StarategicPurposeChart = generateOverAllPieChart(newdataset.Tables[2], "1");
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(StarategicPurposeChart);

            string VolatileChart = generateOverAllPieChart(newdataset.Tables[4], "2");
            iTextSharp.text.Image volatilejpg = iTextSharp.text.Image.GetInstance(VolatileChart);

            iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\dashbar.png");
            //document.Add(jpg); //add an image to the created pdf document

            chunk = new Chunk("\n\nStrategic Purpose", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
            cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);


            cell = new iTextSharp.text.Cell();
            cell.Add(GetCenterTable(newdataset.Tables[3]));
            Table.AddCell(cell);

            chunk = new Chunk("\n\n         Volatility Profile", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
            cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            jpg.Border = 0;
            jpg.ScaleToFit(220f, 400f);
            jpg.SetAbsolutePosition(45f, 90f);
            //jpg.Alignment = Image.TEXTWRAP | Image.ALIGN_RIGHT;
            jpg.IndentationLeft = 9f;
            jpg.SpacingAfter = 9f;
            //jpg.BorderWidthTop = 36f;
            //jpg.BorderColorTop = Color.WHITE;
            document.Add(jpg);

            volatilejpg.ScaleToFit(220f, 400f);
            volatilejpg.SetAbsolutePosition(590f, 90f);
            //jpg.Alignment = Image.TEXTWRAP | Image.ALIGN_RIGHT;
            volatilejpg.IndentationLeft = 9f;
            volatilejpg.SpacingAfter = 9f;
            //volatilejpg.BorderWidthTop = 36f;
            // volatilejpg.BorderColorTop = Color.WHITE;
            document.Add(volatilejpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(138f, 300f);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(421f, 300f);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(703f, 300f);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            /*
            chunk = new Chunk("Strategic Purpose \n\n\n", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            chunk = new Chunk("");
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);


            chunk = new Chunk("Volatile Profile", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            chunk = new Chunk("");
            cell = new iTextSharp.text.Cell();
            cell.Add(jpg);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            cell = new iTextSharp.text.Cell();
            cell.Add(GetCenterTable(newdataset.Tables[3]));
            // cell.Rowspan = 2;
            Table.AddCell(cell);

            cell = new iTextSharp.text.Cell();
            cell.Add(volatilejpg);
            cell.Border = 0;
            Table.AddCell(cell);
            */

            Paragraph pSpace = new Paragraph();
            pSpace.Add("\n");
            document.Add(loTable);
            //document.Add(pSpace);
            document.Add(Table);
            document.Close();

            try
            {
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }
            return fsFinalLocation.Replace(".xls", ".pdf");
        }
        else
        {
            //lblError.Text = "Record not found";
            return "Record not found";
        }
    }

    public void generatePDFFinal()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = String.Empty;
        //String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);
        String lsSQL = getFinalSp();
        // Response.Write(lsSQL);
        newdataset = clsDB.getDataSet(lsSQL);
        DataTable table = newdataset.Tables[1];
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 50, 50, 30, 8);//10,10
        String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        AddFooter(document);

        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
                                         //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        //png.SetAbsolutePosition(40, 810);//540
        //png.ScalePercent(10);
        //document.Add(png);

        double cr = Convert.ToDouble(table.Rows[6][1]);
        if (cr == 0)
        {
            lsTotalNumberofColumns = 12 + "";
        }
        else
        {
            lsTotalNumberofColumns = 13 + "";
        }
        lsTotalNumberofColumns = 15 + "";

        iTextSharp.text.Table Table = new iTextSharp.text.Table(3, 2);
        iTextSharp.text.Cell cell = new Cell();

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(15, 8);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        setTableProperty(loTable);
        //setTableProperty1(Table);

        int liTotalPage = 1;// (newdataset.Tables[0].Rows.Count / liPageSize);
        int liCurrentPage = 0;
        liPageSize = 38;

        iTextSharp.text.Chunk lochunk = new Chunk();
        iTextSharp.text.Chunk chunk = new Chunk();

        if (table.Rows.Count > 0)
        {
            if (txtAsofdate.Text != "")
                lsDateName = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";

            string MarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(table.Rows[0][1])) + RoundValue(Convert.ToDecimal(table.Rows[3][1])) + RoundValue(Convert.ToDecimal(table.Rows[5][1])) + RoundValue(Convert.ToDecimal(table.Rows[2][1])) + RoundValue(Convert.ToDecimal(table.Rows[7][1])) + RoundValue(Convert.ToDecimal(table.Rows[4][1])) + RoundValue(Convert.ToDecimal(table.Rows[8][1])));
            string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0][2])) + RoundPercent(Convert.ToDecimal(table.Rows[3][2])) + RoundPercent(Convert.ToDecimal(table.Rows[5][2]) + RoundPercent(Convert.ToDecimal(table.Rows[2][2])) + RoundPercent(Convert.ToDecimal(table.Rows[7][2])) + RoundPercent(Convert.ToDecimal(table.Rows[4][2])) + RoundPercent(Convert.ToDecimal(table.Rows[8][2]))));
            string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(table.Rows[1][1])) + RoundValue(Convert.ToDecimal(table.Rows[10][1])) + RoundValue(Convert.ToDecimal(table.Rows[6][1])));
            string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[1][2])) + RoundPercent(Convert.ToDecimal((table.Rows[10][2])) + RoundPercent(Convert.ToDecimal(table.Rows[6][2]))));

            #region Data Values
            double ConcentratedHolding = 0;
            //string DashLines = "|\n|\n|";
            string DashLines = "";
            //string UpperDashLines = "|\n|";
            string UpperDashLines = "";
            for (int i = 0; i < 8; i++)
            {
                int colsize = 15;
                for (int j = 0; j < colsize; j++)
                {
                    string Text = "";
                    string lsfamilyName = "";
                    if (ddlAllocationGroup.SelectedValue != "0")
                    {
                        lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
                    }
                    else
                        lsfamilyName = ddlHousehold.SelectedItem.Text;
                    if (i == 0 && j == 0)
                    {
                        lochunk = new Chunk("\n" + lsfamilyName + Text + "", setFontsAll(14, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.Leading = 10F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 1 && j == 0)
                    {
                        lochunk = new Chunk("PORTFOLIO CONSTRUCTION" + Text + "", setFontsAll(12, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        loCell.Add(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
                        loCell.Leading = 11F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 0)
                    {
                        string Hdr1 = Convert.ToString(table.Rows[0]["Strategic Purpose"]);
                        lochunk = new Chunk(Hdr1 + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();

                        loCell.Colspan = 8;
                        j = j + 7;
                        loCell.Border = 0;
                        loCell.Add(lochunk);
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 7)
                    {
                        lochunk = new Chunk(UpperDashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 8)
                    {
                        string Hdr2 = Convert.ToString(table.Rows[2]["Strategic Purpose"]);
                        lochunk = new Chunk(Hdr2 + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Colspan = 5;
                        j = j + 4;
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 13)
                    {
                        lochunk = new Chunk(UpperDashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 14)
                    {
                        string Hdr3 = Convert.ToString(table.Rows[8]["Strategic Purpose"]);
                        lochunk = new Chunk(Hdr3 + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 0)
                    {
                        lochunk = new Chunk("Marketable Strategies\n\n" + MarketablePercent + Text + "%\n$" + MarketableValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 2)
                    {
                        string Cash = Convert.ToString(table.Rows[0][0]);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[0][1]));
                        lochunk = new Chunk("" + Cash + "\n\n" + RoundUp(Convert.ToString(table.Rows[0][2])) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        //loCell.BorderWidthBottom = 2F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 4)
                    {
                        string Fixed_Income = Convert.ToString(table.Rows[3][0]);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[3][1]));
                        lochunk = new Chunk("" + Fixed_Income + "\n\n" + RoundUp(Convert.ToString(table.Rows[3][2])) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 6) //Hedged Strategies Heading
                    {
                        string Hedged_Strategies = Convert.ToString(table.Rows[5][0]);
                        string HedgedStr = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[5][1]));
                        lochunk = new Chunk("" + Hedged_Strategies + "\n\n" + RoundUp(Convert.ToString(table.Rows[5][2])) + Text + "%\n$" + HedgedStr + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 7)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 8)
                    {
                        string Domestic_Equity = Convert.ToString(table.Rows[2][0]);
                        string Domequity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[2][1]));
                        lochunk = new Chunk("" + Domestic_Equity + "\n\n" + RoundUp(Convert.ToString(table.Rows[2][2])) + Text + "%\n$" + Domequity + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 10) //International Equity Heading
                    {
                        string International_Equity = Convert.ToString(table.Rows[7][0]);
                        string Intquity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[7][1]));
                        lochunk = new Chunk("" + International_Equity + "\n\n" + RoundUp(Convert.ToString(table.Rows[7][2])) + Text + "%\n$" + Intquity + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 12)
                    {
                        string Global_Opportunistic = Convert.ToString(table.Rows[4][0]);
                        string GlobOpp = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[4][1]));
                        lochunk = new Chunk("" + Global_Opportunistic + "\n\n" + RoundUp(Convert.ToString(table.Rows[4][2])) + Text + "%\n$" + GlobOpp + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        //SetBorder(loCell, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 13)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 14)
                    {
                        string Liquid_Real_Assets = Convert.ToString(table.Rows[8][0]);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[8][1]));
                        lochunk = new Chunk("" + Liquid_Real_Assets + "\n\n" + RoundUp(Convert.ToString(table.Rows[8][2])) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 0)
                    {
                        lochunk = new Chunk("");
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.BorderWidthBottom = 1F;
                        loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 0)
                    {
                        lochunk = new Chunk("Non-Marketable Strategies\n" + NonMarketablePercent + Text + "%\n$" + NonMarketableValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 6 && j == 7)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 8)
                    {
                        ConcentratedHolding = Convert.ToDouble(table.Rows[1][1]);

                        if (ConcentratedHolding == 0)
                        {
                            string Private_Equity = Convert.ToString(table.Rows[10][0]);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[10][1]));
                            lochunk = new Chunk("" + Private_Equity + "\n\n" + RoundUp(Convert.ToString(table.Rows[10][2])) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }
                        else
                        {
                            string Concentrated_Holdings = Convert.ToString(table.Rows[1][0]);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[1][1]));
                            lochunk = new Chunk("" + Concentrated_Holdings + "\n\n" + RoundUp(Convert.ToString(table.Rows[1][2])) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }

                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 10)
                    {
                        if (ConcentratedHolding != 0)
                        {
                            string Private_Equity = Convert.ToString(table.Rows[10][0]);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[10][1]));
                            lochunk = new Chunk("Private Equity\n\n" + RoundUp(Convert.ToString(table.Rows[10][2])) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                            SetBorder(loCell, true, true, true, true);

                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loTable.AddCell(loCell);
                        }
                        else
                        {
                            lochunk = new Chunk("", setFontsAll(7, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);
                            loCell.Border = 0;
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loTable.AddCell(loCell);
                        }

                    }
                    else if (i == 6 && j == 13)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 10f;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 14)
                    {
                        string Illiquid_Real_Assets = Convert.ToString(table.Rows[6][0]);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[6][1]));
                        lochunk = new Chunk("Illiquid Real Assets\n\n" + RoundUp(Convert.ToString(table.Rows[6][2])) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else
                    {
                        lochunk = new Chunk(Text, setFontsAll(10, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);
                    }
                }
            }
            #endregion

            lsTotalNumberofColumns = 3 + "";
            setTableProperty(Table);

            string StarategicPurposeChart = generateOverAllPieChart(newdataset.Tables[2], "1");
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(StarategicPurposeChart);

            string VolatileChart = generateOverAllPieChart(newdataset.Tables[4], "2");
            iTextSharp.text.Image volatilejpg = iTextSharp.text.Image.GetInstance(VolatileChart);

            iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\dashbar.png");
            //document.Add(jpg); //add an image to the created pdf document

            chunk = new Chunk("\n\nStrategic Purpose", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
            cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);


            cell = new iTextSharp.text.Cell();
            cell.Add(GetCenterTable(newdataset.Tables[3]));
            Table.AddCell(cell);

            chunk = new Chunk("\n\nVolatility Profile", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
            cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            jpg.Border = 0;
            jpg.ScaleToFit(220f, 400f);
            jpg.SetAbsolutePosition(45f, 100f);
            //jpg.Alignment = Image.TEXTWRAP | Image.ALIGN_RIGHT;
            jpg.IndentationLeft = 9f;
            jpg.SpacingAfter = 9f;
            //jpg.BorderWidthTop = 36f;
            //jpg.BorderColorTop = Color.WHITE;
            document.Add(jpg);

            volatilejpg.ScaleToFit(220f, 400f);
            volatilejpg.SetAbsolutePosition(590f, 100f);
            //jpg.Alignment = Image.TEXTWRAP | Image.ALIGN_RIGHT;
            volatilejpg.IndentationLeft = 9f;
            volatilejpg.SpacingAfter = 9f;
            //volatilejpg.BorderWidthTop = 36f;
            // volatilejpg.BorderColorTop = Color.WHITE;
            document.Add(volatilejpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(421f, 308f);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(703f, 308f);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            /*
            chunk = new Chunk("Strategic Purpose \n\n\n", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            chunk = new Chunk("");
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);


            chunk = new Chunk("Volatile Profile", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            chunk = new Chunk("");
            cell = new iTextSharp.text.Cell();
            cell.Add(jpg);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);

            cell = new iTextSharp.text.Cell();
            cell.Add(GetCenterTable(newdataset.Tables[3]));
            // cell.Rowspan = 2;
            Table.AddCell(cell);

            cell = new iTextSharp.text.Cell();
            cell.Add(volatilejpg);
            cell.Border = 0;
            Table.AddCell(cell);
            */

        }
        else
        {
            lblError.Text = "Record not found";
            return;
        }

        Paragraph pSpace = new Paragraph();
        pSpace.Add("\n");
        document.Add(loTable);
        //document.Add(pSpace);
        document.Add(Table);

        if (newdataset.Tables[0].Rows.Count > 0)
        {
            document.Close();

            FileInfo loFile = new FileInfo(ls);
            try
            {
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

                Response.Write("<script>");
                string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
                Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                Response.Write("</script>");

            }
            catch (Exception exc)
            {
                Response.Write(exc.Message);
            }
        }
    }

    public iTextSharp.text.Table GetCenterTable(DataTable dt)
    {
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(5, 18);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        iTextSharp.text.Chunk lochunk = new Chunk();

        lsTotalNumberofColumns = 5 + "";

        setTableProperty(loTable);

        int rowsize = dt.Rows.Count + 1;
        for (int i = 0; i < rowsize; i++)
        {
            for (int j = 0; j < 5; j++)
            {
                if (i == 0 && j == 0)
                {
                    lochunk = new Chunk("", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 11F;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 1)
                {
                    lochunk = new Chunk("Current Allocation", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 11F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 2)
                {
                    lochunk = new Chunk("Current Allocation (%)", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 11F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 3)
                {
                    lochunk = new Chunk("Target Allocation (%)", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 11F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 4)
                {
                    lochunk = new Chunk("Volatility Profile", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 11F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i != 0)
                {
                    string ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                    string ColUUID = Convert.ToString(dt.Rows[i - 1]["StrategicUID"]);
                    if (j == 0)
                    {
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);

                    }
                    else if (j == 1)
                    {

                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                        if (ColValue != "")
                            ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                    }
                    else if (j == 2)
                    {
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                        if (ColValue != "")
                            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(ColValue));
                    }
                    else if (j == 3)
                    {
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                        if (ColValue != "")
                            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(ColValue));
                    }
                    else if (j == 4)
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);

                    if (Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)
                        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    else
                        lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));

                    loCell = new iTextSharp.text.Cell();
                    loCell.Add(lochunk);
                    loCell.Leading = 4F;

                    if (ColUUID.ToLower() == "fffb5207-6075-e211-aa29-0019b9e7ee05" && Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)//Risk Reduction Strategies
                    {
                        loCell.Leading = 4F;
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                    }
                    else if (ColUUID.ToLower() == "277c510e-6075-e211-aa29-0019b9e7ee05" && Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)//Growth Strategies
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                    else if (ColUUID.ToLower() == "07f0821c-6075-e211-aa29-0019b9e7ee05" && Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)//Economic Hedges
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));

                    loCell.Border = 0;
                    if (i == rowsize - 1)
                    {
                        loCell.BorderWidthTop = 0.5F;
                        loCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    }
                    if (j == 0)
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;

                    if ((j == 1 || j == 2 || j == 3) && Convert.ToString(dt.Rows[i - 1]["TotalFlg"]) == "True")
                    {
                        loCell.BorderWidthTop = 0.5F;
                        loCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    }
                    loTable.AddCell(loCell);
                }
            }
        }

        if (ddlAllocationGroup.SelectedValue == "0")
            loTable.DeleteColumn(3);

        return loTable;
    }

    public string generateOverAllPieChart(DataTable table, string ChartType)
    {
        /*  Chart Type
         * Strategic Purpose == 1;
         * Volatile Profile == 2; */

        string Header = string.Empty;
        if (ChartType == "1")
            Header = "";//"Strategic Purpose"
        else
            Header = "";//Volatility Profile
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";

        AppDomain.CurrentDomain.Load("JCommon");
        org.jfree.data.general.DefaultPieDataset myDataSet = new org.jfree.data.general.DefaultPieDataset();
        if (table.Rows.Count > 0)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (Convert.ToString(table.Rows[i][2]) != "")
                {
                    string ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(table.Rows[i][2]));
                    //myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[i][2]), 1)));
                    //if (ChartType == "1")
                    //    myDataSet.setValue(Convert.ToString(table.Rows[i][3]), Convert.ToDouble(ColValue));
                    //else
                    myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(ColValue));
                }
            }
        }
        JFreeChart pieChart = ChartFactory.createPieChart("", myDataSet, false, true, false);

        pieChart.setBackgroundPaint(java.awt.Color.white);
        pieChart.setBorderVisible(false);

        pieChart.setTitle(new org.jfree.chart.title.TextTitle(Header, new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 18)));

        PiePlot ColorConfigurator = (PiePlot)pieChart.getPlot();
        ColorConfigurator.setLabelBackgroundPaint(System.Drawing.Color.White);// ColorConfigurator.getLabelPaint()
        ColorConfigurator.setLabelOutlinePaint(System.Drawing.Color.White);
        ColorConfigurator.setLabelShadowPaint(System.Drawing.Color.White);
        //ColorConfigurator.setBackgroundPaint(System.Drawing.Color.White);
        ColorConfigurator.setOutlinePaint(System.Drawing.Color.White);
        //ColorConfigurator.setOutlineStroke setOutlinePaint(System.Drawing.Brush.);
        ColorConfigurator.setLabelFont(new System.Drawing.Font("Frutiger-Roman", 12));//, System.Drawing.FontStyle.Bold

        ColorConfigurator.setCircular(false);
        //ColorConfigurator.setOutlineVisible(false);  fix to remove border
        ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardPieSectionLabelGenerator("{0} {1}%"));
        //ColorConfigurator.setLabelGenerator(null);
        //ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardCategoryItemLabelGenerator("{0}"));
        //ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardCategoryItemLabelGenerator("{0} =  {1}%"));
        //ColorConfigurator.setInteriorGap(0.30);

        ColorConfigurator.setInteriorGap(0.15);
        ColorConfigurator.setLabelGap(0);

        java.util.List keys = myDataSet.getKeys();

        for (int i = 0; i < keys.size(); i++)
        {
            //if (keys.get(i).ToString() == "Risk Reduction Strategies")
            //    ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
            //if (keys.get(i).ToString() == "Growth Strategies")
            //    ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
            //if (keys.get(i).ToString() == "Economic Hedges")
            //    ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
            //if (keys.get(i).ToString() == "Cash")
            //    ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#D9D9D9"));
            //if (keys.get(i).ToString() == "High")
            //    ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#C4BD97"));
            //if (keys.get(i).ToString() == "Low")
            //    ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#E6B9B8"));
            //if (keys.get(i).ToString() == "Moderate")
            //    ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#FAC090"));

            if (keys.get(i).ToString().Contains("Risk"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
            if (keys.get(i).ToString().Contains("Growth"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
            if (keys.get(i).ToString().Contains("Economic"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
            if (keys.get(i).ToString().Contains("Cash"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#D9D9D9"));
            if (keys.get(i).ToString().Contains("High"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#C4BD97"));
            if (keys.get(i).ToString().Contains("Low"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#E6B9B8"));
            if (keys.get(i).ToString().Contains("Moderate"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#FAC090"));
        }
        ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
        java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
        ChartUtilities.writeChartAsPNG(jos, pieChart, 480, 330);

        return fsFinalLocation.Replace(".xls", ".png");
    }

    private string GetFilteredValue(DataTable dt, string Value, string QryColumnName)
    {
        return GetFilteredValue(dt, Value, QryColumnName, 0);
    }

    private string GetFilteredValue(DataTable dt, string Value, string QryColumnName, int GetDataFromColumn)
    {
        string retVal = "";
        DataView dv = new DataView(dt);
        dv.RowFilter = "" + QryColumnName + " = '" + Value + "'";
        dt = dv.ToTable();
        retVal = Convert.ToString(dt.Rows[0][GetDataFromColumn]);
        return retVal;
    }
    private string GetSeprator(string strText)
    {
        string Seprater = "\n\n";
        int length = strText.Length;
        if (length > 21)
            Seprater = "\n";
        return Seprater;
    }

    public void AddFooter(iTextSharp.text.Document document, string strText)
    {
        strText = "Gresham Advisors, LLC | 333 W. Wacker Dr. Suite 700 | Chicago, IL 60606 | P 312.960.0200 | F 312.960.0204 | www.greshampartners.com";
        Phrase footPhraseImg = new Phrase(strText, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
        footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
        footer.Alignment = Element.ALIGN_LEFT;
        document.Footer = footer;
    }

    private string getFinalSp()
    {
        // private string getFinalSp(ReportType Type, string ddlHouseholdValue, string ddlHouseholdText, string txtAsofdate, string ddlCashValue, string ddlReportFlgValue, string ddlAllocationGroupValue, string ddlAllocationGroupText, string ddlReport1and2Value, string ddlAllAssetValue)
        string lsSQL = "";
        string houseHold = "";
        string AsOfDates = txtAsofdate.Text.Trim() == "" ? "null" : "'" + txtAsofdate.Text.Trim() + "'";
        object HouseHold = ddlHousehold.SelectedValue == "0" ? "null" : "'" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "'";
        object AllocationGroup = ddlAllocationGroup.SelectedValue == "0" ? "null" : "'" + ddlAllocationGroup.SelectedItem.Text.Replace("'", "''") + "'";
        //string LegalEntity = LegalEntityId == "" ? "null" : "'" + LegalEntityId + "'";
        //string Fund = FundId == "" ? "null" : "'" + FundId + "'";
        string ReportRollUpGrp = ddlReportRollupgrp.SelectedValue == "0" ? "null" : ddlReportRollupgrp.SelectedValue;
        string GreshamAdvFlg = "'" + ddlGreshamAdvisedFlg.SelectedItem.Text + "'";

        if (Request.QueryString.Count > 0)
        {
            if (Request.QueryString["type"] == "new")
                lsSQL = "SP_R_CONSTRUCTIONCHART_EXCEL_NEW_GA @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + ",@GreshamAdvisedFlagTxt = '" + GreshamAdvFlg + "',@Reportrollupgroupid=" + ReportRollUpGrp + "";
        }
        else
            lsSQL = "SP_R_CONSTRUCTIONCHART_EXCEL @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + ",@GreshamAdvisedFlagTxt = 'TIA'";//old report
        return lsSQL;

    }

    #region Common Methods

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
    private string RoundUp(string lsFormatedString)
    {
        lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
        return lsFormatedString;
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
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        }
        if (italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanai.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\Fverdanaz.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        }

        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //}
        //if (bold == 1 && italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        //}
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


    public iTextSharp.text.Font setFontsAllFrutiger(int size, int bold, int italic)
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

    public void setTableProperty1(iTextSharp.text.Table fotable)
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
                int[] headerwidths2 = { 70, 18 };
                fotable.SetWidths(headerwidths2);
                fotable.Width = 88;
                break;
            case "3":
                int[] headerwidths3 = { 41, 60, 40 };
                fotable.SetWidths(headerwidths3);
                fotable.Width = 100;
                break;
            case "4":
                int[] headerwidths4 = { 45, 5, 15, 15 };
                fotable.SetWidths(headerwidths4);
                fotable.Width = 80;
                break;
            case "5":
                int[] headerwidths5 = { 22, 12, 12, 12, 12 };
                fotable.SetWidths(headerwidths5);
                fotable.Width = 100;
                break;
            case "6":
                int[] headerwidths6 = { 27, 11, 11, 8, 5, 7 };
                fotable.SetWidths(headerwidths6);
                fotable.Width = 70;
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
                int[] headerwidths13 = { 12, 2, 15, 2, 20, 20, 2, 15, 5, 15, 15, 2, 15 };
                fotable.SetWidths(headerwidths13);
                fotable.Width = 100; break;
            case "14":
                int[] headerwidths14 = { 30, 9 };
                fotable.SetWidths(headerwidths14);
                fotable.Width = 100;
                break;
            case "15":
                int[] headerwidths15 = { 15, 2, 15, 2, 15, 2, 15, 2, 15, 2, 15, 2, 15, 2, 15 };
                fotable.SetWidths(headerwidths15);
                fotable.Width = 100; break;
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


    public iTextSharp.text.Font Font8Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(8, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(8, 0, 0);
    }

    public iTextSharp.text.Font Font7GreyItalic()
    {
        return setFontsAll(7, 0, 1, new iTextSharp.text.Color(216, 216, 216));
    }




    public void AddFooter(iTextSharp.text.Document document)
    {
        Phrase footPhraseImg = new Phrase("Gresham Advisors, LLC | 333 W. Wacker Dr. Suite 700 | Chicago, IL 60606 | P 312.960.0200 | F 312.960.0204 | www.greshampartners.com", setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
        footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
        footer.Alignment = Element.ALIGN_CENTER;
        document.Footer = footer;
    }

    public void AddHeader(iTextSharp.text.Document document)
    {
        Paragraph pimage = new Paragraph();
        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png"));
        SignatureJpg.SetAbsolutePosition(45, 557);//540
                                                  //SignatureJpg.ScaleToFit(45, 557);
        SignatureJpg.ScalePercent(10);
        pimage.Add(SignatureJpg);

        //Phrase footPhraseImg = new Phrase(, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter Header = new HeaderFooter(pimage, false);
        Header.Border = iTextSharp.text.Rectangle.NO_BORDER;
        Header.Alignment = Element.ALIGN_CENTER;
        document.Header = Header;
    }

    public string RemoveStyle(string html)
    {
        //start by completely removing all unwanted tags 
        //System.Text.RegularExpressions.Regex.Replace(html, @"<(.|\n)*?>", string.Empty);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"<[/]?(font|span|xml|del|ins|[ovwxp]:\w+)[^>]*?>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //then run another pass over the html (twice), removing unwanted attributes 
        html = System.Text.RegularExpressions.Regex.Replace(html, @"(mso-bidi-font-style: normal)*<( )*font-family: Verdana( )*(/)*>", "\r", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"(mso-bidi-font-weight: normal)*<( )*font-family: Verdana( )*>", "\r", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        html = System.Text.RegularExpressions.Regex.Replace(html, @"<([^>]*)(?:class|[ovwxp]:\w+)=(?:'[^']*'|""[^""]*""|[^>]+)([^>]*)>", "<$1$2>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"<([^>]*)(?:class|[ovwxp]:\w+)=(?:'[^']*'|""[^""]*""|[^>]+)([^>]*)>", "<$1$2>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //html.Replace("<p", "\n").Replace("</p>", "\n").Replace("<li", "\n").Replace("</li>", "\n").Replace("b", "@").Replace("</b>", "@").Replace("<i", "@").Replace("</i>", "@");
        return html;
    }

    //public void AddStyle(string selector, string styles)
    //{

    //    string STYLE_DEFAULT_TYPE = "style";
    //    this._Styles.LoadTagStyle(selector, STYLE_DEFAULT_TYPE, styles);
    //}

    #endregion


}
