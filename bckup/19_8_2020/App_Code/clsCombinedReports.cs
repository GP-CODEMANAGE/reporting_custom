
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
using System.Globalization;
using System.Web.UI.DataVisualization.Charting;
using System.Linq;
using System.Collections.Generic;

/// <summary>
/// Summary description for clsCombinedReports
/// </summary>
public class clsCombinedReports
{
    #region General Declaration
    //public string lsSQL = "";
    Boolean fbCheckExcel = false;
    public StreamWriter sw = null;
    public string strDescription = string.Empty;
    //bool bProceed = true;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName;
    DB clsDB = new DB();
    string IncDate = "";
    string ColorTIA1 = "#558ED5"; //Blue
    string ColorNetInvestedCap = "#77933C"; //Green
    string ColorTIA2 = "#558ED5"; //Blue
    string ColorInflationAdjInvCap = "#E46C0A";//Orange
    double extendedrangePerc = 1.05;
    double max1 = 0.0;
    double min1 = 0.0;
    Logs lg = new Logs();
    public enum ReportType
    {
        PortfolioConsChart = 1,
        CommitmentSchedule = 2,
        AllocationGroupPieChart = 3,
        InvestmentObjectiveChart = 4,
        OverallPieChart = 5,
        DirectMgrDetail = 6,
        PortfolioConsChartNEW = 7,
        PerfAnalytics = 8,
        Rpt1LineChart = 9,
        Rpt1Table = 10,
        Rpt3LineChart = 11,
        Rpt3BarChart = 12,
        Rpt3Table1 = 13,
        Rpt3Table2 = 14,
        Rpt4ColumnChart = 15,
        Rpt4ShapeChart = 16,
        Rpt4TableLT = 17,
        Rpt4TableST = 18,
        RptMarketablePerf = 19,
        RptAssetsPerformanceSummary = 20,
        PortfolioConsChartV2 = 21,
        RptNonMarketablePerf = 22,	  //added 2_1_2019 Non Marketable (DYNAMO)
    }
    #endregion

    #region Properties
    private string HouseHold_Value = string.Empty;
    private string HouseHold_Text = string.Empty;
    private string AllocationGroup_Value = string.Empty;
    private string AllocationGroup_Text = string.Empty;
    private string AsOf_Date = string.Empty;
    private string lsFamilies_Name = string.Empty;
    private string lsDate_Name = string.Empty;
    private string _LeagalEntity = string.Empty;
    private string _Fund = string.Empty;
    private string _FooterText = string.Empty;
    private string _Footerlocation = string.Empty;

    private string _ClientFooterTxt = string.Empty;
    private string _Ssi_GreshamClientFooter = string.Empty;

    private string _CommitmentReportHeader = string.Empty;
    private string ReportRollUpGroup_Value = string.Empty;
    private string _GreshamAdvisedFlag = string.Empty;
    private string _PorthFolioConChartReportVer = string.Empty;
    private string _Chart1 = string.Empty;
    private string _Chart2 = string.Empty;
    private string _AssetClassCSV = string.Empty;
    private string _ReportRollupGroupIdName = string.Empty;
    private string _ReportingName = string.Empty;
    private string _ReportSource = string.Empty;
    private string _PriorDate = string.Empty;

    //added 2_1_2019 Non Marketable (DYNAMO)
    private string _ReportRollupGroupId = string.Empty;
    private string _HouseholdId = string.Empty;
    private string _FundIRR = string.Empty;
    private string _HHParameterTxt = string.Empty;
    private string _ReportingID = string.Empty;
    private string _ReportName = string.Empty;

    //added 8_14_2019 batch Issue(Mixing of Reports)
    private string _TempFolderPath = string.Empty;

    private string _LogFileName = string.Empty;
    public string LogFileName
    {
        get
        {
            return _LogFileName;
        }
        set
        {
            _LogFileName = value;
        }
    }
    public string CommitmentReportHeader
    {
        get
        {
            return _CommitmentReportHeader;
        }
        set
        {
            _CommitmentReportHeader = value;
        }
    }

    public string FooterText
    {
        get
        {
            return _FooterText;
        }
        set
        {
            _FooterText = value;
        }
    }

    public string Footerlocation
    {
        get
        {
            return _Footerlocation;
        }
        set
        {
            _Footerlocation = value;
        }
    }


    public string ClientFooterTxt
    {
        get
        {
            return _ClientFooterTxt;
        }
        set
        {
            _ClientFooterTxt = value;
        }
    }

    public string Ssi_GreshamClientFooter
    {
        get
        {
            return _Ssi_GreshamClientFooter;
        }
        set
        {
            _Ssi_GreshamClientFooter = value;
        }
    }

    public string LegalEntityId
    {
        get
        {
            return _LeagalEntity;
        }
        set
        {
            _LeagalEntity = value;
        }
    }

    public string FundId
    {
        get
        {
            return _Fund;
        }
        set
        {
            _Fund = value;
        }
    }

    public string HouseHoldValue
    {
        get
        {
            return HouseHold_Value;
        }
        set
        {
            HouseHold_Value = value;
        }
    }

    public string HouseHoldText
    {
        get
        {
            return HouseHold_Text;
        }
        set
        {
            HouseHold_Text = value;
        }
    }

    public string ReportRollUpGroupValue
    {
        get
        {
            return ReportRollUpGroup_Value;
        }
        set
        {
            ReportRollUpGroup_Value = value;
        }
    }

    public string AllocationGroupValue
    {
        get
        {
            return AllocationGroup_Value;
        }
        set
        {
            AllocationGroup_Value = value;
        }
    }

    public string AllocationGroupText
    {
        get
        {
            return AllocationGroup_Text;
        }
        set
        {
            AllocationGroup_Text = value;
        }
    }

    public string AsOfDate
    {
        get
        {
            return AsOf_Date;
        }
        set
        {
            AsOf_Date = value;
        }
    }

    public string lsFamiliesName
    {
        get
        {
            return lsFamilies_Name;
        }
        set
        {
            lsFamilies_Name = value;
        }
    }

    public string lsDateName
    {
        get
        {
            return lsDate_Name;
        }
        set
        {
            lsDate_Name = value;
        }
    }

    public string GreshamAdvisedFlag
    {
        get
        {
            return _GreshamAdvisedFlag;
        }
        set
        {
            _GreshamAdvisedFlag = value;
        }
    }

    public string PortFolioConChartRptVer
    {
        get
        {
            return _PorthFolioConChartReportVer;
        }
        set
        {
            _PorthFolioConChartReportVer = value;
        }
    }

    public string Chart1
    {
        get
        {
            return _Chart1;
        }
        set
        {
            _Chart1 = value;
        }
    }

    public string Chart2
    {
        get
        {
            return _Chart2;
        }
        set
        {
            _Chart2 = value;
        }
    }

    public string AssetClassCSV
    {
        get
        {
            return _AssetClassCSV;
        }
        set
        {
            _AssetClassCSV = value;
        }
    }

    public string ReportRollupGroupIdName
    {
        get
        {
            return _ReportRollupGroupIdName;
        }
        set
        {
            _ReportRollupGroupIdName = value;
        }
    }

    public string ReportingName
    {
        get
        {
            return _ReportingName;
        }
        set
        {
            _ReportingName = value;
        }
    }

    public string ReportSource
    {

        get
        {
            return _ReportSource;
        }
        set
        {
            _ReportSource = value;
        }
    }

    public string PriorDate
    {
        get
        {
            return _PriorDate;
        }
        set
        {
            _PriorDate = value;
        }
    }
    //added 2_1_2019 Non marketable (DYNAMO)
    public string ReportRollupGroupId
    {
        get
        {
            return _ReportRollupGroupId;
        }
        set
        {
            _ReportRollupGroupId = value;
        }
    }
    public string HouseholdId
    {
        get
        {
            return _HouseholdId;
        }
        set
        {
            _HouseholdId = value;
        }
    }
    public string FundIRR
    {
        get
        {
            return _FundIRR;
        }
        set
        {
            _FundIRR = value;
        }
    }
    public string HHParameterTxt
    {
        get
        {
            return _HHParameterTxt;
        }
        set
        {
            _HHParameterTxt = value;
        }
    }
    public string ReportingID
    {
        get
        {
            return _ReportingID;
        }
        set
        {
            _ReportingID = value;
        }
    }
    public string ReportName
    {
        get
        {
            return _ReportName;
        }
        set
        {
            _ReportName = value;
        }
    }
    public string TempFolderPath
    {
        get
        {
            return _TempFolderPath;
        }
        set
        {
            _TempFolderPath = value;
        }
    }
    #endregion

    #region Code to Generate Reports

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
        //Commented 8_14_2019(Batch Mixup issue)
        DataTable table = newdataset.Tables[1];
        // Random rnd = new Random();
        //string strRndNumber = Convert.ToString(rnd.Next());
        //string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + strRndNumber;

        //String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\PC_" + strGUID + ".xls";

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
                                                                                                                              //   String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));
        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);
        double cr = Convert.ToDouble(table.Rows[6][1]);
        if (cr == 0)
        {
            lsTotalNumberofColumns = 12 + "";
        }
        else
        {
            lsTotalNumberofColumns = 13 + "";
        }
        // lsTotalNumberofColumns = 12 + "";

        iTextSharp.text.Table Table = new iTextSharp.text.Table(13, 3);
        iTextSharp.text.Cell cell = new Cell();

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(13, 15);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        setTableProperty1(loTable);
        setTableProperty1(Table);

        int liTotalPage = 1;// (newdataset.Tables[0].Rows.Count / liPageSize);
        int liCurrentPage = 0;
        liPageSize = 38;

        iTextSharp.text.Chunk lochunk = new Chunk();
        iTextSharp.text.Chunk chunk = new Chunk();

        if (table.Rows.Count > 0)
        {
            Decimal DirectionalPerc = Math.Round(Convert.ToDecimal(table.Rows[2][2]) + Convert.ToDecimal(table.Rows[3][2]), 1);
            Decimal NonDirPerc = Math.Round(Convert.ToDecimal(table.Rows[4][2]) + Convert.ToDecimal(table.Rows[5][2]), 1);
            string DirectValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[2][1]) + Convert.ToInt32(table.Rows[3][1]));
            string NonDirectValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[4][1]) + Convert.ToInt32(table.Rows[5][1]));
            Decimal TotalPerc = Math.Round(DirectionalPerc + NonDirPerc, 1);
            string CorePortfTotal = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(DirectValue) + Convert.ToDecimal(NonDirectValue));

            Decimal CashValue = Convert.ToDecimal(table.Rows[0][1]);
            Decimal FixedIncomeValue = Convert.ToDecimal(table.Rows[1][1]);
            Decimal DomesticEqueityValue = Convert.ToDecimal(table.Rows[2][1]);
            Decimal GlobalOppValue = Convert.ToDecimal(table.Rows[4][1]);
            Decimal InternationEqValue = Convert.ToDecimal(table.Rows[3][1]);
            Decimal HedgeValue = Convert.ToDecimal(table.Rows[5][1]);
            Decimal ConcentratedValue = Convert.ToDecimal(table.Rows[6][1]);
            Decimal LiquidValue = Convert.ToDecimal(table.Rows[7][1]);
            Decimal ILiquidValue = Convert.ToDecimal(table.Rows[8][1]);
            Decimal PrivateEqValue = Convert.ToDecimal(Convert.ToDecimal(table.Rows[9][1]));
            int CashPerc = Convert.ToInt32(Math.Round(Convert.ToDecimal(table.Rows[0][2])));
            int FixedIncomPerc = Convert.ToInt32(Math.Round(Convert.ToDecimal(table.Rows[1][2])));
            int ConcPerc = Convert.ToInt32(Math.Round(Convert.ToDecimal(table.Rows[6][2])));
            int LiquidPerc = Convert.ToInt32(Math.Round(Convert.ToDecimal(table.Rows[7][2])));
            int ILLiquidPerc = Convert.ToInt32(Math.Round(Convert.ToDecimal(table.Rows[8][2])));
            int PrivatePerc = Convert.ToInt32(Math.Round(Convert.ToDecimal(table.Rows[9][2])));

            //Decimal rtgh = CashValue + FixedIncomeValue + Convert.ToDecimal(CorePortfTotal) + ConcentratedValue + LiquidValue;
            Decimal rtgh = CashValue + FixedIncomeValue + Convert.ToDecimal(table.Rows[2][1]) + Convert.ToDecimal(table.Rows[3][1]) + Convert.ToDecimal(table.Rows[4][1]) + Convert.ToDecimal(table.Rows[5][1]) + ConcentratedValue + LiquidValue;
            string LiquidAssetValue = String.Format("{0:#,###0;(#,###0)}", rtgh);
            string ILLiquidAssetValue = String.Format("{0:#,###0;(#,###0)}", ILiquidValue + PrivateEqValue);
            Decimal LiquidAssetPerc = 0;     // TotalPerc + CashPerc + FixedIncomPerc + ConcPerc + LiquidPerc;
            Decimal ILLiquidAssetPerc = 0;  // ILLiquidPerc + PrivatePerc;

            Decimal dLiquidAssetPerc = Math.Round(Convert.ToDecimal(table.Rows[2][2]) + Convert.ToDecimal(table.Rows[3][2]) + Convert.ToDecimal(table.Rows[4][2]) + Convert.ToDecimal(table.Rows[5][2]) + Convert.ToDecimal(table.Rows[0][2]) + Convert.ToDecimal(table.Rows[1][2]) + Convert.ToDecimal(table.Rows[6][2]) + Convert.ToDecimal(table.Rows[7][2]), 1);
            Decimal dILLiquidAssetPerc = Math.Round(Convert.ToDecimal(table.Rows[8][2]) + Convert.ToDecimal(table.Rows[9][2]), 1);
            LiquidAssetPerc = dLiquidAssetPerc;
            ILLiquidAssetPerc = dILLiquidAssetPerc;

            string FinalValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(rtgh) + Convert.ToDecimal(ILiquidValue + PrivateEqValue));

            if (AsOfDate != "")
                lsDateName = Convert.ToDateTime(AsOfDate).ToString("MMMM dd, yyyy") + "";

            #region Data Values

            for (int i = 0; i < 16; i++)
            {
                int colsize = 13;
                for (int j = 0; j < colsize; j++)
                {
                    // string Text = "i=" + i.ToString() +": j="+ j.ToString()+",";// Convert.ToString(newdataset.Tables[0].Rows[i][j]);
                    string Text = "";// i.ToString() + ":" + j.ToString();
                    string lsfamilyName = "";

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
                        lochunk = new Chunk("Cash" + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        SetBorder(loCell, false, true, false, false);
                        loCell.Add(lochunk);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 2)
                    {
                        lochunk = new Chunk("Fixed Income" + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("Core Equity Portfolio	" + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("Concentrated" + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, false, true, false, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 9)
                    {
                        lochunk = new Chunk("Real Assets" + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("Private Equity" + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 2;
                        SetBorder(loCell, false, true, false, false);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 4)
                    {
                        lochunk = new Chunk("Directional" + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 3;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 5)
                    {
                        lochunk = new Chunk("Non-Directional" + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        //loCell.Colspan = 3;
                        loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 0)
                    {
                        lochunk = new Chunk("Cash\n\n\n" + RoundUp(Convert.ToString(table.Rows[0][2])) + Text + "%", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("Fixed Income\n\n\n" + RoundUp(Convert.ToString(table.Rows[1][2])) + Text + "%", setFontsAll(9, 0, 0));
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
                        string Domequity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[2][1]));
                        lochunk = new Chunk("Domestic Equity\n\n" + RoundUp(Convert.ToString(table.Rows[2][2])) + Text + "%\n$" + Domequity + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B8CCE4"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 5)
                    {
                        string GlobOpp = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[4][1]));
                        lochunk = new Chunk("Global Opportunistic\n\n" + RoundUp(Convert.ToString(table.Rows[4][2])) + Text + "%\n$" + GlobOpp + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("Concentrated Positions\n\n" + RoundUp(Convert.ToString(table.Rows[6][2])) + Text + "%", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("Liquid Real Assets\n\n" + RoundUp(Convert.ToString(table.Rows[7][2])) + Text + "%", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("Illiquid Real Assets\n\n" + RoundUp(Convert.ToString(table.Rows[8][2])) + Text + "%", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, false, false, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Leading = 30f;
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
                        lochunk = new Chunk("Private Equity\n\n\n" + RoundUp(Convert.ToString(table.Rows[9][2])) + Text + "%", setFontsAll(9, 0, 0));
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
                        string Intquity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[3][1]));
                        lochunk = new Chunk("International Equity\n\n" + RoundUp(Convert.ToString(table.Rows[3][2])) + Text + "%\n$" + Intquity + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B8CCE4"));

                        SetBorder(loCell, false, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 5 && j == 5) //Hedged Strategies Heading
                    {
                        string HedgedStr = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[5][1]));
                        lochunk = new Chunk("Hedged Strategies\n\n" + RoundUp(Convert.ToString(table.Rows[5][2])) + Text + "%\n$" + HedgedStr + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B8CCE4"));
                        SetBorder(loCell, false, true, false, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }

                    else if (i == 6 && j == 4) //Directional %
                    {
                        string strDirectPerct = RoundUp(Convert.ToString((Convert.ToDecimal(table.Rows[2][2]) + Convert.ToDecimal(table.Rows[3][2]))));
                        lochunk = new Chunk(strDirectPerct + "%" + Text + "", setFontsAll(9, 0, 0));
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
                        string strNonDirPerc = RoundUp(Convert.ToString((Convert.ToDecimal(table.Rows[4][2]) + Convert.ToDecimal(table.Rows[5][2]))));
                        lochunk = new Chunk(strNonDirPerc + "%" + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("$" + DirectValue + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("$" + NonDirectValue + Text + "", setFontsAll(9, 0, 0));
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
                        string strTotalperc = RoundUp(Convert.ToString(Convert.ToDecimal(table.Rows[2][2]) + Convert.ToDecimal(table.Rows[3][2]) + Convert.ToDecimal(table.Rows[4][2]) + Convert.ToDecimal(table.Rows[5][2])));
                        lochunk = new Chunk(strTotalperc + "%" + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("$" + String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[0][1])) + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 9 && j == 2)//Fixed Income Value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[1][1])) + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 9 && j == 4)// Core Portfolio value
                    {
                        lochunk = new Chunk("$" + CorePortfTotal + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("$" + String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[6][1])) + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 9 && j == 9)// Liquid Real Assets value
                    {
                        lochunk = new Chunk("$" + String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[7][1])) + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("$" + String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[8][1])) + Text + "", setFontsAll(9, 0, 0));
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
                        lochunk = new Chunk("$" + String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(table.Rows[9][1])) + Text + "", setFontsAll(9, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 10 && j == 0)// | border
                    {
                        lochunk = new Chunk("  ");
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
                        string strLiquidAssetPerc = RoundUp(Convert.ToString(Convert.ToDecimal(table.Rows[2][2]) + Convert.ToDecimal(table.Rows[3][2]) + Convert.ToDecimal(table.Rows[4][2]) + Convert.ToDecimal(table.Rows[5][2]) + Convert.ToDecimal(table.Rows[0][2]) + Convert.ToDecimal(table.Rows[1][2]) + Convert.ToDecimal(table.Rows[6][2]) + Convert.ToDecimal(table.Rows[7][2])));
                        if (cr == 0)
                        {
                            lochunk = new Chunk("                                     " + "Liquid Assets\n                                     $" + LiquidAssetValue + "\n                                   " + strLiquidAssetPerc + "" + Text + "%", setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            lochunk = new Chunk("Liquid Assets\n$" + LiquidAssetValue + "\n " + strLiquidAssetPerc + "" + Text + "%", setFontsAll(9, 0, 0));
                        }

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
                        string strILLiquidAssetPerc = RoundUp(Convert.ToString(Convert.ToDecimal(table.Rows[8][2]) + Convert.ToDecimal(table.Rows[9][2])));
                        lochunk = new Chunk("Illiquid Assets\n$" + ILLiquidAssetValue + "\n" + strILLiquidAssetPerc + "%" + Text + "", setFontsAll(9, 0, 0));
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
                    //else if (i == 13 && j == 0)// -- border
                    //{
                    //    chunk = new Chunk("");
                    //    cell = new iTextSharp.text.Cell();
                    //    cell.Add(chunk);
                    //    cell.BackgroundColor = iTextSharp.text.Color.WHITE;

                    //    cell.Colspan = 13;
                    //    j = j + 12;
                    //    cell.BorderWidthLeft = 1F;
                    //    cell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    cell.BorderWidthRight = 1F;
                    //    cell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    cell.BorderWidthBottom = 1F;
                    //    cell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    //    //loCell.Height = 5;
                    //    Table.AddCell(cell);
                    //}

                    //else if (i == 14 && j == 0)// -- border
                    //{
                    //    lochunk = new Chunk("");
                    //    cell.Add(lochunk);
                    //    cell.BackgroundColor = iTextSharp.text.Color.WHITE;
                    //    cell.Colspan = 6;
                    //    j = j + 5;
                    //    cell.BorderWidthLeft = 1F;
                    //    cell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                    //    Table.AddCell(cell);
                    //    //lochunk = new Chunk("");
                    //    //loCell = new iTextSharp.text.Cell();
                    //    //loCell.Add(lochunk);
                    //    //loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                    //    //loCell.Colspan = 6;
                    //    //j = j + 5;
                    //    ////loCell.BorderWidthLeft = 1F;
                    //    ////loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    //loCell.BorderWidthRight = 1F;
                    //    //loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    //loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    //    ////loCell.Height =iTextSharp.text.Cell.
                    //    //loTable.AddCell(loCell);
                    //}
                    //else if (i == 14 && j == 6)// -- border
                    //{

                    //    lochunk = new Chunk("");
                    //    cell.Add(lochunk);
                    //    cell.BackgroundColor = iTextSharp.text.Color.WHITE;
                    //    cell.Colspan = 7;
                    //    j = j + 6;
                    //    cell.BorderWidthLeft = 1F;
                    //    cell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                    //    Table.AddCell(cell);

                    //    //lochunk = new Chunk("");
                    //    //loCell = new iTextSharp.text.Cell();
                    //    //loCell.Add(lochunk);
                    //    //loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                    //    //loCell.Colspan = 7;
                    //    //j = j + 6;
                    //    ////loCell.BorderWidthLeft = 1F;
                    //    ////loCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    //loCell.BorderWidthRight = 1F;
                    //    //loCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                    //    //loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    //    ////loCell.Height =iTextSharp.text.Cell.
                    //    //loTable.AddCell(loCell);
                    //}
                    //else if (i == 15 && j == 0)// Total Portfolio Heading
                    //{
                    //    if (cr == 0)
                    //    {
                    //        lochunk = new Chunk("                                               " + "Total Portfolio\n                                                $" + FinalValue + "" + Text + "", setFontsAll(9, 0, 0));
                    //    }
                    //    else
                    //    {
                    //        lochunk = new Chunk("Total Portfolio\n $" + FinalValue + "" + Text + "", setFontsAll(9, 0, 0));
                    //    }
                    //    loCell = new iTextSharp.text.Cell();

                    //    loCell.Add(lochunk);
                    //    loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                    //    loCell.Colspan = 13;
                    //    j = j + 12;
                    //    loCell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                    //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    //    loTable.AddCell(loCell);
                    //}
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
            #endregion

            #region Set Table Center Border

            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 13; j++)
                {
                    if (i == 0 && j == 0)// -- border
                    {
                        //loCell = new iTextSharp.text.Cell();
                        chunk = new Chunk(" ");
                        cell = new iTextSharp.text.Cell();
                        cell.Add(chunk);
                        cell.BackgroundColor = iTextSharp.text.Color.WHITE;

                        cell.Colspan = 13;
                        j = j + 12;
                        cell.BorderWidthLeft = 1F;
                        cell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        cell.BorderWidthRight = 1F;
                        cell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        cell.BorderWidthBottom = 1F;
                        cell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        Table.AddCell(cell);
                    }
                    else if (i == 1 && j == 0)// -- border
                    {
                        chunk = new Chunk(" ");
                        cell = new iTextSharp.text.Cell();
                        cell.Add(chunk);
                        cell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        cell.Colspan = 6;
                        j = j + 5;
                        cell.BorderWidthRight = 1F;
                        cell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.White);
                        cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        Table.AddCell(cell);
                    }
                    else if (i == 2 && j == 0)// -- border
                    {
                        chunk = new Chunk(" ");
                        cell = new iTextSharp.text.Cell();
                        cell.Add(chunk);
                        cell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        cell.Colspan = 7;
                        j = j + 6;
                        cell.BorderWidthLeft = 1F;
                        cell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        Table.AddCell(cell);
                    }
                    else if (i == 3 && j == 0)// Total Portfolio Heading
                    {
                        if (cr == 0)
                        {
                            chunk = new Chunk("      " + "Total Portfolio\n    $" + FinalValue + "", setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            chunk = new Chunk("   Total Portfolio\n     $" + FinalValue + "", setFontsAll(9, 0, 0));
                        }
                        cell = new iTextSharp.text.Cell();

                        cell.Add(chunk);
                        cell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        cell.Colspan = 13;
                        j = j + 12;
                        cell.Border = 0;//iTextSharp.text.Cell.RECTANGLE;
                        cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        Table.AddCell(cell);
                    }
                }
            }



            #endregion

            /**** to remove Concentrated Positions if zero  ****/
            double c = Convert.ToDouble(table.Rows[6][1]);
            if (c == 0)
            {
                loTable.DeleteColumn(7); //Concentrated Positions
                loTable.Width = 80f;
            }
            // loTable.DeleteColumn(8);

            /**** end ****/

            document.Add(loTable);
            document.Add(Table);

            if (newdataset.Tables[0].Rows.Count > 0)
            {
                document.Close();

                //Commented 8_14_2019(Batch Mixup issue)
                //FileInfo loFile = new FileInfo(ls);
                //try
                //{
                //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

                //}
                //catch (Exception exc)
                //{
                //}

            }
            return fsFinalLocation.Replace(".xls", ".pdf");
        }
        else
        {
            return "Record not found";
        }

    }

    private string RoundUp(string lsFormatedString)
    {
        if (lsFormatedString != "")
        {
            lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
            return lsFormatedString;
        }
        else
            return "";
    }

    public string generateCommittmentSchReport()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        //String lsFooterTxt = "See notes for this illustration located in the Appendix under Commitment Schedule for important information.";
        String lsFooterTxt = FooterText;

        String lsFooterlocation = Footerlocation;

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
            //lblError.Text = "Record not found";
            return "Record not found";
        }
        //lodataset.Tables.Add(table);
        DataTable loInsertblankRow = table.Copy();
        //loInsertblankRow.Tables.Add(table);
        //lodataset.Tables[0].Clear();
        table.Clear();
        table = null;
        table = loInsertblankRow.Clone();

        //Random rnd = new Random();
        //Convert.ToString(rnd.Next());

        // string strGUID = Guid.NewGuid().ToString() + "_" + Convert.ToString(rnd.Next());
        //string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss");
        // strGUID = strGUID.Substring(0, 5);
        // String fsFinalLocation = @"C:\Reports\" + strGUID + ".xls";
        //Commented 8_14_2019(Batch Mixup issue)
        // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

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
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));
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
                //  loInsertdataset.Tables[0].Rows.Count % liPageSize
                if (liRowCount != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty, lsFooterlocation, String.Empty, Ssi_GreshamClientFooter));
                    document.NewPage();
                    SetTotalPageCount("Commitment Schedule");
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
                        if (liColumnCount == 9 && lsFormatedString == "")
                        {
                            lsFormatedString = "N/A";
                        }
                        else
                        {
                            lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                        }
                    }
                    else
                    {
                        if (liColumnCount == 9 && lsFormatedString == "")
                        {
                            lsFormatedString = "N/A";
                        }
                        else
                        {
                            lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                        }
                    }

                    if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
                    {
                        if (liColumnCount == 9 && lsFormatedString == "")
                        {
                            lsFormatedString = "";
                        }
                        else
                        {
                            lsFormatedString = "";
                        }
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

                if (liColumnCount == 9)
                {
                    if (lsFormatedString == "N/A")
                    {
                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "3")
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                            lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                        }

                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "2")
                        {
                            lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                        }
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    }
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
                                    if (liColumnCount == 9 && Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["AverageQuarterlyDist"]) == "")
                                    {
                                        lsFormatedString = "N/A";
                                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                        lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                    }
                                    else
                                    {
                                        lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                        lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                    }
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
                                    if (liColumnCount == 9 && Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["AverageQuarterlyDist"]) == "")
                                    {
                                        //loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                        lsFormatedString = "N/A";
                                        //lochunk = new Chunk(lsFormatedString, Font18Bold(lsFormatedString));

                                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                        loCell.VerticalAlignment = 4;
                                        loCell.Leading = 10f;
                                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                        //lsFormatedString = "TOTAL COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                        lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                                    }
                                    else
                                    {
                                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                        lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                        lochunk = new Chunk(lsFormatedString, Font19Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                    }
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
                        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "")
                        {
                            if (liColumnCount == 9 && Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["AverageQuarterlyDist"]) == "")
                            {
                                lsFormatedString = "";
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                            }
                        }
                    }

                }
                if (checkTrue(lodataset, liRowCount, "_OrderNmb") && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                {
                    if (liColumnCount == 0)
                    {
                        String abc = "          " + "          " + lodataset.Tables[0].Rows[liRowCount][0].ToString();
                        lochunk = new Chunk(abc, Font7Grey());
                    }
                    else
                    {
                        lochunk = new Chunk(lsFormatedString, Font7Grey());
                    }
                }

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
                    SetTotalPageCount("Commitment Schedule");
                    document.Add(loTable);
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt, lsFooterlocation, ClientFooterTxt, Ssi_GreshamClientFooter));
                }
            }
            catch (Exception Ex)
            {

            }
        }
        document.Close();

        //try
        //{

        //    FileInfo loFile = new FileInfo(ls);
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch
        //{ }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    public string generateInvestmentObjectiveChart()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = FooterText;
        String lsFooterLocation = Footerlocation;
        String lsSQL = getFinalSp(ReportType.InvestmentObjectiveChart);
        newdataset = clsDB.getDataSet(lsSQL);

        for (int i = 0; i < newdataset.Tables[0].Rows.Count; i++)
        {
            DataRow newRow = newdataset.Tables[0].NewRow();
            newRow["Asset Class"] = "Directional";// for borderstyle

            DataRow NonDir = newdataset.Tables[0].NewRow();
            NonDir["Asset Class"] = "Non Directional";// for borderstyle

            if (i != 0)
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "1" && Convert.ToString(newdataset.Tables[0].Rows[i - 1]["IndicatorFlg"]) != "1")
                {
                    newdataset.Tables[0].Rows.InsertAt(newRow, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }
                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "2" && Convert.ToString(newdataset.Tables[0].Rows[i - 1]["IndicatorFlg"]) != "2")
                {
                    newdataset.Tables[0].Rows.InsertAt(NonDir, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }
            }
            else
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "1")
                {

                    newdataset.Tables[0].Rows.InsertAt(newRow, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }
                if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "2")
                {
                    newdataset.Tables[0].Rows.InsertAt(NonDir, i);
                    newdataset.Tables[0].AcceptChanges();
                    i++;
                }
            }
        }
        for (int j = 0; j < newdataset.Tables[0].Rows.Count; j++)
        {
            if (Convert.ToString(newdataset.Tables[0].Rows[j]["IndicatorFlg"]) == "1")
            {
                newdataset.Tables[0].Rows[j].BeginEdit();
                newdataset.Tables[0].Rows[j]["Asset Class"] = "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + Convert.ToString(newdataset.Tables[0].Rows[j]["Asset Class"]);
                newdataset.Tables[0].Rows[j].EndEdit();
                newdataset.Tables[0].AcceptChanges();
            }

            if (Convert.ToString(newdataset.Tables[0].Rows[j]["IndicatorFlg"]) == "2")
            {
                newdataset.Tables[0].Rows[j].BeginEdit();
                newdataset.Tables[0].Rows[j]["Asset Class"] = "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + Convert.ToString(newdataset.Tables[0].Rows[j]["Asset Class"]);
                newdataset.Tables[0].Rows[j].EndEdit();
                newdataset.Tables[0].AcceptChanges();
            }
        }
        newdataset = AddTotalsInvestmentObjectiveChart(newdataset);

        if (newdataset.Tables[0].Rows.Count < 1)
        {
            //lblError.Text = "No Record Found";
            return "No Record Found";
        }

        DataSet loInsertblankRow = newdataset.Copy();

        newdataset = loInsertblankRow.Clone();

        //Random rnd = new Random();


        //string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + Convert.ToString(rnd.Next());

        // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\IOC_" + strGUID + ".xls";
        int liBlankCounter = 0;

        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Columns.Count; liBlankRow++)
        {
            if (liBlankRow == 7)
            {
                loInsertblankRow.Tables[0].Columns.RemoveAt(liBlankRow);

                for (int i = 0; i < loInsertblankRow.Tables[0].Columns.Count; i++)
                {
                    if (i == 7)
                    {
                        loInsertblankRow.Tables[0].Columns.RemoveAt(i);
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
        loInsertdataset.AcceptChanges();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));
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
        String SQL = getFinalSp(ReportType.InvestmentObjectiveChart);

        newdataset = clsDB.getDataSet(SQL);

        newdataset = AddTotalsInvestmentObjectiveChart(newdataset);

        for (int liRowCount = 0; liRowCount < loInsertblankRow.Tables[0].Rows.Count; liRowCount++)
        {
            if (liRowCount % liPageSize == 0)
            {
                document.Add(loTable);

                if (liRowCount != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty, String.Empty, String.Empty, String.Empty));
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
                    if (liColumnCount == 1)
                    {
                        lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                    }
                    else if (lsFormatedString == "0" && liColumnCount == 4)
                    {
                        lsFormatedString = String.Format("", Convert.ToDecimal(lsFormatedString));
                    }
                    else if (liColumnCount == loInsertdataset.Tables[0].Columns.Count)
                    {
                        lsFormatedString = String.Format("${0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(lsFormatedString));
                    }
                    else
                    {
                        lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(lsFormatedString));
                    }
                }
                catch
                {
                }
                lochunk = new Chunk(lsFormatedString, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                loCell = new iTextSharp.text.Cell();
                loCell.Border = 0;
                loCell.NoWrap = true;
                //loCell.VerticalAlignment=0;
                loCell.VerticalAlignment = 5;

                setGreyBorder(lodataset, loCell, liRowCount);
                loCell.Leading = 4f;//6

                loCell.UseBorderPadding = true;

                if (liColumnCount == 1 || liColumnCount == 4 || liColumnCount == 3 || liColumnCount == 5 || liColumnCount == 6)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                }
                else if (liColumnCount == 2)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                }

                if (liColumnCount == 2 || liColumnCount == 3 || liColumnCount == 5 || liColumnCount == 6)
                {
                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "1")
                    {
                        //loCell.EnableBorderSide(2);
                    }
                    else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        try
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                            //loCell.EnableBorderSide(1);
                            if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) == "0" && liColumnCount == 5)
                            {
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                //loCell.EnableBorderSide(1);
                                string CurrentAllocation = String.Format("", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            }
                            else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "0")
                            {
                                string CurrentAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            }
                        }
                        catch
                        {

                        }
                    }
                }
                else if (liColumnCount == 1 || liColumnCount == 4)
                {
                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "1")
                    {
                        //loCell.EnableBorderSide(1);
                    }
                    else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        //loCell.EnableBorderSide(1);
                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                        {
                            string CurrentAllocation = String.Format("${0:#,###0;(#,###0.0)}", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                            //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                            lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                        }
                    }
                }
                else if (liColumnCount == 0)
                {
                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["_LineFlg"]) == "2")
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        string CurrentAllocation = Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]);
                        lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                    }

                }
                loCell.Add(lochunk);
                loTable.AddCell(loCell);
            }

            try
            {
                if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
                {
                    SetTotalPageCount("Investment Objective Chart");
                    document.Add(loTable);
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt, lsFooterLocation, ClientFooterTxt, Ssi_GreshamClientFooter));
                }
            }
            catch (Exception Ex)
            {

            }
        }

        document.Close();

        //try
        //{

        //    FileInfo loFile = new FileInfo(ls);
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch
        //{ }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    public string generateOverAllPieChart()
    {
        Random rnd = new Random();
        Convert.ToString(rnd.Next());

        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + Convert.ToString(rnd.Next());
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

        //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".xls";

        DataSet newdataset;
        DB clsDB = new DB();
        String lsSQL = getFinalSp(ReportType.OverallPieChart);
        newdataset = clsDB.getDataSet(lsSQL);
        DataTable table = newdataset.Tables[1];
        AppDomain.CurrentDomain.Load("JCommon");
        org.jfree.data.general.DefaultPieDataset myDataSet = new org.jfree.data.general.DefaultPieDataset();
        if (table.Rows.Count > 0)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (Convert.ToString(table.Rows[i][2]) != "0")
                {
                    myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[i][2]), 1)));
                }
            }
            //if (Convert.ToString(table.Rows[0][2]) != "0")
            //    myDataSet.setValue("Cash and Equivalents", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[0][2]), 1)));
            //if (Convert.ToString(table.Rows[1][2]) != "0")
            //    myDataSet.setValue("Fixed Income", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[1][2]), 1)));
            //if (Convert.ToString(table.Rows[2][2]) != "0")
            //    myDataSet.setValue("Low Volatility Hedged Strategies", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[2][2]), 1)));
            //if (Convert.ToString(table.Rows[3][2]) != "0")
            //    myDataSet.setValue("Domestic Equity", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[3][2]), 1)));
            //if (Convert.ToString(table.Rows[4][2]) != "0")
            //    myDataSet.setValue("International Equity", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[4][2]), 1)));
            //if (Convert.ToString(table.Rows[5][2]) != "0")
            //    myDataSet.setValue("Global Opportunistic", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[5][2]), 1)));
            //if (Convert.ToString(table.Rows[6][2]) != "0")
            //    myDataSet.setValue("Hedged Strategies", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[6][2]), 1)));
            //if (Convert.ToString(table.Rows[7][2]) != "0")
            //    myDataSet.setValue("Concentrated Positions", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[7][2]), 1)));
            //if (Convert.ToString(table.Rows[8][2]) != "0")
            //    myDataSet.setValue("Liquid Real Assets", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[8][2]), 1)));
            //if (Convert.ToString(table.Rows[9][2]) != "0")
            //    myDataSet.setValue("Illiquid Real Assets", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[9][2]), 1)));
            //if (Convert.ToString(table.Rows[10][2]) != "0")
            //    myDataSet.setValue("Private Equitys", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[10][2]), 1)));
            //if (Convert.ToString(table.Rows[11][2]) != "0")
            //    myDataSet.setValue("Other Assets", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[11][2]), 1)));
        }
        JFreeChart pieChart = ChartFactory.createPieChart3D(lsFamiliesName + "\n" + lsDateName, myDataSet, false, true, false);

        pieChart.setBackgroundPaint(java.awt.Color.white);
        pieChart.setBorderVisible(false);

        pieChart.setTitle(new org.jfree.chart.title.TextTitle("", new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12)));

        PiePlot ColorConfigurator = (PiePlot)pieChart.getPlot();
        ColorConfigurator.setLabelBackgroundPaint(System.Drawing.Color.White);// ColorConfigurator.getLabelPaint()
        ColorConfigurator.setLabelOutlinePaint(System.Drawing.Color.White);
        ColorConfigurator.setLabelShadowPaint(System.Drawing.Color.White);

        ColorConfigurator.setLabelFont(new System.Drawing.Font("Frutiger55", 10));

        ColorConfigurator.setCircular(false);
        ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardPieSectionLabelGenerator("{0} =  {1}%"));

        java.util.List keys = myDataSet.getKeys();

        for (int i = 0; i < keys.size(); i++)
        {
            if (keys.get(i).ToString() == "Cash and Equivalents")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4271A5"));
            if (keys.get(i).ToString() == "Fixed Income")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#AD4542"));
            if (keys.get(i).ToString() == "Domestic Equity")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#84A24A"));
            if (keys.get(i).ToString() == "International Equity")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#73558C"));
            if (keys.get(i).ToString() == "Global Opportunistic")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4296AD"));
            if (keys.get(i).ToString() == "Hedged Strategies")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#DE8239"));
            if (keys.get(i).ToString() == "Concentrated Positions")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#94A6CE"));
            if (keys.get(i).ToString() == "Liquid Real Assets")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#94A6CE"));
            if (keys.get(i).ToString() == "Illiquid Real Assets")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CE9294"));
            if (keys.get(i).ToString() == "Private Equitys")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#B5CB94"));
            if (keys.get(i).ToString() == "Other Assets")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CEBEDE"));
        }
        ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
        java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
        ChartUtilities.writeChartAsPNG(jos, pieChart, 800, 350);
        //CreatePDFFile();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        //   String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        String ls = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        if (AsOfDate != "")
            lsDateName = Convert.ToDateTime(AsOfDate).ToString("MMMM dd, yyyy") + "";

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

        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(fsFinalLocation.Replace(".xls", ".png"));

        //Give space before image
        jpg.SpacingBefore = 2f;

        document.Add(jpg); //add an image to the created pdf document

        document.Add(addFooter(lsDateTime, 1, 1, liPageSize - 5, true, FooterText, Footerlocation, ClientFooterTxt, Ssi_GreshamClientFooter));

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

    public string generateAllocationGroupPieChart()
    {

        Random rnd = new Random();

        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + Convert.ToString(rnd.Next()); ;
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        //Commented 8_14_2019(Batch Mixup issue)
        //String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\AL_" + strGUID + ".xls";
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".xls";

        DataSet newdataset;
        DB clsDB = new DB();
        String lsSQL = getFinalSp(ReportType.AllocationGroupPieChart);
        newdataset = clsDB.getDataSet(lsSQL);
        DataTable table = newdataset.Tables[1];
        AppDomain.CurrentDomain.Load("JCommon");
        org.jfree.data.general.DefaultPieDataset myDataSet = new org.jfree.data.general.DefaultPieDataset();
        if (table.Rows.Count > 0)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (Convert.ToString(table.Rows[i][2]) != "0")
                {
                    myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[i][2]), 1)));
                }
            }
            //if (Convert.ToString(table.Rows[0][2]) != "0")
            //    myDataSet.setValue("Cash and Equivalents", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[0][2]), 1)));
            //if (Convert.ToString(table.Rows[1][2]) != "0")
            //    myDataSet.setValue("Fixed Income", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[1][2]), 1)));
            //if (Convert.ToString(table.Rows[2][2]) != "0")
            //    myDataSet.setValue("Low Volatility Hedged Strategies", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[2][2]), 1)));
            //if (Convert.ToString(table.Rows[3][2]) != "0")
            //    myDataSet.setValue("Domestic Equity", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[3][2]), 1)));
            //if (Convert.ToString(table.Rows[4][2]) != "0")
            //    myDataSet.setValue("International Equity", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[4][2]), 1)));
            //if (Convert.ToString(table.Rows[5][2]) != "0")
            //    myDataSet.setValue("Global Opportunistic", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[5][2]), 1)));
            //if (Convert.ToString(table.Rows[6][2]) != "0")
            //    myDataSet.setValue("Hedged Strategies", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[6][2]), 1)));
            //if (Convert.ToString(table.Rows[7][2]) != "0")
            //    myDataSet.setValue("Concentrated Positions", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[7][2]), 1)));
            //if (Convert.ToString(table.Rows[8][2]) != "0")
            //    myDataSet.setValue("Liquid Real Assets", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[8][2]), 1)));
            //if (Convert.ToString(table.Rows[9][2]) != "0")
            //    myDataSet.setValue("Illiquid Real Assets", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[9][2]), 1)));
            //if (Convert.ToString(table.Rows[10][2]) != "0")
            //    myDataSet.setValue("Private Equitys", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[10][2]), 1)));
            //if (Convert.ToString(table.Rows[11][2]) != "0")
            //    myDataSet.setValue("Other Assets", Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[11][2]), 1)));
        }
        JFreeChart pieChart = ChartFactory.createPieChart3D(lsFamiliesName + "\n" + lsDateName, myDataSet, false, true, false);

        pieChart.setBackgroundPaint(java.awt.Color.white);
        pieChart.setBorderVisible(false);

        pieChart.setTitle(new org.jfree.chart.title.TextTitle("", new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 12)));

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
            if (keys.get(i).ToString() == "Cash and Equivalents")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4271A5"));
            if (keys.get(i).ToString() == "Fixed Income")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#AD4542"));
            if (keys.get(i).ToString() == "Domestic Equity")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#84A24A"));
            if (keys.get(i).ToString() == "International Equity")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#73558C"));
            if (keys.get(i).ToString() == "Global Opportunistic")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#4296AD"));
            if (keys.get(i).ToString() == "Hedged Strategies")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#DE8239"));
            if (keys.get(i).ToString() == "Concentrated Positions")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#94A6CE"));
            if (keys.get(i).ToString() == "Liquid Real Assets")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#94A6CE"));
            if (keys.get(i).ToString() == "Illiquid Real Assets")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CE9294"));
            if (keys.get(i).ToString() == "Private Equitys")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#B5CB94"));
            if (keys.get(i).ToString() == "Other Assets")
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml("#CEBEDE"));
        }
        ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
        java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
        ChartUtilities.writeChartAsPNG(jos, pieChart, 800, 350);
        //CreatePDFFile();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        String ls = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        lsTotalNumberofColumns = 4 + "";
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(4, 4);   // 2 rows, 2 columns        
        setTableProperty(loTable);

        if (AsOfDate != "")
            lsDateName = Convert.ToDateTime(AsOfDate).ToString("MMMM dd, yyyy") + "";

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
        loTable.AddCell(loCell);

        document.Add(loTable);

        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(fsFinalLocation.Replace(".xls", ".png"));

        //Give space before image
        jpg.SpacingBefore = 2f;

        document.Add(jpg); //add an image to the created pdf document

        document.Add(addFooter(lsDateTime, 1, 1, liPageSize - 5, true, FooterText, Footerlocation, ClientFooterTxt, Ssi_GreshamClientFooter));

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

    public string generateDirectMgrDetail()
    {
        HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);

        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = FooterText;
        String lsFooterLocation = Footerlocation;
        //String lsFooterTxt = "See notes for this illustration located in the Appendix under Commitment Schedule for important information.";

        //String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);
        String lsSQL = getFinalSp(ReportType.DirectMgrDetail);
        // Response.Write(lsSQL);
        newdataset = clsDB.getDataSet(lsSQL);

        DataSet dupdataset = newdataset.Copy();


        if (newdataset.Tables[0].Rows.Count < 1)
        {
            return "No Record Found";
        }


        if (newdataset.Tables[0].Rows.Count > 0)
        {
            for (int j = 0; j < newdataset.Tables[0].Rows.Count; j++)
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[j]["Security"]) != "")
                {
                    newdataset.Tables[0].Rows[j].BeginEdit();
                    newdataset.Tables[0].Rows[j]["Security"] = "\t " + "\t " + "\t " + "\t " + "\t " + Convert.ToString(newdataset.Tables[0].Rows[j]["Security"]);
                    newdataset.Tables[0].Rows[j].EndEdit();
                    newdataset.Tables[0].AcceptChanges();
                }
            }
        }

        #region comment
        //for (int i = 0; i < newdataset.Tables[0].Rows.Count; i++)
        //{
        //    DataRow newRow = newdataset.Tables[0].NewRow();
        //    newRow["Asset Class"] = "Directional";// for borderstyle

        //    DataRow NonDir = newdataset.Tables[0].NewRow();
        //    NonDir["Asset Class"] = "Non Directional";// for borderstyle

        //    if (i != 0)
        //    {
        //        if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "1" && Convert.ToString(newdataset.Tables[0].Rows[i - 1]["IndicatorFlg"]) != "1")
        //        {

        //            newdataset.Tables[0].Rows.InsertAt(newRow, i);
        //            newdataset.Tables[0].AcceptChanges();
        //            i++;
        //        }

        //        if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "2" && Convert.ToString(newdataset.Tables[0].Rows[i - 1]["IndicatorFlg"]) != "2")
        //        {
        //            newdataset.Tables[0].Rows.InsertAt(NonDir, i);
        //            newdataset.Tables[0].AcceptChanges();
        //            i++;
        //        }
        //    }
        //    else
        //    {
        //        if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "1")
        //        {

        //            newdataset.Tables[0].Rows.InsertAt(newRow, i);
        //            newdataset.Tables[0].AcceptChanges();
        //            i++;
        //        }
        //        if (Convert.ToString(newdataset.Tables[0].Rows[i]["IndicatorFlg"]) == "2" )
        //        {
        //            newdataset.Tables[0].Rows.InsertAt(NonDir, i);
        //            newdataset.Tables[0].AcceptChanges();
        //            i++;
        //        }
        //    }


        //}

        //for (int j = 0; j < newdataset.Tables[0].Rows.Count; j++)
        //{
        //    if (Convert.ToString(newdataset.Tables[0].Rows[j]["IndicatorFlg"]) == "1")
        //    {
        //        newdataset.Tables[0].Rows[j].BeginEdit();
        //        newdataset.Tables[0].Rows[j]["Asset Class"] = "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + Convert.ToString(newdataset.Tables[0].Rows[j]["Asset Class"]);
        //        newdataset.Tables[0].Rows[j].EndEdit();
        //        newdataset.Tables[0].AcceptChanges();
        //    }

        //    if (Convert.ToString(newdataset.Tables[0].Rows[j]["IndicatorFlg"]) == "2")
        //    {
        //        newdataset.Tables[0].Rows[j].BeginEdit();
        //        newdataset.Tables[0].Rows[j]["Asset Class"] = "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + Convert.ToString(newdataset.Tables[0].Rows[j]["Asset Class"]);
        //        newdataset.Tables[0].Rows[j].EndEdit();
        //        newdataset.Tables[0].AcceptChanges();
        //    }
        //}

        #endregion

        newdataset = AddTotalsDirectMgr(newdataset);

        if (newdataset.Tables[0].Rows.Count < 1)
        {
            //lblError.Text = "No Record Found";
            return "No Record Found";
        }

        DataSet loInsertblankRow = newdataset.Copy();

        newdataset = loInsertblankRow.Clone();


        Random rnd = new Random();


        // string strGUID = Guid.NewGuid().ToString();
        //  string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + Convert.ToString(rnd.Next());
        // strGUID = strGUID.Substring(0, 5);
        // String fsFinalLocation = @"C:\Reports\" + strGUID + ".xls";

        // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + ".xls";
        String fsFinalLocation = TempFolderPath + System.Guid.NewGuid().ToString() + ".xls";
        //String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\PC_" + strGUID + ".xls";
        int liBlankCounter = 0;

        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Columns.Count; liBlankRow++)
        {
            if (liBlankRow == 7)
            {
                loInsertblankRow.Tables[0].Columns.RemoveAt(liBlankRow);

                for (int i = 0; i < loInsertblankRow.Tables[0].Columns.Count; i++)
                {
                    if (i == 7)
                    {
                        loInsertblankRow.Tables[0].Columns.RemoveAt(i);
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
        //String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        String ls = TempFolderPath + "\\" + System.Guid.NewGuid().ToString() + ".pdf";
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

        //String SQL = getFinalSp(ReportType.DirectMgrDetail);
        // Response.Write(lsSQL);
        //newdataset = clsDB.getDataSet(lsSQL);
        DataSet newdataset1 = AddTotalsDirectMgr(dupdataset);

        for (int liRowCount = 0; liRowCount < loInsertblankRow.Tables[0].Rows.Count; liRowCount++)
        {
            if (liRowCount % liPageSize == 0)
            {
                document.Add(loTable);

                if (liRowCount != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty, String.Empty, String.Empty, String.Empty));
                    document.NewPage();
                    SetTotalPageCount("Direct Mgr Detail Report");
                }

                loInsertdataset.AcceptChanges();
                setHeaderDirectMgrDetails(document, loInsertdataset, newdataset1);
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
                    if (liColumnCount == 1)
                    {
                        lsFormatedString = String.Format(" {0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                    }
                    else if (lsFormatedString == "0" && liColumnCount == 4)
                    {
                        lsFormatedString = String.Format("", Convert.ToDecimal(lsFormatedString));
                    }
                    //else if (lsFormatedString == "0" && liColumnCount == 3)
                    //{
                    //    lsFormatedString = String.Format("", Convert.ToDecimal(lsFormatedString));
                    //}
                    else if (liColumnCount == loInsertdataset.Tables[0].Columns.Count)
                    {
                        lsFormatedString = String.Format("${0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(lsFormatedString));
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
                //loCell.VerticalAlignment=0;
                loCell.VerticalAlignment = 5;

                setGreyBorder(lodataset, loCell, liRowCount);
                loCell.Leading = 4f;//6

                loCell.UseBorderPadding = true;

                if (liColumnCount == 1 || liColumnCount == 2 || liColumnCount == 5 || liColumnCount == 6)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    loCell.EnableBorderSide(1);
                    loCell.BorderColor = iTextSharp.text.Color.LIGHT_GRAY;
                }
                else if (liColumnCount == 0)
                {
                    loCell.EnableBorderSide(1);
                    loCell.BorderColor = iTextSharp.text.Color.LIGHT_GRAY;
                }
                if (loInsertblankRow.Tables[0].Rows.Count - 1 == liRowCount)
                {
                    loCell.DisableBorderSide(1);
                }

                if (liColumnCount == 0 || liColumnCount == 1 || liColumnCount == 2)
                {
                    if (loInsertblankRow.Tables[0].Rows.Count == 1)
                    {
                        loCell.DisableBorderSide(1);
                    }
                }

                if (liColumnCount == 1 || liColumnCount == 2)
                {

                    if (loInsertblankRow.Tables[0].Rows.Count - 1 == liRowCount)
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        loCell.DisableBorderSide(1); //remove border
                        if (liColumnCount == 1)
                        {
                            if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                            {
                                string CurrentAllocation = String.Format("{0:#,###0;(#,###0.0)}", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                                //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                            }
                        }
                        else if (liColumnCount == 2)
                        {
                            if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                            {
                                string CurrentAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                                //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                            }
                        }
                    }
                    else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["PctPortfolio"]) == "" || Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["MarketValue"]) == "")
                    {
                        try
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                            //loCell.EnableBorderSide(1);
                            if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) == "0" && liColumnCount == 5)
                            {
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                //loCell.EnableBorderSide(1);
                                string CurrentAllocation = String.Format("", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            }
                            else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "0")
                            {
                                string CurrentAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                                lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            }

                        }
                        catch
                        {

                        }
                    }
                }
                else if (liColumnCount == 1 || liColumnCount == 2)
                {
                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["Security"]) == "")
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        //loCell.EnableBorderSide(1);
                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                        {
                            string CurrentAllocation = String.Format("${0:#,###0;(#,###0.0)}", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                            //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                            lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                        }
                    }
                    else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["Security"]) == "")
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        //loCell.EnableBorderSide(1);
                        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                        {
                            string CurrentAllocation = String.Format("${0:#,###0;(#,###0.0)}", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount])));
                            //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                            lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                            //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                        }
                    }
                }
                else if (liColumnCount == 0)
                {
                    if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liRowCount]["Security"]).ToUpper() == "Total Porfolio".ToUpper())
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                        //loCell.EnableBorderSide(1);
                        string CurrentAllocation = Convert.ToString(lodataset.Tables[0].Rows[liRowCount][liColumnCount]);
                        //string SuggestedAllocation = String.Format("{0:#,###0;(#,###0.0)}%", Convert.ToDecimal(Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["Suggested Allocation"])));
                        lochunk = new Chunk(CurrentAllocation, setFontsAll(8, 1, 0));
                        //lochunk = new Chunk(SuggestedAllocation, setFontsAll(8, 1, 0));
                    }

                }

                //else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_LineFlg"]) == "0")
                //{
                //    loCell.DisableBorderSide(-1);
                //}
                /////

                #region Not in Use
                /*=========START WITH BOLD AND SUPERBOLD FLAG========*/
                //if (checkTrue(lodataset, liRowCount, "_OrderNmb") || checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                //    try
                //    {
                //        if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1)
                //        {
                //            lsFormatedString = String.Format("${0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
                //        }
                //        else
                //        {
                //            lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
                //        }
                //    }
                //    catch
                //    {

                //    }

                //    //changed on 02/25/2011
                //    //lochunk = new Chunk(lsFormatedString, Font9Bold());
                //    lochunk = new Chunk(lsFormatedString, Font8Bold());
                //    #region Commented
                //    if (!lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE"))
                //    {
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                //        lochunk = new Chunk(lsFormatedString, Font8Bold());
                //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                //        if (lsFormatedString.Length > 25)
                //        {
                //            if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //            {
                //                //decrease columncount by 1 to adjust the Colspan. eg: NON-INVESTMENT ASSETS/LOOK-THROUGHS
                //                loCell.Colspan = 2;
                //                colsize = colsize - 1;
                //            }
                //        }
                //        setBottomWidthWhite(loCell);

                //    } /*=========IF END OF BOLD AND SUPERBOLD FLAG========*/
                //    else
                //    {
                //        if (lodataset.Tables[0].Rows[liRowCount][0].ToString() == "NET CHANGE")
                //        {
                //            setGreyBorder(loCell);
                //            //added on 28Feb2011 to change font size for total
                //            if (liColumnCount != 0)
                //            {
                //                lochunk = new Chunk(lsFormatedString, Font7Bold());
                //            }
                //        }
                //    }

                //    if (lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE %"))
                //    {

                //        lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                //        try
                //        {
                //            lsFormatedString = String.Format("{0:#,###0.0%;(#,###0.0%)}", Convert.ToDecimal(lsFormatedString) / 100);
                //        }
                //        catch
                //        {

                //        }
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                //        lochunk = new Chunk(lsFormatedString, Font8Bold());
                //        //added on 28Feb2011 to change font size for total
                //        if (liColumnCount != 0)
                //        {
                //            lochunk = new Chunk(lsFormatedString, Font7Bold());
                //        }


                //    }
                //    #endregion

                //}
                //else
                //{
                //    if (liColumnCount == 0 && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //    {
                //        String abc = "" + lodataset.Tables[0].Rows[liRowCount][0].ToString();
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(abc, Font9Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                //        lochunk = new Chunk(abc, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));

                //        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "")
                //        {
                //            //loCell.EnableBorderSide(0);
                //            lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "6")
                //        {
                //            checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(216, 216, 216));
                //            lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]), setFontsAll(8, 1, 0));
                //            //lochunk.SetBackground(iTextSharp.text.Color.LIGHT_GRAY);#B7DDE8 new iTextSharp.text.Color(216, 216, 216)
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "3")
                //        {

                //            if (liRowCount == lodataset.Tables[0].Rows.Count - 5)
                //            {
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                lsFormatedString = "TOTAL PROPOSED/CONFIRMED COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }
                //            else if (liRowCount == lodataset.Tables[0].Rows.Count - 4)
                //            {
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                lsFormatedString = "TOTAL PROPOSED/CONFIRMED COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }
                //            else //if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                //            {
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                lsFormatedString = "TOTAL COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }

                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "2")
                //        {
                //            if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                //            {
                //                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 0, 0));
                //            }
                //            else
                //            {
                //                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 0, 0));
                //            }

                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "7")
                //        {
                //            if (liRowCount == lodataset.Tables[0].Rows.Count - 2)
                //            {
                //                loCell.VerticalAlignment = 4;
                //                loCell.Leading = 10f;
                //                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                lsFormatedString = "TOTAL COMMITMENTS ";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(8, 1, 0));
                //            }
                //            else
                //            {
                //                //loCell.EnableBorderSide(1);
                //                lsFormatedString = "";//String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                lochunk = new Chunk(lsFormatedString, setFontsAll(7, 1, 0));
                //            }

                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "8")
                //        {
                //            loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                //            //lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                //        }
                //    }
                //    else if (liColumnCount != 0 && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //    {
                //        if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "2")
                //        {
                //            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                //            {
                //                try
                //                {
                //                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                    lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                }
                //                catch
                //                {

                //                }
                //            }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "3")
                //        {
                //            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                //            {
                //                try
                //                {
                //                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                    lochunk = new Chunk(lsFormatedString, Font19Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                }
                //                catch
                //                {

                //                }
                //            }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "6")
                //        {

                //            try
                //            {
                //                checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(216, 216, 216));
                //                lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]), setFontsAll(9, 1, 0));
                //                //lochunk.SetBackground(iTextSharp.text.Color.LIGHT_GRAY);
                //            }
                //            catch
                //            { }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "7")
                //        {
                //            if (Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]) != "")
                //            {
                //                try
                //                {
                //                    //loCell.EnableBorderSide(1);

                //                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                //                    lsFormatedString = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                    lochunk = new Chunk(lsFormatedString, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount])));
                //                }
                //                catch
                //                {

                //                }
                //            }
                //        }
                //        else if (Convert.ToString(lodataset.Tables[0].Rows[liRowCount]["_OrderNmb"]) == "8")
                //        {
                //            loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                //            //lochunk = new Chunk(abc, Font18Bold(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["Investment"])));
                //        }


                //    }



                //}
                //if (checkTrue(lodataset, liRowCount, "_OrderNmb") && !checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    if (liColumnCount == 0)
                //    {
                //        String abc = "          " + "          " + lodataset.Tables[0].Rows[liRowCount][0].ToString();
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(abc, Font8Grey());
                //        lochunk = new Chunk(abc, Font7Grey());
                //    }
                //    else
                //    {
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(lsFormatedString, Font8Grey());
                //        lochunk = new Chunk(lsFormatedString, Font7Grey());
                //    }
                //}

                ////CONDITION FOR SUPERBOLDFLAG
                ////checkTrue(lodataset, liRowCount, "_OrderNmb", loCell, new iTextSharp.text.Color(183, 221, 232));
                ////====added on 28Feb2011 to change font size for total====
                //if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    if (liColumnCount != 0)
                //    {
                //        lochunk = new Chunk(lsFormatedString, Font7Bold());
                //    }
                //}
                ///*=====END=====*/

                //if (checkTrue(lodataset, liRowCount, "_OrderNmb"))
                //{
                //    if (liColumnCount == 0)
                //    {
                //        String abc = "          " + "          " + "Total";
                //        //changed on 02/25/2011
                //        //lochunk = new Chunk(abc, Font8Normal());
                //        lochunk = new Chunk(abc, Font7Normal());
                //    }
                //    setTopWidthBlack(loCell);
                //    setBottomWidthWhite(loCell);

                //}
                #endregion
                loCell.Add(lochunk);
                loTable.AddCell(loCell);
            }

            try
            {
                if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
                {
                    document.Add(loTable);
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt, lsFooterLocation, ClientFooterTxt, Ssi_GreshamClientFooter));
                    SetTotalPageCount("Direct Mgr Detail Report");
                }
            }
            catch (Exception Ex)
            {

            }
        }
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

    public string generatePortfolioConstChartNEW()
    {
        string ReportPath = string.Empty;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;

        String lsSQL = getFinalSp(ReportType.PortfolioConsChartNEW);

        newdataset = clsDB.getDataSet(lsSQL);

        clsPorthfolioConsChartNew objPortConstChartNew = new clsPorthfolioConsChartNew();

        objPortConstChartNew.HouseHoldText = HouseHoldText;
        objPortConstChartNew.AllocationGroupText = AllocationGroup_Text;
        objPortConstChartNew.AsOfDate = AsOfDate;
        objPortConstChartNew.FooterText = FooterText;
        objPortConstChartNew.GreshamAdvisedFlag = GreshamAdvisedFlag;
        //objPortConstChartNew.RptVersion = PortFolioConChartRptVer;

        if (AllocationGroup_Text != "")
            lsFamiliesName = AllocationGroup_Text;
        else
            lsFamiliesName = HouseHoldText;

        objPortConstChartNew.lsFamiliesName = CommitmentReportHeader;

        if (PortFolioConChartRptVer == "old")
        {
            ReportPath = objPortConstChartNew.generatePDFFinal(newdataset, TempFolderPath);
        }
        else
        {
            ReportPath = objPortConstChartNew.generatePDFFinal_New(newdataset, TempFolderPath);
        }

        return ReportPath;
    }

    public string generatePortfolioConstChartV2()
    {
        string ReportPath = string.Empty;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;

        String lsSQL = getFinalSp(ReportType.PortfolioConsChartV2);

        newdataset = clsDB.getDataSet(lsSQL);

        clsPorthfolioConsChartNew objPortConstChartNew = new clsPorthfolioConsChartNew();

        objPortConstChartNew.HouseHoldText = HouseHoldText;
        objPortConstChartNew.AllocationGroupText = AllocationGroup_Text;
        objPortConstChartNew.AsOfDate = AsOfDate;
        objPortConstChartNew.FooterText = FooterText.Replace("v2.1", "");
        objPortConstChartNew.GreshamAdvisedFlag = GreshamAdvisedFlag;

        objPortConstChartNew.Footerlocation = Footerlocation;
        objPortConstChartNew.ClientFooterTxt = ClientFooterTxt;
        objPortConstChartNew.Ssi_GreshamClientFooter = Ssi_GreshamClientFooter;
        //objPortConstChartNew.RptVersion = PortFolioConChartRptVer;

        if (AllocationGroup_Text != "")
            lsFamiliesName = AllocationGroup_Text;
        else
            lsFamiliesName = HouseHoldText;

        objPortConstChartNew.lsFamiliesName = CommitmentReportHeader;

        if (PortFolioConChartRptVer == "old")
        {
            ReportPath = objPortConstChartNew.generatePDFFinal(newdataset, TempFolderPath);

        }
        else
        {
            ReportPath = objPortConstChartNew.generatePDFFinalV2(newdataset, TempFolderPath);

        }

        return ReportPath;
    }


    public string generatePerfAnalyticsRpt1()
    {
        #region REPORT 1

        string _AsOfDate = "";
        if (!string.IsNullOrEmpty(AsOfDate))
        {
            DateTime asofDT = Convert.ToDateTime(AsOfDate);
            _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);
        }

        Random rnd = new Random();

        string date2 = System.DateTime.Today.ToString();

        string strGUID = DateTime.Parse(date2).ToString("yyyyMMdd") + "_PERF_ANALYTICS_" + DateTime.Now.ToString("yyyy-MM-dd-HHmmssfff") + "_" + Convert.ToString(rnd.Next());
        // String fsFinalLocation = HttpContext.Current.Server.MapPath("~/ExcelTemplate/pdfOutput/" + strGUID + ".pdf");

        iTextSharp.text.Document pdoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate(), -23, -20, 43, 0);//10,10
                                                                                                                                //  String ls = HttpContext.Current.Server.MapPath("~/ExcelTemplate/pdfOutput/ls_" + strGUID + ".pdf");
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";

        PdfWriter writer = PdfWriter.GetInstance(pdoc, new FileStream(fsFinalLocation, FileMode.Create));

        pdoc.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        pdoc.Add(png);



        DB clsDB = new DB();

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




        Paragraph HeadingFamilyName = new Paragraph(lsFamiliesName.ToString().Replace("''", "'"), setFontsAll(14, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph HeadingRow2 = new Paragraph("TOTAL INVESTMENT ASSETS", setFontsAll(10, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph HeadingRow3 = new Paragraph("Are My Investment Assets Keeping Pace with Inflation?", setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
        Paragraph HeadingRow4 = new Paragraph(_AsOfDate, setFontsAll(10, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

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

        iTextSharp.text.Image chartimg1 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~") + @"\images\Gresham_Logo.png");
        iTextSharp.text.Image chartimg2 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~") + @"\images\Gresham_Logo.png");

        string filename1 = getLineChartReport1();
        string filename2 = filename1.Replace("OP_", "A_OP_");

        // string filename2 = filename1;
        //Document document = new Document();

        //JFreeChart chart = TEMP();
        //BufferedImage bufferedImage = chart.createBufferedImage(500, 500);
        //iTextSharp.text.Image image1 = iTextSharp.text.Image.GetInstance(bufferedImage);


        Paragraph PR1Row1Cell1 = new Paragraph("Total Investment Assets vs. Baseline", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        PR1Row1Cell1.SetAlignment("center");
        PR1Row1Cell1.SpacingAfter = 2f;
        PR1Row1Cell1.Leading = 10f;

        Paragraph PR1Row1Cell3 = new Paragraph("Total Investment Assets vs. Inflation Adj. Baseline", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
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



        DataSet dsTableRpt = clsDB.getDataSet(getFinalSp(ReportType.Rpt1Table));

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
        // Paragraph PTRRow2 = new Paragraph("as of " + _AsOfDate, setFontsAll(8, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));//commented - Basecamp request 7_10_2019
        Paragraph PTRRow3 = new Paragraph("Since        ", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow4Col1 = new Paragraph(" ", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow4Col2 = new Paragraph("Year-to Date", setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow4Col3 = new Paragraph("  " + InceptionDt, setFontsAll(8, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow5Col1 = new Paragraph("Beginning Value", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow5Col2 = new Paragraph(YearToDateMnyRow1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow5Col3 = new Paragraph(InceptionMnyRow1, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        // Paragraph PTRRow6Col1 = new Paragraph("Contributions/Withdrawals", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow6Col1 = new Paragraph("Contributions/(Withdrawals)", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));//changed - Basecamp request 7_10_2019
        Paragraph PTRRow6Col2 = new Paragraph(YearToDateMnyRow2, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow6Col3 = new Paragraph(InceptionMnyRow2, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        Paragraph PTRRow7Col1 = new Paragraph("Baseline", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow7Col2 = new Paragraph(YearToDateMnyRow3, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow7Col3 = new Paragraph(InceptionMnyRow3, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        // Paragraph PTRRow8Col1 = new Paragraph("Increase/Decrease", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow8Col1 = new Paragraph("Increase/(Decrease)", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));//changed - Basecamp request 7_10_2019
        Paragraph PTRRow8Col2 = new Paragraph(YearToDateMnyRow4, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow8Col3 = new Paragraph(InceptionMnyRow4, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        Paragraph PTRRow9Col1 = new Paragraph("Ending Value", setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow9Col2 = new Paragraph(YearToDateMnyRow5, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PTRRow9Col3 = new Paragraph(InceptionMnyRow5, setFontsAll(8, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

        PTRHeading.Leading = 10;
        // PTRRow2.Leading = 10;//commented - Basecamp request 7_10_2019
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
        // PTRRow2.SetAlignment("center");//commented - Basecamp request 7_10_2019
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
        // CellTRrow2.AddElement(PTRRow2);//commented - Basecamp request 7_10_2019
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
        Chunk PFooterRow1P2 = new Chunk(" Includes all investment assets, Gresham advised and non-Gresham advised. Figure depicts overall asset base after spending and taxes. ", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PFooterRow2P1 = new Chunk("Excludes residences, personal property and \"below the line\" assets.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PFooterRow3P1 = new Chunk("Baseline: ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PFooterRow3P2 = new Chunk("Beginning Total Investment Assets, adjusted for significant non-investment contributions or withdrawals, as appropriate.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PFooterRow4P1 = new Chunk("Infl. Adj. Initial TIA (Inflation Adjusted Initial Total Investment Assets): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PFooterRow4P2 = new Chunk("Depicts the value of your overall investable asset base if it remained intact (funds not withdrawn for non-investment needs) ", setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PFooterRow5P1 = new Chunk("and adjusted for inflation. ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PFooterRow6P1 = new Chunk("Infl. Adj. Baseline (Inflation Adjusted Baseline): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
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
        //   CellFooterRow4.AddElement(PFooterRow4);
        //CellFooterRow5.AddElement(PFooterRow5);
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
        //   LoFooter.AddCell(CellFooterRow4);
        LoFooter.AddCell(CellFooterRow5);
        LoFooter.AddCell(CellFooterRow6);
        LoFooter.AddCell(CellFooterRow7);

        LoFooter.WidthPercentage = 100f;
        LoFooter.TotalWidth = 100f;
        LoFooter.TotalWidth = 700f;


        /* Commented footer after requirement from Jeanne --28th June 2016
         *
         //LoFooter.WriteSelectedRows(0, 7, 55, 100, writer.DirectContent);
         *
         */

        //  loHeadingRow7.AddElement(LoFooter);
        //LoHeader.AddCell(loHeadingRow7);

        pdoc.Add(LoHeader);
        //  pdoc.Add(LoFooter);
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

        //Gresham Footer for Batch 
        //    PdfPTable gTable = addFooter(lsDateTime, 1, 1, liPageSize - 1, true, FooterText, "1", Footerlocation, ClientFooterTxt, Ssi_GreshamClientFooter);
        PdfPTable TabFooter = null;
        PdfPTable TabFooter1 = null;

        if (Ssi_GreshamClientFooter == "3")
        {

            TabFooter = addFooterClientGoal(lsDateTime, true, FooterText, Footerlocation, false, 0, 0, ClientFooterTxt, Ssi_GreshamClientFooter);
            TabFooter1 = addFooterClientGoal1(lsDateTime, true, FooterText, Footerlocation, false, 0, 0, ClientFooterTxt, Ssi_GreshamClientFooter);
        }
        else
        {
            TabFooter = addFooterClientGoal(lsDateTime, true, FooterText, Footerlocation, false, 0, 0, ClientFooterTxt, Ssi_GreshamClientFooter);
        }
        //  TabFooter.WidthPercentage = 100f;
        TabFooter.TotalWidth = 100f;
        //  TabFooter.TotalWidth = 830;

        SetTotalPageCount("Client Goals");

        if (Ssi_GreshamClientFooter == "3")
        {
            TabFooter1.WidthPercentage = 100f;
            //TabFooter.TotalWidth = 100f;
            TabFooter1.TotalWidth = 700f;
        }

        // loTable.SpacingAfter = 12f;
        // lsFooterLocation = "100000001";


        if (Ssi_GreshamClientFooter == "1")
        {
            if (Footerlocation == "100000001")
            {
                pdoc.Add(TabFooter);
            }
            else
            {
                TabFooter.WriteSelectedRows(0, 7, 55, 40, writer.DirectContent);
            }
        }

        else if (Ssi_GreshamClientFooter == "2")
        {
            TabFooter.WriteSelectedRows(0, 7, 55, 40, writer.DirectContent);
        }

        else if (Ssi_GreshamClientFooter == "3")
        {
            pdoc.Add(TabFooter);
            // TabFooter.WriteSelectedRows(0, 7, 55, 0, writer.DirectContent);
            TabFooter1.WriteSelectedRows(0, 7, 60, 40, writer.DirectContent);

            /// TabFooter.WriteSelectedRows()

        }

        else if (Ssi_GreshamClientFooter == "4")
        {
            TabFooter.WriteSelectedRows(0, 7, 55, 40, writer.DirectContent);
        }

        //  pdoc.Add(TabFooter);
        //gTable.WidthPercentage = 100f;
        //gTable.TotalWidth = 100f;
        //gTable.TotalWidth = 700f;
        //gTable.WriteSelectedRows(0, 7, 55, 40, writer.DirectContent);
        pdoc.Close();

        //added 18-05-2018 (Sasmit- Cleanup JUNKFILES)
        try
        {
            if (filename1 != "")
            {
                File.Delete(filename1);
            }
            if (filename2 != "")
            {
                File.Delete(filename2);
            }
        }
        catch (Exception ex)
        {
        }

        //try
        //{

        //    FileInfo loFile = new FileInfo(ls);
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch
        //{ }
        return fsFinalLocation.Replace(".xls", ".pdf");
        #endregion
    }


    public string generatePerfAnalyticsRpt3()
    {
        string _AsOfDate = "";
        if (!string.IsNullOrEmpty(AsOfDate))
        {
            DateTime asofDT = Convert.ToDateTime(AsOfDate);
            _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);
        }

        Random rnd = new Random();
        string date2 = System.DateTime.Today.ToString();
        //Commented 8_14_2019(Batch Mixup issue)
        // string strGUID = DateTime.Parse(date2).ToString("yyyyMMdd") + "_PERF_ANALYTICS_" + DateTime.Now.ToString("yyyy-MM-dd-HHmmssfff") + "_" + Convert.ToString(rnd.Next()); ;
        //  String fsFinalLocation = HttpContext.Current.Server.MapPath("~/ExcelTemplate/pdfOutput/" + strGUID + ".pdf");

        iTextSharp.text.Document pdoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate(), -23, -20, 43, 0);//10,10
                                                                                                                                //added 8_14_2019(Batch Mixup issue)                                                                                                                  // String ls = HttpContext.Current.Server.MapPath("~/ExcelTemplate/pdfOutput/ls_" + strGUID + ".pdf");
        string strGUID = Guid.NewGuid().ToString();
        string fsFinalLocation = Path.Combine(TempFolderPath + "\\" + "ls_" + strGUID + ".pdf");

        PdfWriter writer = PdfWriter.GetInstance(pdoc, new FileStream(fsFinalLocation, FileMode.Create));

        pdoc.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        pdoc.Add(png);

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

        //if Family name is greater 
        if (lsFamiliesName.Length >= 20)
        {
            int[] widthR3Row3Temp = { 1, 42, 10, 10, 10, 12, 10 };
            LoR3Row3Temp.SetWidths(widthR3Row3Temp);
        }
        else
        {
            int[] widthR3Row3Temp = { 1, 32, 10, 10, 10, 12, 20 };
            LoR3Row3Temp.SetWidths(widthR3Row3Temp);
        }


        int[] widthR3Footer = { 100 };
        LoR3LoFooter.SetWidths(widthR3Footer);

        LoR3Row3Temp.TotalWidth = 100f;
        LoR3Row3Temp.WidthPercentage = 100f;


        Paragraph PR3FamilyName = new Paragraph(lsFamiliesName.ToString().Replace("''", "'"), setFontsAll(14f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR3HeadingRow2 = new Paragraph("GRESHAM ADVISED ASSETS", setFontsAll(10f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR3HeadingRow3 = new Paragraph("How Have My Gresham Advised Assets Performed?", setFontsAll(12f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
        Paragraph PR3HeadingRow4 = new Paragraph(_AsOfDate, setFontsAll(10f, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

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

        loR3HeadingRow4.PaddingBottom = 10f;

        LoR3Header.AddCell(loR3FamilyName);
        LoR3Header.AddCell(loR3HeadingRow2);
        LoR3Header.AddCell(loR3HeadingRow3);
        LoR3Header.AddCell(loR3HeadingRow4);

        Paragraph PR3Row1Cell1 = new Paragraph("Growth of My Gresham Advised Assets (GAA)", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        PR3Row1Cell1.SetAlignment("center");
        PR3Row1Cell1.SpacingAfter = 2f;
        PR3Row1Cell1.Leading = 10f;

        Paragraph PR3Row3Cell1 = new Paragraph("Annual Performance of Gresham Advised Assets (GAA)", setFontsAll(9f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        PR3Row3Cell1.SetAlignment("center");
        PR3Row3Cell1.SpacingAfter = 2f;
        PR3Row3Cell1.Leading = 10f;

        iTextSharp.text.Image chartimg3 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~") + @"\images\Gresham_Logo.png");
        string filename3 = getLineChartReport3();
        chartimg3 = iTextSharp.text.Image.GetInstance(filename3);

        iTextSharp.text.Image chartimg4 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~") + @"\images\Gresham_Logo.png");
        // string filename4 = getBarChartTAB3();
        string filename4 = getBarChartReport3();
        chartimg4 = iTextSharp.text.Image.GetInstance(filename4);

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



        DataSet dsTableRpt3 = clsDB.getDataSet(getFinalSp(ReportType.Rpt3Table1));

        IncDate = dsTableRpt3.Tables[0].Rows[0]["InceptionDate"].ToString();
        DataSet dsTableCPI = clsDB.getDataSet(getFinalSp(ReportType.Rpt3Table2));
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
                //Commented 1_30_2019 - Basecamp Request
                // if (num1Year == 0)
                //    str1year = "N/A";
                //else
                str1year = num1Year.ToString("P1", CultureInfo.InvariantCulture);
            }
            else
                str1year = "N/A";

            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["3 Year"])))
            {
                double num3Year = Convert.ToDouble(dsTableRpt3.Tables[0].Rows[0]["3 Year"].ToString());
                //Commented 1_30_2019 - Basecamp Request
                //if (num3Year == 0)
                //  str3year = "N/A";
                //else
                str3year = num3Year.ToString("P1", CultureInfo.InvariantCulture);
            }
            else
                str3year = "N/A";


            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["5 Year"])))
            {
                double num5Year = Convert.ToDouble(dsTableRpt3.Tables[0].Rows[0]["5 Year"].ToString());
                //Commented 1_30_2019 - Basecamp Request              
                // if (num5Year == 0)
                //     str5year = "N/A";
                // else
                str5year = num5Year.ToString("P1", CultureInfo.InvariantCulture);
            }
            else
                str5year = "N/A";

            if (!string.IsNullOrEmpty(Convert.ToString(dsTableRpt3.Tables[0].Rows[0]["Since Inception"])))
            {
                double numSinceInc = Convert.ToDouble(dsTableRpt3.Tables[0].Rows[0]["Since Inception"].ToString());
                //Commented 1_30_2019 - Basecamp Request
                // if (numSinceInc == 0)
                //     strSinceInc = "N/A";
                // else
                strSinceInc = numSinceInc.ToString("P1", CultureInfo.InvariantCulture);
            }
            else
                strSinceInc = "N/A";

            //  FamName = dsTableRpt3.Tables[0].Rows[0][0].ToString();
            FamName = lsFamiliesName.Replace("''", "'") + " GAA";
            IncDate = dsTableRpt3.Tables[0].Rows[0]["InceptionDate"].ToString();
        }

        if (dsTableCPI.Tables[0].Rows.Count > 0)
        {

            if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["1 Year"])))
            {
                double num1Year = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["1 Year"].ToString());
                //Commented 1_30_2019 - Basecamp Request
                // if (num1Year == 0)
                //     str1yearCPI = "N/A";
                // else
                str1yearCPI = num1Year.ToString("N1", CultureInfo.InvariantCulture) + " %";
            }
            else
                str1yearCPI = "N/A";

            if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["3 Year"])))
            {
                double num3Year = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["3 Year"].ToString());
                //Commented 1_30_2019 - Basecamp Request
                // if (num3Year == 0)
                //     str3yearCPI = "N/A";
                //  else
                str3yearCPI = num3Year.ToString("N1", CultureInfo.InvariantCulture) + " %";
            }
            else
                str3yearCPI = "N/A";

            if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["5 Year"])))
            {
                double num5Year = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["5 Year"].ToString());
                //Commented 1_30_2019 - Basecamp Request
                // if (num5Year == 0)
                //    str5yearCPI = "N/A";
                // else
                str5yearCPI = num5Year.ToString("N1", CultureInfo.InvariantCulture) + " %";
            }
            else
                str5yearCPI = "N/A";

            if (!string.IsNullOrEmpty(Convert.ToString(dsTableCPI.Tables[0].Rows[0]["Since Inception"])))
            {
                double numSinceInc = Convert.ToDouble(dsTableCPI.Tables[0].Rows[0]["Since Inception"].ToString());
                //Commented 1_30_2019 - Basecamp Request
                //  if (numSinceInc == 0)
                //      strSinceIncCPI = "N/A";
                //  else
                strSinceIncCPI = numSinceInc.ToString("N1", CultureInfo.InvariantCulture) + " %";
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
        Chunk PR3FooterRow2P2 = new Chunk(" Total dollars invested, adjusted for contributions and withdrawals.", setFontsAll(7.5f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PR3FooterRow3P1 = new Chunk("Performance:", setFontsAll(7.5f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PR3FooterRow3P2 = new Chunk(" Performance is shown net of all manager fees but gross of Gresham's fee, which covers a wide range of services. ", setFontsAll(7.5f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
        Chunk PR3FooterRow4P1 = new Chunk("Inflation Adj. Net Invested Capital (Inflation Adjusted Net Invested Capital): ", setFontsAll(7.5f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#969696"))));
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
        LoR3LoFooter.AddCell(CellR3FooterRow4);
        LoR3LoFooter.AddCell(CellR3FooterRow3);
        // LoR3LoFooter.WidthPercentage = 100f;
        // LoR3LoFooter.TotalWidth = 100f;
        LoR3LoFooter.TotalWidth = 700;


        //  LoR3LoFooter.WriteSelectedRows(0, 3, 55, -100, writer.DirectContent);

        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        pdoc.Add(LoR3Header);
        pdoc.Add(LoR3Row1);
        pdoc.Add(LoR3Row2);
        pdoc.Add(LoR3Row3);

        /* Commented footer after requirement from Jeanne --28th June 2016
         * 
            pdoc.Add(LoR3LoFooter);
         * 
         */
        //Gresham Footer for Batch 
        PdfPTable gTable = addFooterAbsoluteReturn(lsDateTime, 1, 1, liPageSize - 3, true, FooterText, "1", Footerlocation, ClientFooterTxt, Ssi_GreshamClientFooter);
        pdoc.Add(gTable);

        pdoc.Close();

        //added 18-05-2018 (Sasmit- Cleanup JUNKFILES)
        try
        {
            if (filename3 != "")
            {
                File.Delete(filename3);
            }
            if (filename4 != "")
            {
                File.Delete(filename4);
            }
        }
        catch (Exception ex)
        {

        }
        SetTotalPageCount("Absolute Returns");

        //Commented 8_14_2019(Batch Mixup issue)
        //try
        //{

        //    FileInfo loFile = new FileInfo(ls);
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch
        //{ }

        #endregion

        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    public string generatePerfAnalyticsRpt4()
    {
        string _AsOfDate = "";
        if (!string.IsNullOrEmpty(AsOfDate))
        {
            DateTime asofDT = Convert.ToDateTime(AsOfDate);
            _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);
        }

        Random rnd = new Random();
        string date2 = System.DateTime.Today.ToString();

        string strGUID = DateTime.Parse(date2).ToString("yyyyMMdd") + "_PERF_ANALYTICS_" + DateTime.Now.ToString("yyyy-MM-dd-HHmmssfff") + "_" + Convert.ToString(rnd.Next());
        //String fsFinalLocation = HttpContext.Current.Server.MapPath("~/ExcelTemplate/pdfOutput/" + strGUID + ".pdf");

        iTextSharp.text.Document pdoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate(), -23, -20, 43, 0);//10,10
        //String ls = HttpContext.Current.Server.MapPath("~/ExcelTemplate/pdfOutput/ls_" + strGUID + ".pdf");
        String fsFinalLocation = TempFolderPath + "\\" + "ls_" + Guid.NewGuid().ToString() + ".pdf";

        PdfWriter writer = PdfWriter.GetInstance(pdoc, new FileStream(fsFinalLocation, FileMode.Create));

        pdoc.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        pdoc.Add(png);

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

        string GrpName = ReportRollupGroupIdName;

        string strAssetClass = AssetClassCSV;

        // LoR4Legends.SpacingBefore = 30f;

        // LoR4Row2.TotalWidth = 100f;
        // LoR4Row2.WidthPercentage = 100f;


        Paragraph PR4FamilyName = new Paragraph(lsFamiliesName.ToString().Replace("''", "'"), setFontsAll(14f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4HeadingRow2 = new Paragraph("GRESHAM ADVISED ASSETS", setFontsAll(10f, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
        Paragraph PR4HeadingRow3 = new Paragraph("How Has Gresham Reduced My Portfolio's Risk?", setFontsAll(12f, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
        Paragraph PR4HeadingRow4 = new Paragraph(_AsOfDate, setFontsAll(10f, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

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
        iTextSharp.text.Image chartimg6 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~") + @"\images\Gresham_Logo.png");
        //  string filename6 = GetChart5Left(out Mindate);
        string filename6 = getColumnChartReport4(lsFamiliesName);
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

        iTextSharp.text.Image chartimg5 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~") + @"\images\Gresham_Logo.png");
        string filename5 = getShapeChartReport4("4");
        chartimg5 = iTextSharp.text.Image.GetInstance(filename5);


        chartimg5.ScalePercent(25, 25);
        //chartimg5.ScalePercent(75, 75);
        //  chartimg6.ScalePercent(75, 75);

        //Get Inception date 
        DataSet dsInc = clsDB.getDataSet(getFinalSp(ReportType.Rpt3Table1));

        IncDate = Convert.ToString(dsInc.Tables[0].Rows[0]["InceptionDate"]);

        DateTime dtInc = DateTime.Parse(IncDate, new CultureInfo("en-US"));
        DateTime dtfixed = DateTime.Parse("12/31/2010", new CultureInfo("en-US"));
        string strR4Row1Cell1Heading = "";
        if (dtInc < dtfixed)
            strR4Row1Cell1Heading = "Performance vs. Volatility (since 01/01/2011)";
        else
            strR4Row1Cell1Heading = "Performance vs. Volatility (since " + dtInc.ToString("MM/dd/yyyy") + ")";



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

        iTextSharp.text.Image chartimg8 = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("~") + @"\images\Gresham_Logo.png");
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

        DataSet dsTableRpt5 = clsDB.getDataSet(getFinalSp(ReportType.Rpt4TableLT));

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

        /* Commented footer after requirement from Jeanne --28th June 2016
         * 
              LoR4LoFooter.WriteSelectedRows(0, 3, 55, -100, writer.DirectContent);
         * 
         */

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



        if (dtInc < dtfixed)
        {
            pdoc.Add(LoR4Header);
            pdoc.Add(LoR4Row1);
            pdoc.Add(LoR4Row2);
            pdoc.Add(LoR4Row3);
            /* Commented footer after requirement from Jeanne --28th June 2016
             * 
             pdoc.Add(LoR4LoFooter);
             * 
             */
        }
        else
        {
            //pdoc.NewPage();
            pdoc.Add(LoR4Header);
            pdoc.Add(LoR4Row1);
            pdoc.Add(LoR4Row2Short);
            pdoc.Add(LoR4Row3);
            /* Commented footer after requirement from Jeanne --28th June 2016
             * 
            pdoc.Add(LoR4LoFooter);
             */
        }


        #endregion

        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

        //Gresham Footer for Batch 
        PdfPTable gTable = addFooter(lsDateTime, 1, 1, liPageSize - 5, true, FooterText, "1", Footerlocation, ClientFooterTxt, Ssi_GreshamClientFooter);
        pdoc.Add(gTable);

        pdoc.Close();

        //added 18-05-2018 (Sasmit- Cleanup JUNKFILES)
        try
        {
            if (filename5 != "")
            {
                File.Delete(filename5);
            }
            if (filename6 != "")
            {
                File.Delete(filename6);
            }
            if (filename8 != "")
            {
                File.Delete(filename8);
            }
        }
        catch (Exception ex)
        {
        }
        SetTotalPageCount("Capital Protection");
        //try
        //{

        //    FileInfo loFile = new FileInfo(ls);
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch
        //{ }

        return fsFinalLocation.Replace(".xls", ".pdf");


    }

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
            qry = getFinalSp(clsCombinedReports.ReportType.Rpt4TableLT);

        else
            qry = getFinalSp(clsCombinedReports.ReportType.Rpt4TableST);
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


    private string getLineChartReport1()
    {
        System.Web.UI.DataVisualization.Charting.Chart LineChartNEW1 = new System.Web.UI.DataVisualization.Charting.Chart();
        LineChartNEW1.Height = 400;
        LineChartNEW1.Height = 800;
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
        LineChartNEW1.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
        LineChartNEW1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 1;
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 1;
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        LineChartNEW1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW1.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW1.ChartAreas[0].AxisX.LineWidth = 2;

        LineChartNEW1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        LineChartNEW1.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        LineChartNEW1.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
        LineChartNEW1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        LineChartNEW1.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW1.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 1;
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

        System.Web.UI.DataVisualization.Charting.Chart LineChartNEW2 = new System.Web.UI.DataVisualization.Charting.Chart();
        LineChartNEW2.Height = 400;
        LineChartNEW2.Height = 800;
        LineChartNEW2.BorderlineDashStyle = ChartDashStyle.Solid;
        LineChartNEW2.Visible = false;


        LineChartNEW2.Titles.Add(new System.Web.UI.DataVisualization.Charting.Title("Total Investment Assets vs. Inflation Adj. Baseline"));
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
        LineChartNEW2.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        LineChartNEW2.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW2.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW2.ChartAreas[0].AxisX.LineWidth = 2;

        LineChartNEW2.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        LineChartNEW2.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        LineChartNEW2.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
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
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();


        //String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";
        //String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\P_OP_" + strGUID + ".xls";

        DB clsDB = new DB();
        //string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        DataSet ds = clsDB.getDataSet(getFinalSp(ReportType.Rpt1LineChart));

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

        LineChartNEW2.DataSource = dtatble;
        LineChartNEW2.DataBind();

        // Set series chart type
        LineChartNEW1.Series[0].ChartType = SeriesChartType.Line;
        LineChartNEW1.Series[1].ChartType = SeriesChartType.Line;

        LineChartNEW2.Series[0].ChartType = SeriesChartType.Line;
        LineChartNEW2.Series[1].ChartType = SeriesChartType.Line;


        if (ds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Double s1val = 0.0;
                Double s2val = 0.0;
                Double s4val = 0.0;

                if (Convert.ToString(ds.Tables[0].Rows[i]["value"]) != "")
                    s1val = Convert.ToDouble(ds.Tables[0].Rows[i]["value"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                if (Convert.ToString(ds.Tables[0].Rows[i]["NetInvestments"]) != "")
                    s2val = Convert.ToDouble(ds.Tables[0].Rows[i]["NetInvestments"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                if (Convert.ToString(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"]) != "")
                    s4val = Convert.ToDouble(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                Dt1 = ds.Tables[0].Rows[i]["Date"].ToString();
                sDay = DateTime.Parse(Dt1).ToString("dd");
                sMonth = DateTime.Parse(Dt1).ToString("MM");
                sYear = DateTime.Parse(Dt1).ToString("yyyy");

                DateTime dtDate = Convert.ToDateTime(Dt1);

                LineChartNEW1.Series["Series1"].Points.AddXY(dtDate, s1val);
                LineChartNEW1.Series["Series2"].Points.AddXY(dtDate, s2val);

                LineChartNEW2.Series["Series1"].Points.AddXY(dtDate, s1val);
                LineChartNEW2.Series["Series2"].Points.AddXY(dtDate, s4val);

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


            //Set Max value of Y-Axis -- Right Line chart 
            if (max1 > max4)
                maxx2 = RoundToMax(max1);

            if (max4 > max1)
                maxx2 = RoundToMax(max4);


            //Max value of both graph to be same
            if (maxx1 > maxx2)
            {
                LineChartNEW1.ChartAreas[0].AxisY.Maximum = maxx1;
                LineChartNEW2.ChartAreas[0].AxisY.Maximum = maxx1;
            }
            else
            {
                LineChartNEW1.ChartAreas[0].AxisY.Maximum = maxx2;
                LineChartNEW2.ChartAreas[0].AxisY.Maximum = maxx2;
            }


            if (maxx1 > 5000000 && maxx1 < 60000000)
                LineChartNEW1.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            if (maxx2 > 5000000 && maxx2 < 60000000)
                LineChartNEW2.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            Double S1LastValue = 0.0;
            Double S2LastValue = 0.0;
            Double S3LastValue = 0.0;

            if (ds.Tables[0].Rows.Count > 1)
            {
                if (Convert.ToString(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["value"]) != "")
                    S1LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["value"].ToString());

                if (Convert.ToString(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["NetInvestments"]) != "")
                    S2LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["NetInvestments"].ToString());

                if (Convert.ToString(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"]) != "")
                    S3LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"].ToString());
            }

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
            LineChartNEW1.Series[1].Name = "Baseline";


            //LineChartNEW1.Series[0].BorderColor = S;
            LineChartNEW1.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW1.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);


            LineChartNEW1.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChartNEW1.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S2LastValue.ToString("C0");

            LineChartNEW1.Series[0].Points[ds.Tables[0].Rows.Count - 1].Font = new System.Drawing.Font("Frutiger55", 7F, System.Drawing.FontStyle.Bold);
            LineChartNEW1.Series[1].Points[ds.Tables[0].Rows.Count - 1].Font = new System.Drawing.Font("Frutiger55", 7F, System.Drawing.FontStyle.Bold);

            LineChartNEW1.Series[0].SmartLabelStyle.Enabled = true;
            LineChartNEW1.Series[0].SmartLabelStyle.Enabled = true;

            //Added 1st if condition to show labels with value below 5000000- ChartIssue Basecamp 4_8_2019
            if (S1LastValue <= 5000000 || S2LastValue <= 5000000)
            {
                LineChartNEW1.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                LineChartNEW1.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            }
            else if (S1LastValue > S2LastValue)
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
            LineChartNEW2.ChartAreas[0].AxisX.IsMarginVisible = false;
            LineChartNEW2.ChartAreas[0].AxisX.IsStartedFromZero = true;

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


            LineChartNEW2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);
            LineChartNEW2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);

            LineChartNEW2.Series[0].BorderWidth = 10;
            LineChartNEW2.Series[1].BorderWidth = 10;


            LineChartNEW2.Series[0].Color = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW2.Series[1].Color = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);


            LineChartNEW2.Series[0].Name = "Total Investment Assets (TIA)";
            LineChartNEW2.Series[1].Name = "Inflation Adj. Baseline";


            //LineChartNEW2.Series[0].BorderColor = S;
            LineChartNEW2.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW2.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);


            LineChartNEW2.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChartNEW2.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S3LastValue.ToString("C0");

            LineChartNEW2.Series[0].Points[ds.Tables[0].Rows.Count - 1].Font = new System.Drawing.Font("Frutiger55", 7F, System.Drawing.FontStyle.Bold);
            LineChartNEW2.Series[1].Points[ds.Tables[0].Rows.Count - 1].Font = new System.Drawing.Font("Frutiger55", 7F, System.Drawing.FontStyle.Bold);

            //Added 1st if condition to show labels with value below 5000000- ChartIssue Basecamp 4_8_2019
            if (S1LastValue <= 5000000 || S3LastValue <= 5000000)
            {
                LineChartNEW2.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                LineChartNEW2.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            }
            else if (S1LastValue > S3LastValue)
            {
                LineChartNEW2.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                LineChartNEW2.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }
            else
            {
                LineChartNEW2.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                LineChartNEW2.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            }

            //Remove Extra for label to display 
            LineChartNEW2.Series[0].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW2.Series[0].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChartNEW2.Series[0].SmartLabelStyle.CalloutLineWidth = 0;

            LineChartNEW2.Series[1].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW2.Series[1].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChartNEW2.Series[1].SmartLabelStyle.CalloutLineWidth = 0;

            if (ds.Tables[0].Rows.Count < 12) //1years --MONTHLY
            {
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else if (ds.Tables[0].Rows.Count >= 12 && ds.Tables[0].Rows.Count < 36) //Quaterly
            {
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.Interval = 3;
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else
            {
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Years;
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1; //To start with december
                LineChartNEW2.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Days;
            }


            LineChartNEW2.Series[0].IsVisibleInLegend = true;
            LineChartNEW2.Series[1].IsVisibleInLegend = true;


            LineChartNEW2.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChartNEW2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
        }


        //LineChartNEW1.ChartAreas[0].Position.X = 5;
        //LineChartNEW1.ChartAreas[0].Position.Y = 8;
        //LineChartNEW1.ChartAreas[0].Position.Height = 82;
        //LineChartNEW1.ChartAreas[0].Position.Width = 97;

        //LineChartNEW2.ChartAreas[0].Position.X = 5;
        //LineChartNEW2.ChartAreas[0].Position.Y = 8;
        //LineChartNEW2.ChartAreas[0].Position.Height = 82;
        //LineChartNEW2.ChartAreas[0].Position.Width = 97;


        Random rnd = new Random();
        string RNum = Convert.ToString(rnd.Next(999999999));
        string strGuid = Guid.NewGuid().ToString();
        //string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + RNum + ".bmp";
        //string filename1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_OP_" + RNum + ".bmp";
        string filename = TempFolderPath + "\\" + "OP_" + strGuid + ".bmp";
        string filename1 = TempFolderPath + "\\" + "A_OP_" + strGuid + ".bmp";
        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(1300, 800);
        Bitmap bm1 = new Bitmap(1300, 800);

        bm.SetResolution(300, 300);
        bm1.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);
        System.Drawing.Graphics gGraphics1 = System.Drawing.Graphics.FromImage(bm1);


        LineChartNEW1.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 1300, 800));
        LineChartNEW2.Paint(gGraphics1, new System.Drawing.Rectangle(0, 0, 1300, 800));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);
        bm1.Save(filename1, System.Drawing.Imaging.ImageFormat.Bmp);


        //  Chart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in LineChartNEW1.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }

        return filename;
    }

    private string getLineChartReport3()
    {

        System.Web.UI.DataVisualization.Charting.Chart LineChartNEW = new System.Web.UI.DataVisualization.Charting.Chart();
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
        //  LineChartNEW.Series.Add(new Series());// commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019


        LineChartNEW.ChartAreas.Add(new ChartArea());
        LineChartNEW.ChartAreas[0].BorderColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW.ChartAreas[0].BackSecondaryColor = System.Drawing.Color.Transparent;
        LineChartNEW.ChartAreas[0].BackColor = System.Drawing.Color.Transparent;
        LineChartNEW.ChartAreas[0].ShadowColor = System.Drawing.Color.Transparent;

        LineChartNEW.ChartAreas[0].AxisY.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW.ChartAreas[0].AxisY.LineWidth = 2;
        LineChartNEW.ChartAreas[0].AxisY.LabelStyle.Format = "{C0}";
        LineChartNEW.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        LineChartNEW.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        LineChartNEW.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        LineChartNEW.ChartAreas[0].AxisX.LineWidth = 2;

        LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        LineChartNEW.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        LineChartNEW.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
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

        //Commented 8_14_2019(Batch Mixup issue)
        //Random rand = new Random();
        //string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();


        //String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";
        //String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\P_2OP_" + strGUID + ".xls";

        DB clsDB = new DB();
        //string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        //string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        //string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
        DataSet ds = clsDB.getDataSet(getFinalSp(ReportType.Rpt3LineChart));

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
        // LineChartNEW.Series[2].ChartType = SeriesChartType.Line;// commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019

        //   double max1 = 0.0, max2 = 0.0, max3 = 0.0;// commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
        double max1 = 0.0, max2 = 0.0;
        double maxx1 = 0.0, minn1 = 0.0;
        double min1 = 0.0, min2 = 0.0;

        if (ds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {

                Double s1val = 0.0;
                Double s2val = 0.0;
                // Double s3val = 0.0;// commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019

                if (Convert.ToString(ds.Tables[0].Rows[i]["Gresham Advised Assets"]) != "")
                    s1val = Convert.ToDouble(ds.Tables[0].Rows[i]["Gresham Advised Assets"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                if (Convert.ToString(ds.Tables[0].Rows[i]["Net Invested Capital"]) != "")
                    s2val = Convert.ToDouble(ds.Tables[0].Rows[i]["Net Invested Capital"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

                // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
                //if (Convert.ToString(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"]) != "")
                //    s3val = Convert.ToDouble(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"].ToString(), System.Globalization.CultureInfo.InvariantCulture);


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
                // LineChartNEW.Series["Series3"].Points.AddXY(dtDate, s3val);// commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019

                if (i == 0)
                {
                    min1 = s1val;
                    min2 = s2val;
                }

                //To set max point on chart
                if (s1val > max1)
                    max1 = s1val;

                if (s2val > max2)
                    max2 = s2val;

                if (s1val < min1)
                    min1 = s1val;

                if (s2val < min2)
                    min2 = s2val;


                // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
                //    if (s3val > max3)
                //        max3 = s3val;
            }

            // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
            ////Set Max value of Y-Axis -- Left Line chart 
            //if (max1 > max2 && max1 > max3)
            //    maxx1 = RoundToMax(max1);

            //if (max2 > max1 && max2 > max3)
            //    maxx1 = RoundToMax(max2);

            //if (max3 > max1 && max3 > max2)
            //    maxx1 = RoundToMax(max3);

            // changed -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
            //Set Max value of Y-Axis -- Left Line chart 
            if (max1 > max2)
            {
                maxx1 = RoundToMax_AbsoulteReturns(max1);
            }
            else
            {
                maxx1 = RoundToMax_AbsoulteReturns(max2);

            }
            if (min1 < min2)
            {
                minn1 = FloorToMinimum_AbsoulteReturns(min1);
            }
            else
            {
                minn1 = FloorToMinimum_AbsoulteReturns(min2);
            }
            if (double.IsNaN(minn1))
            {
                minn1 = 0;
            }
            LineChartNEW.ChartAreas[0].AxisY.Minimum = minn1;
            if (maxx1 != 0.0)
            {
                LineChartNEW.ChartAreas[0].AxisY.Maximum = maxx1;
            }
            if (maxx1 > 5000000 && maxx1 < 50000000)
            {  // LineChartNEW.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;
            }
            Double S1LastValue = 0.0;
            Double S2LastValue = 0.0;
            //  Double S3LastValue = 0.0;

            if (ds.Tables[0].Rows.Count > 0)
            {
                if (Convert.ToString(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Gresham Advised Assets"]) != "")
                    S1LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Gresham Advised Assets"].ToString());

                if (Convert.ToString(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Net Invested Capital"]) != "")
                    S2LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Net Invested Capital"].ToString());

                // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
                //if (Convert.ToString(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"]) != "")
                //    S3LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"].ToString());
            }

            LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);
            LineChartNEW.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);

            LineChartNEW.Series[0].BorderWidth = 10;
            LineChartNEW.Series[1].BorderWidth = 10;
            // LineChartNEW.Series[2].BorderWidth = 10; // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019//

            LineChartNEW.Series[0].Color = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW.Series[1].Color = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);
            // LineChartNEW.Series[2].Color = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap); // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019

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
            //  LineChartNEW.Series[2].Name = "Inflation Adj. Net Invested Capital"; // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019

            #region added -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
            //To Show Labels of end points
            LineChartNEW.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChartNEW.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S2LastValue.ToString("C0");
            // LineChartNEW.Series[0]["LabelStyle"] = "Top";
            // LineChartNEW.Series[1]["LabelStyle"] = "Top";
            #endregion
            //LineChartNEW.Series[0].BorderColor = S;
            LineChartNEW.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChartNEW.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);
            // LineChartNEW.Series[2].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap); // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019

            //   LineChartNEW.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            //   LineChartNEW.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S2LastValue.ToString("C0");
            //   LineChartNEW.Series[2].Points[ds.Tables[0].Rows.Count - 1].Label = S3LastValue.ToString("C0");

            // commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
            //LineChartNEW.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label.PadLeft(200);
            //int S1 = 0, S2 = 0, S3 = 0;
            //double MaxPoint = 0.0, MinPoint = 0.0;
            //double[] values = { S1LastValue, S2LastValue, S3LastValue };
            //Array.Sort(values);
            //double minval = values[0];
            //double midval = values[1];
            //double maxval = values[2];


            /*
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
                                LineChartNEW.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                            else
                                LineChartNEW.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                        }
            */

            #region  commented -basecamp request to remove Inflation Adj. Net Invested Capital 7_10_2019
            //LineChartNEW.Series[2].SmartLabelStyle.IsMarkerOverlappingAllowed = true;
            //LineChartNEW.Series[1].SmartLabelStyle.IsMarkerOverlappingAllowed = true;
            //LineChartNEW.Series[0].SmartLabelStyle.IsMarkerOverlappingAllowed = true;

            //int PerTopY_to_max = 100 - (int)Math.Round((double)(100 * maxval) / maxx1);
            //int Permax_to_mid = 100 - (int)Math.Round((double)(100 * midval) / maxval);
            //int Permid_to_min = 100 - (int)Math.Round((double)(100 * minval) / midval);
            //int Permax_to_min = 100 - (int)Math.Round((double)(100 * minval) / maxval);
            //int Permin_to_minY = 100 - (int)Math.Round((double)(100 * 0) / minval);

            //int TopAnnValue = 0;
            //if (PerTopY_to_max <= 5)
            //    TopAnnValue = 4;
            //else if (PerTopY_to_max > 10 && PerTopY_to_max < 15)
            //    TopAnnValue = 7;
            //else if (PerTopY_to_max >= 15 && PerTopY_to_max < 25)
            //    TopAnnValue = 7;
            //else if (PerTopY_to_max >= 25 && PerTopY_to_max < 30)
            //    TopAnnValue = 10;
            //else if (PerTopY_to_max >= 30 && PerTopY_to_max < 40)
            //    TopAnnValue = 14;
            //else
            //    TopAnnValue = 17;


            //int midAnnValue = 0;
            //int minAnnValue = 0;
            ////Difference between maximum datapoint and mid datapoint is less then display after top value
            //if (Permax_to_mid <= 20)
            //    midAnnValue = TopAnnValue + 6;
            //else if (Permax_to_mid > 20 && Permax_to_mid < 30)
            //    midAnnValue = TopAnnValue + 12;
            //else if (Permax_to_mid >= 30 && Permax_to_mid < 40)
            //    midAnnValue = TopAnnValue + 15;
            //else if (Permax_to_mid >= 40 && Permax_to_mid < 50)
            //    midAnnValue = TopAnnValue + 18;
            //else if (Permax_to_mid >= 50 && Permax_to_mid < 55)
            //    midAnnValue = TopAnnValue + 24;
            //else
            //    midAnnValue = TopAnnValue + 18;

            //if (Permid_to_min <= 5)
            //    minAnnValue = midAnnValue + 10;
            //else if (Permid_to_min > 5 && Permid_to_min <= 10)
            //    minAnnValue = midAnnValue + 18;
            //else if (Permid_to_min > 10 && Permid_to_min <= 20)
            //    minAnnValue = midAnnValue + 14;
            //else if (Permid_to_min > 20 && Permid_to_min < 30)
            //    minAnnValue = midAnnValue + 15;
            //else if (Permid_to_min >= 30 && Permid_to_min < 40)
            //    minAnnValue = midAnnValue + 16;
            //else
            //    minAnnValue = midAnnValue + 15;

            //if (minAnnValue >= 51 && minAnnValue <= 55) //51 = X axis 
            //    minAnnValue = minAnnValue - 4;

            //System.Web.UI.DataVisualization.Charting.TextAnnotation TxtAnnMax = new System.Web.UI.DataVisualization.Charting.TextAnnotation();
            //TxtAnnMax.Text = maxval.ToString("C0");
            //TxtAnnMax.X = 90;
            //TxtAnnMax.Y = TopAnnValue;
            //TxtAnnMax.Font = new System.Drawing.Font("Frutiger55", 7, System.Drawing.FontStyle.Bold);
            //TxtAnnMax.ForeColor = System.Drawing.Color.Black;
            //LineChartNEW.Annotations.Add(TxtAnnMax);


            //System.Web.UI.DataVisualization.Charting.TextAnnotation TxtAnnMid = new System.Web.UI.DataVisualization.Charting.TextAnnotation();
            //TxtAnnMid.Text = midval.ToString("C0");
            //TxtAnnMid.X = 90;
            //TxtAnnMid.Y = midAnnValue;
            //TxtAnnMid.Font = new System.Drawing.Font("Frutiger55", 7, System.Drawing.FontStyle.Bold);
            //TxtAnnMid.ForeColor = System.Drawing.Color.Black;
            //LineChartNEW.Annotations.Add(TxtAnnMid);


            //System.Web.UI.DataVisualization.Charting.TextAnnotation TxtAnnMin = new System.Web.UI.DataVisualization.Charting.TextAnnotation();
            //TxtAnnMin.Text = minval.ToString("C0");
            //TxtAnnMin.X = 90;
            //TxtAnnMin.Y = minAnnValue;
            //TxtAnnMin.Font = new System.Drawing.Font("Frutiger55", 7, System.Drawing.FontStyle.Bold);
            //TxtAnnMin.ForeColor = System.Drawing.Color.Black;
            //LineChartNEW.Annotations.Add(TxtAnnMin);

            #endregion
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

            //LineChartNEW.Series[0].SmartLabelStyle.Enabled = false;
            //LineChartNEW.Series[1].SmartLabelStyle.Enabled = false;
            //LineChartNEW.Series[2].SmartLabelStyle.Enabled = false;

            //Remove Extra for label to display 
            LineChartNEW.Series[0].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW.Series[0].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.Black;
            LineChartNEW.Series[0].SmartLabelStyle.CalloutLineWidth = 1;
            //  LineChartNEW.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopRight;

            LineChartNEW.Series[1].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChartNEW.Series[1].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.Black;
            LineChartNEW.Series[1].SmartLabelStyle.CalloutLineWidth = 1;

            // LineChartNEW.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopRight;

            // commented - basecamp request to remove Inflation Adj.Net Invested Capital 7_10_2019
            //LineChartNEW.Series[2].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            //LineChartNEW.Series[2].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.Black;
            //LineChartNEW.Series[2].SmartLabelStyle.CalloutLineWidth = 1;
            //// LineChartNEW.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopRight;

            //  LineChartNEW.Series[2].SmartLabelStyle.IsOverlappedHidden = false;// commented - basecamp request to remove Inflation Adj.Net Invested Capital 7_10_2019
            LineChartNEW.Series[1].SmartLabelStyle.IsOverlappedHidden = false;
            LineChartNEW.Series[0].SmartLabelStyle.IsOverlappedHidden = false;

            LineChartNEW.Series[0].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;
            LineChartNEW.Series[1].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;
            // LineChartNEW.Series[2].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;// commented - basecamp request to remove Inflation Adj.Net Invested Capital 7_10_2019

            LineChartNEW.Series[0].IsVisibleInLegend = true;
            LineChartNEW.Series[1].IsVisibleInLegend = true;
            // LineChartNEW.Series[2].IsVisibleInLegend = true;// commented - basecamp request to remove Inflation Adj.Net Invested Capital 7_10_2019



            LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChartNEW.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
        }



        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        //  string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_" + RNum + ".bmp";
        string filename = TempFolderPath + "\\" + "A_" + Guid.NewGuid().ToString() + ".bmp";
        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(2600, 600);

        bm.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);

        LineChartNEW.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 2600, 600));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        //  Chart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in LineChartNEW.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;
    }


    private string getBarChartReport3()
    {


        System.Web.UI.DataVisualization.Charting.Chart BarChart1 = new System.Web.UI.DataVisualization.Charting.Chart();
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
        BarChart1.ChartAreas[0].AxisY.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        BarChart1.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        BarChart1.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        BarChart1.ChartAreas[0].AxisX.LineWidth = 2;
        BarChart1.ChartAreas[0].AxisX.Interval = 1;
        BarChart1.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy";
        BarChart1.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;
        BarChart1.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8f, FontStyle.Regular);

        BarChart1.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#CFCFCF");
        BarChart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        BarChart1.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;
        //System.Random rand = new System.Random();
        //string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();

        //String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";

        // JFreeChart chart = ChartFactory.createBarChart(
        DB clsDB = new DB();

        //string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        //string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        //string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

        DataSet ds = clsDB.getDataSet(getFinalSp(ReportType.Rpt3BarChart));

        string Dt1;
        string sDay;
        string sMonth;
        string sYear;

        DataTable dtatble = ds.Tables[0];

        BarChart1.DataSource = dtatble;
        BarChart1.DataBind();

        // Set series chart type
        BarChart1.Series[0].ChartType = SeriesChartType.Column;
        double[] _values = new double[ds.Tables[0].Rows.Count];
        if (ds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Double s1val = 0.0;
                string strReturn = Convert.ToString(ds.Tables[0].Rows[i]["Return"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Return"]);
                if (strReturn != "")
                {
                    s1val = Convert.ToDouble(strReturn, System.Globalization.CultureInfo.InvariantCulture);
                    _values[i] = s1val;
                }
                Dt1 = ds.Tables[0].Rows[i]["Dates"].ToString();
                sDay = "01";
                sMonth = "01";
                sYear = ds.Tables[0].Rows[i]["Dates"].ToString();

                //chart 1
                //   s1.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s1val);

                string d = Convert.ToInt16(sDay) + "/" + Convert.ToInt16(sMonth) + "/" + Convert.ToInt32(sYear);
                BarChart1.Series[0].Points.AddXY(sYear, s1val * 100);
            }
        }

        double maxValue = _values.Max() * 100;
        double minValue = _values.Min() * 100;


        BarChart1.ChartAreas[0].AxisY.Interval = 10;
        //   maxValue = maxValue + 10;
        //   minValue = minValue - 10;



        //    BarChart1.ChartAreas[0].AxisY.Maximum = maxValue;
        //   BarChart1.ChartAreas[0].AxisY.Minimum = minValue;

        maxValue = maxValue + (int)(maxValue * 0.25);
        if (minValue < 0)
            minValue = minValue + (int)(minValue * 0.25);
        else
            minValue = minValue - (int)(minValue * 0.25);

        BarChart1.ChartAreas[0].AxisY.Maximum = Math.Ceiling(maxValue / (Double)10) * 10;
        BarChart1.ChartAreas[0].AxisY.Minimum = Math.Floor(minValue / (Double)10) * 10;


        BarChart1.Series[0].IsValueShownAsLabel = true;
        BarChart1.Series[0].LabelFormat = "{0.0}%";
        BarChart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true; //disabled inner gridlines
        BarChart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false; //disabled inner gridlines
        BarChart1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 1;

        BarChart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false; //disabled inner gridlines
        BarChart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false; //disabled inner gridlines

        BarChart1.ChartAreas[0].AxisX.IsStartedFromZero = true;
        BarChart1.ChartAreas[0].AxisX.IsMarginVisible = true;
        //clsDB.getConfiguration();

        //System.Random rnd = new System.Random();
        //string RNum = Convert.ToString(rnd.Next(999999999));

        //string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_" + RNum + ".bmp";
        string filename = TempFolderPath + "\\" + "A_" + Guid.NewGuid().ToString() + ".bmp";
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

    private string getShapeChartReport4(string rptno)
    {
        System.Web.UI.DataVisualization.Charting.Chart ShapeChartRptNEW4 = new System.Web.UI.DataVisualization.Charting.Chart();
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
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineWidth = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.Size = 1;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;

        ShapeChartRptNEW4.ChartAreas[0].AxisX.LineColor = System.Drawing.ColorTranslator.FromHtml("#868686");
        ShapeChartRptNEW4.ChartAreas[0].AxisX.LineWidth = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.Title = "Annualized Volatility (Standard Deviation)";
        ShapeChartRptNEW4.ChartAreas[0].AxisX.TitleFont = new System.Drawing.Font("Frutiger55", 9, FontStyle.Bold);
        ShapeChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.Format = "{N0}%";
        ShapeChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.IsEndLabelVisible = true;

        ShapeChartRptNEW4.ChartAreas[0].AxisX.MajorGrid.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineColor = System.Drawing.ColorTranslator.FromHtml("#404040");
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineWidth = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.Size = 1;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.MinorTickMark.LineDashStyle = ChartDashStyle.Solid;


        System.Random rand = new System.Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";

        double Xmax = 0.0;
        double Ymax = 0.0;
        double axismax = 0.0;
        DB clsDB = new DB();
        //string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        //string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        //string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

        DataSet ds = clsDB.getDataSet(getFinalSp(ReportType.Rpt4ShapeChart));

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
        double[] valuesYaxis = { s1Y, s2Y, s3Y, s4Y };
        double maxval = values.Max();
        double minval = valuesYaxis.Min();

        double yaxismaxvalue = valuesYaxis.Max();
        yaxismaxvalue = Math.Ceiling(yaxismaxvalue);

        maxval = Math.Ceiling(maxval);

        if (maxval % 2 != 0)
            maxval++;

        ShapeChartRptNEW4.Series[0].Points.AddXY(s1X, s1Y);
        ShapeChartRptNEW4.Series[1].Points.AddXY(s2X, s2Y);
        ShapeChartRptNEW4.Series[2].Points.AddXY(s3X, s3Y);
        ShapeChartRptNEW4.Series[3].Points.AddXY(s4X, s4Y);

        ShapeChartRptNEW4.Series[0].Points[0].Label = dt.Rows[0]["Name"].ToString();
        ShapeChartRptNEW4.Series[1].Points[0].Label = dt.Rows[1]["Name"].ToString();
        ShapeChartRptNEW4.Series[2].Points[0].Label = dt.Rows[2]["Name"].ToString();
        ShapeChartRptNEW4.Series[3].Points[0].Label = dt.Rows[3]["Name"].ToString();

        // Set point chart type
        ShapeChartRptNEW4.Series[0].ChartType = SeriesChartType.Point;
        ShapeChartRptNEW4.Series[1].ChartType = SeriesChartType.Point;
        ShapeChartRptNEW4.Series[2].ChartType = SeriesChartType.Point;
        ShapeChartRptNEW4.Series[3].ChartType = SeriesChartType.Point;


        // Enable data points labels
        if (s1X == 0.0 && s1Y == 0.0)
            ShapeChartRptNEW4.Series[0]["LabelStyle"] = "Right";

        if (s2X == 0.0 && s2Y == 0.0)
            ShapeChartRptNEW4.Series[1]["LabelStyle"] = "Right";

        if (s3X == 0.0 && s3Y == 0.0)
            ShapeChartRptNEW4.Series[2]["LabelStyle"] = "Right";

        if (s4X == 0.0 && s4Y == 0.0)
            ShapeChartRptNEW4.Series[3]["LabelStyle"] = "Right";

        // ShapeChartRptNEW4.Series["Series1"].IsValueShownAsLabel = true;
        //  ShapeChartRptNEW4.Series["Series1"]["LabelStyle"] = "Center";

        // Set marker size
        ShapeChartRptNEW4.Series[0].MarkerSize = 10;
        ShapeChartRptNEW4.Series[1].MarkerSize = 10;
        ShapeChartRptNEW4.Series[2].MarkerSize = 10;
        ShapeChartRptNEW4.Series[3].MarkerSize = 10;

        ShapeChartRptNEW4.ChartAreas[0].Position.X = 0;
        ShapeChartRptNEW4.ChartAreas[0].Position.Y = 2;
        ShapeChartRptNEW4.ChartAreas[0].Position.Height = 95;
        ShapeChartRptNEW4.ChartAreas[0].Position.Width = 95;



        ShapeChartRptNEW4.ChartAreas[0].AxisX.Maximum = maxval;

        //  if (minval < 2) //If Annualized return is less than 0 
        //   ShapeChartRptNEW4.ChartAreas[0].AxisY.Maximum = 5;
        //  else
        ShapeChartRptNEW4.ChartAreas[0].AxisY.Maximum = yaxismaxvalue + 1;
        if (minval <= 2)
        {
            double minfloor = Math.Floor(minval);

            ShapeChartRptNEW4.ChartAreas[0].AxisY.Minimum = minfloor - 1;
        }
        //minval - (int)(Math.Abs(minval) * 0.50); ;

        // ShapeChartRptNEW4.ChartAreas[0].InnerPlotPosition.Height = 80;
        // ShapeChartRptNEW4.ChartAreas[0].InnerPlotPosition.Width = 80;

        // Set marker shape
        ShapeChartRptNEW4.Series[0].MarkerStyle = MarkerStyle.Square; //Total GAA
        ShapeChartRptNEW4.Series[1].MarkerStyle = MarkerStyle.Circle; //Marketable GAA
        ShapeChartRptNEW4.Series[2].MarkerStyle = MarkerStyle.Diamond; //Strategic Benchmark
        ShapeChartRptNEW4.Series[3].MarkerStyle = MarkerStyle.Triangle; //MSCI

        // Set marker color 
        ShapeChartRptNEW4.Series[0].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#548ACF");
        ShapeChartRptNEW4.Series[1].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#8064A2");
        ShapeChartRptNEW4.Series[2].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");
        ShapeChartRptNEW4.Series[3].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#17375E");

        // Set marker border -  Strategic Benchmark
        ShapeChartRptNEW4.Series[2].MarkerBorderWidth = 1;
        ShapeChartRptNEW4.Series[2].MarkerBorderColor = System.Drawing.ColorTranslator.FromHtml("#215968");

        ShapeChartRptNEW4.ChartAreas[0].AxisX.Interval = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisY.Interval = 2;
        ShapeChartRptNEW4.ChartAreas[0].AxisX.Minimum = 0.0;

        ShapeChartRptNEW4.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Regular);
        ShapeChartRptNEW4.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Regular);

        ShapeChartRptNEW4.Titles[0].Font = new System.Drawing.Font("Frutiger65", 9, FontStyle.Bold);
        ShapeChartRptNEW4.Titles[0].Docking = Docking.Top;
        ShapeChartRptNEW4.Titles[0].DockingOffset = -2;

        // Enable 3D
        // ShapeChartRptNEW4.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;

        Random rnd = new Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        //string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + RNum + ".bmp";
        string filename = TempFolderPath + "\\" + "OP_" + Guid.NewGuid().ToString() + ".bmp";

        Bitmap bm = new Bitmap(1280, 1000);
        bm.SetResolution(300, 300);
        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);
        ShapeChartRptNEW4.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 1280, 1000));
        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        foreach (var series in ShapeChartRptNEW4.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }

        return filename;


    }
    public void Write(DataTable dt)
    {
        int[] maxLengths = new int[dt.Columns.Count];

        for (int i = 0; i < dt.Columns.Count; i++)
        {
            maxLengths[i] = dt.Columns[i].ColumnName.Length;

            foreach (DataRow row in dt.Rows)
            {
                if (!row.IsNull(i))
                {
                    int length = row[i].ToString().Length;

                    if (length > maxLengths[i])
                    {
                        maxLengths[i] = length;
                    }
                }
            }
        }

        string val1 = string.Empty;
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            val1 = val1 + "|" + dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2);

            //  sw.Write(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2));
        }
        lg.AddinLogFile(LogFileName, val1);

        string val2 = string.Empty;
        foreach (DataRow row in dt.Rows)
        {
            val2 = "";
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (!row.IsNull(i))
                {
                    val2 = val2 + "|" + row[i].ToString().PadRight(maxLengths[i] + 2);
                    //  sw.Write(row[i].ToString().PadRight(maxLengths[i] + 2));
                }
                else
                {
                    val2 = val2 + "|" + new string(' ', maxLengths[i] + 2);
                    //  sw.Write(new string(' ', maxLengths[i] + 2));
                }
            }
            lg.AddinLogFile(LogFileName, val2);

        }



    }
    private string getColumnChartReport4(string Fname)
    {

        System.Web.UI.DataVisualization.Charting.Chart ColumnChartRptNEW4 = new System.Web.UI.DataVisualization.Charting.Chart();
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
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\CC_" + strGUID + ".xls";

        // JFreeChart chart = ChartFactory.createBarChart(
        DB clsDB = new DB();

        //string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        //string GrpName = ddlGroup.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlGroup.SelectedValue.Replace("'", "''") + "'";

        //string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";

        DataSet ds = clsDB.getDataSet(getFinalSp(ReportType.Rpt4ColumnChart));
        string sqlQuery1 = getFinalSp(ReportType.Rpt4ColumnChart);
        lg.AddinLogFile(LogFileName, sqlQuery1);
        string Dt1;
        string sDay;
        string sMonth;
        string sYear;

        DataTable dtatble = ds.Tables[0];

        ColumnChartRptNEW4.DataSource = dtatble;
        ColumnChartRptNEW4.DataBind();

        // Set series chart type
        ColumnChartRptNEW4.Series[0].ChartType = SeriesChartType.Column;
        ColumnChartRptNEW4.Series[1].ChartType = SeriesChartType.RangeColumn;
        if (ds.Tables[0].Rows.Count > 0)
        {

            Write(ds.Tables[0]);

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Double s1val = 0.0;
                Double s2val = 0.0;
                string strVal1 = Convert.ToString(ds.Tables[0].Rows[i]["Honore"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Honore"]);
                string strVal2 = Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]);

                if (strVal1 != "")
                    s1val = Convert.ToDouble(strVal1, System.Globalization.CultureInfo.InvariantCulture);

                if (strVal2 != "")
                    s2val = Convert.ToDouble(strVal2, System.Globalization.CultureInfo.InvariantCulture);

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

        if (ReportingName.ToString() != "")
            ColumnChartRptNEW4.Series[0].Name = ReportingName.ToString();
        else
            ColumnChartRptNEW4.Series[0].Name = Fname;

        ColumnChartRptNEW4.Series[2].Name = Convert.ToString(ds.Tables[0].Rows[0]["BenchMarkName"]);
        ColumnChartRptNEW4.Series[1].IsVisibleInLegend = false;

        //  ColumnChartRptNEW4.Series[0].BorderWidth = 5;
        // ColumnChartRptNEW4.Series[0].BorderColor = System.Drawing.Color.Aqua;


        ColumnChartRptNEW4.Series[0].XAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Primary;
        ColumnChartRptNEW4.Series[0].YAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Primary;

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
        ColumnChartRptNEW4.Series[1].XAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Secondary;
        ColumnChartRptNEW4.Series[1].YAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Secondary;


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

        // string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_" + RNum + ".bmp";
        string filename = TempFolderPath + "\\" + "A_" + Guid.NewGuid().ToString() + ".bmp";

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

    private string getMarketablePerf()
    {
        liPageSize = 37;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        //string spName = "EXEC SP_R_MGR_REPORT_NEW_PDF 'Gantz Family'  , '31-Mar-2013'  ";
        String lsSQL = getFinalSp(ReportType.RptMarketablePerf);
        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        if (newdataset.Tables[0].Rows.Count < 1 && ReportSource == "AdventInd")
        {
            return "No Record Found";
        }
        Random rand = new Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + rand.Next().ToString();


        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
                                                                                                                              // String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyHHmmss") + System.Guid.NewGuid().ToString() + "MarketablePerformance.pdf";
        String fsFinalLocation = TempFolderPath + "\\" + System.Guid.NewGuid().ToString() + "MarketablePerformance.pdf";
        PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));

        string lsFooterText = FooterText;//footer text is in below method

        string lsFooterLocation = Footerlocation;


        document.Open();

        //  String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        string HHName = CommitmentReportHeader.ToString().Replace("''", "'");
        //string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
        //if (strTitle != "")
        //    HHName = strTitle;

        string strheader = "MARKETABLE PERFORMANCE";
        string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";

        DateTime asofDT = Convert.ToDateTime(AsOfDate);

        string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);

        int rowsize = table.Rows.Count;
        if (rowsize == 0)
            rowsize = 1;
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(9, rowsize);   // 2 rows, 2 columns           
        lsTotalNumberofColumns = "9";
        iTextSharp.text.Cell loCell = new Cell();
        // setTableProperty(loTable);

        #region Table Style
        int[] headerwidths9 = { 29, 7, 8, 18, 10, 8, 8, 9, 9 };
        loTable.SetWidths(headerwidths9);
        loTable.Width = 100;

        loTable.Width = 100;

        loTable.Border = 0;
        loTable.Cellspacing = 0;
        loTable.Cellpadding = 3;
        loTable.Locked = false;

        #endregion

        iTextSharp.text.Paragraph lochunk = new Paragraph();
        iTextSharp.text.Chunk lochunknew = new Chunk();


        int colsize = 9;

        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        double total = (double)table.Rows.Count / liPageSize;
        int liTotalPage = Convert.ToInt32(Math.Ceiling(total));



        if (liTotalPage > 2)
        {
            if (liTotalPage == 3)
            {
                int temptotal = liPageSize + (liPageSize + 6);
                if (table.Rows.Count < temptotal)
                    liTotalPage = 2;
            }
            else
            {
                //  liPageSize = liPageSize + 7;
                int extra = liTotalPage - 2;
                // liPageSize = liPageSize + (extra * 7);
                total = (double)(table.Rows.Count + (extra * 7)) / liPageSize;
                liTotalPage = Convert.ToInt32(Math.Ceiling(total));
            }
        }
        //else if (liTotalPage == 2)
        //{
        //    int temptotal = liPageSize + (liPageSize + 7);
        //    if (table.Rows.Count <= temptotal)
        //        liTotalPage = 1;
        //}

        int liCurrentPage = 0;

        //if (table.Rows.Count % liPageSize != 0)
        //{
        //    liTotalPage = liTotalPage + 1;
        //}
        //else
        //{
        //    liPageSize = 30;
        //    liTotalPage = liTotalPage + 1;
        //}
        bool once = true;
        bool once1 = true;

        int liPageFixedSize = 37;

        for (int i = 0; i < rowsize; i++)
        {
            if (liCurrentPage == 0)
                liPageSize = 37;
            else if (liCurrentPage == liTotalPage - 1)
            {
                if (once)
                {
                    liPageSize = liPageFixedSize * (liCurrentPage + 1);
                    liPageSize = liPageSize + (5 * liCurrentPage);
                    once = false;
                }
            }
            else
            {
                if (once1)
                {
                    liPageSize = liPageFixedSize * (liCurrentPage + 1);
                    liPageSize = liPageSize + (6 * liCurrentPage);
                    once1 = false;
                }
            }

            //if last record of the page is ASSET then it will pushed to next page 
            if (liPageSize < table.Rows.Count)
            {
                if (Convert.ToString(table.Rows[liPageSize - 1]["FundOrder"]) == "0")
                    liPageSize = liPageSize - 1;
            }
            //if Calculated Pagesize matches with Total Record and Total is still less than Current page than add 2 records to next page.
            if (liPageSize >= table.Rows.Count && (liCurrentPage + 1) < liTotalPage && !(liPageSize > table.Rows.Count + 15))
                liPageSize = (liPageSize - (table.Rows.Count - liPageSize)) - 2;


            if (i % liPageSize == 0)
            {

                document.Add(loTable);

                if (i != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    // document.Add(addpageno("", liTotalPage, liCurrentPage, liPageSize, true, ""));
                    //  document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                    document.NewPage();
                    SetTotalPageCount("Marketable Performance");
                    once1 = true;
                }

                loTable = new iTextSharp.text.Table(9, rowsize);
                // int[] headerwidths = { 27, 7, 9, 18, 10, 8, 8, 8, 8 };
                loTable.SetWidths(headerwidths9);
                loTable.Width = 100;

                loTable.Width = 100;

                loTable.Border = 0;
                loTable.Cellspacing = 0;
                loTable.Cellpadding = 3;
                loTable.Locked = false;


                lochunk = new Paragraph(HHName, setFontsAllFrutiger(14, 1, 0));
                loCell = new Cell();
                loCell.Add(lochunk);
                loCell.Colspan = 9;
                loCell.HorizontalAlignment = 1;


                lochunk = new Paragraph(strheader, setFontsAllFrutiger(10, 0, 0));
                loCell.Add(lochunk);

                //    lochunk = new Paragraph(Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                // loCell.Add(lochunk);

                lochunk = new Paragraph(_AsOfDate + "\n", setFontsAllFrutiger(10, 0, 1));
                loCell.Add(lochunk);
                loCell.Border = 0;


                //Report Header only for first Page
                if (liCurrentPage == 0)
                    loTable.AddCell(loCell);

                //  if (liCurrentPage == liTotalPage - 1)
                //  document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, 70, true, FooterText, String.Empty));

                //  AddFooter(document, FooterText);
                //decimal FeePercent = Convert.ToDecimal(newdataset.Tables[1].Rows[0][0]);
                //FeePercent = Math.Round(FeePercent, 0);
                //lochunk = new Chunk("Returns are shown net of all manager fees, but gross of Gresham’s fee, which is currently " + FeePercent + " basis points.  Gresham’s fee covers a range of interrelated investment, planning and advisory services.", setFontsAllFrutiger(7, 1, 0));
                //loCell = new Cell();
                //loCell.Add(lochunk);
                //loCell.Colspan = 9;
                //loCell.HorizontalAlignment = 0;
                //loCell.Border = 0;
                //loCell.Leading = 7f;
                //loTable.AddCell(loCell);

                for (int k = 0; k < colsize; k++)
                {
                    if (k == 0)
                    {
                        lochunk = new Paragraph(" ", setFontsAll(8, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        loCell.Colspan = 5;
                        loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        if (k != 0)
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        else
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        loTable.AddCell(loCell);
                    }
                    else if (k == 5)
                    {

                        lochunk = new Paragraph("Short-Term Performance ", setFontsAll(8, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Colspan = 2;
                        loCell.Leading = 11F;
                        loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        loTable.AddCell(loCell);
                    }
                    else if (k == 7)
                    {

                        lochunk = new Paragraph("Long-Term Performance ", setFontsAll(8, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Colspan = 2;
                        loCell.Leading = 11F;
                        loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        loTable.AddCell(loCell);
                    }

                }

                for (int k = 0; k < colsize; k++)
                {
                    string ColHeader = Convert.ToString(table.Columns[k].ColumnName);
                    if (ColHeader == "SubFundNameTxt")
                        ColHeader = " ";
                    if (ColHeader == "InvestmentStrategyTxt")
                        ColHeader = "Investment Strategy";
                    if (ColHeader == "Ssi_MgrInitialInvestmentDate")
                        ColHeader = "Gresham Initial\nInvestment";
                    if (ColHeader == "QTD")
                        ColHeader = "QTD Return";
                    if (ColHeader == "YTD")
                        ColHeader = "YTD Return";
                    if (ColHeader == "5YrAnn")
                        ColHeader = "5 Yr Return ";
                    if (ColHeader == "5YrAnnVolatility")
                        ColHeader = "5 Yr Volatility ";
                    if (ColHeader == "TotalInvAssetPercent")
                        ColHeader = "% Total Inv. \n Assets";

                    //if (ColHeader.Contains("at"))
                    //    ColHeader = ColHeader.Replace("at", "at\n");

                    lochunk = new Paragraph(ColHeader, setFontsAll(8, 1, 0));
                    loCell = new iTextSharp.text.Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 11F;
                    loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                    //if (k != 0)
                    //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    //else
                    //  if (k == 2)
                    // loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    // else
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;


                    loTable.AddCell(loCell);
                }

                //Gresham Logo only for first page
                if (liCurrentPage == 0)
                {
                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                    png.SetAbsolutePosition(45, 557);//540
                    png.ScalePercent(10);
                    document.Add(png);
                }
            }

            if (table.Rows.Count > 0)
            {
                Chunk chunk2;
                for (int j = 0; j < colsize; j++)
                {
                    string ColValue = Convert.ToString(table.Rows[i][j]);
                    string ColorCode = Convert.ToString(table.Rows[i]["_ColourCode"]);
                    int BoldFlg = Convert.ToInt32(table.Rows[i]["_BoldFlg"]);
                    int ItalicsFlg = Convert.ToInt32(table.Rows[i]["_ItalicsFlg"]);
                    int UnderlineFlg = Convert.ToInt32(table.Rows[i]["_UnderlineFlg"]);
                    //if (ColValue=="Gresham Advised Strategies (excl. cash, fixed income)")
                    //    ColValue = "Gresham Advised Strategies\n(excl. cash, fixed income)";
                    //if (ColValue.Contains("(excl."))
                    //    ColValue = ColValue.Replace("(", "\n(");


                    if (Convert.ToString(table.Rows[i]["FundOrder"]) == "0" || i == 0)
                    {
                        if (ColValue == "N/A")
                            ColValue = "";

                        if ((j == 5 || j == 6 || j == 7 || j == 8) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                ColValue = ColValue.Replace("(", "($");
                            }
                            // else
                            //  ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                            BoldFlg = 1;
                            chunk2 = new Chunk(ColValue.ToUpper(), setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                            //if (UnderlineFlg == 1)
                            //chunk2.SetUnderline(1f, -1f);

                            lochunk = new Paragraph(chunk2);
                            if (j == 5 || j == 6)
                                lochunk.IndentationRight = 22f;
                            if (j == 7 || j == 8)
                                lochunk.IndentationRight = 24f;
                            if (j == 2)
                                lochunk.IndentationRight = 19f;
                        }
                        //else if (j == 8 && ColValue != "")
                        //{
                        //    ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                        //    lochunk = new Chunk(ColValue + "%", setFontsAll(7, 1, 0));
                        //}
                        //else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                        //{
                        //    if (ColValue != "N/A")
                        //        ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                        //    lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        //}
                        else
                        {
                            BoldFlg = 1;
                            if (i == 0 && j == 0)
                            {

                                string[] str = ColValue.Split('(');

                                chunk2 = new Chunk(str[0].ToUpper(), setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                if (UnderlineFlg == 1)
                                    chunk2.SetUnderline(1f, -2f);
                                lochunk = new Paragraph(chunk2);
                                if (str.Length > 1)
                                {
                                    lochunknew = new Chunk("(" + str[1].ToUpper(), setFontsAll(6, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                    if (UnderlineFlg == 1)
                                        lochunknew.SetUnderline(1f, -2f);
                                    lochunk = new Paragraph(lochunknew);
                                }
                            }
                            else
                            {
                                //  lochunk = new Paragraph(ColValue.ToUpper(), setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                if (UnderlineFlg == 1 && j == 0)
                                {
                                    chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                    chunk2.SetUnderline(1f, -2f);
                                }
                                else
                                    chunk2 = new Chunk(ColValue.ToUpper(), setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));

                                lochunk = new Paragraph(chunk2);
                            }
                        }

                        if (UnderlineFlg == 1)
                            lochunk.IndentationLeft = 20f;
                    }
                    else
                    {
                        if (j == 0) //component
                        {
                            chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                            if (UnderlineFlg == 1)
                                chunk2.SetUnderline(1f, -2f);
                            lochunk = new Paragraph(chunk2);
                            lochunk.IndentationLeft = 20f;
                        }
                        else
                        {
                            if ((j == 5 || j == 6 || j == 7 || j == 8) && ColValue != "")
                            {
                                //if (ColValue.Contains("-"))
                                //{
                                //    ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                //    ColValue = ColValue.Replace("(", "($");
                                //}
                                // else
                                //  ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                                chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                //if (UnderlineFlg == 1)
                                //    chunk2.SetUnderline(1f, -2f);

                                lochunk = new Paragraph(chunk2);
                                if (j == 5 || j == 6 || j == 2)
                                    lochunk.IndentationRight = 22f;
                                if (j == 7 || j == 8)
                                    lochunk.IndentationRight = 24f;
                            }
                            //else if (j == 8 && ColValue != "")
                            //{
                            //    lochunk = new Chunk(Convert.ToDecimal(ColValue) + "%", setFontsAll(7, 0, 0));
                            //}
                            //else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                            //{
                            //    if (ColValue != "N/A")
                            //        ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                            //    lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            //}
                            else
                            {
                                chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                //if (UnderlineFlg == 1)
                                //    chunk2.SetUnderline(1f, -2f);
                                lochunk = new Paragraph(chunk2);

                                if (j == 2)
                                    lochunk.IndentationRight = 19f;
                            }
                        }


                    }

                    loCell = new iTextSharp.text.Cell();
                    //if (!string.IsNullOrEmpty(ColorCode))
                    //    lochunk.Font = setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode)));
                    loCell.Add(lochunk);
                    //if ((i == 0 || i == 1) && j == 0)
                    //    loCell.Add(lochunknew);
                    loCell.Border = 0;

                    if (i == 0)
                        loCell.Leading = 8F;
                    else
                    {
                        if (Convert.ToString(table.Rows[i]["FundOrder"]) != "True" && Convert.ToString(table.Rows[i]["FundOrder"]) != "")
                            loCell.Leading = 2F;
                        else
                            loCell.Leading = 4F;
                    }

                    if (i % liPageSize == 0)
                    {
                        loCell.Leading = 8F;
                    }

                    //loCell.VerticalAlignment = 5;


                    if (j == 1 || j == 2 || j == 5 || j == 6 || j == 7 || j == 8)
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    //else if (j == 3)
                    //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    else if (j == 4)
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    else if (j == 3)
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //if (Convert.ToString(table.Rows[i]["Components"]) == "Gresham Advised Values")
                    //    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                    //else if (Convert.ToString(table.Rows[i]["Components"]) == "Marketable Strategies" || Convert.ToString(table.Rows[i]["Components"]) == "Non-Marketable Strategies")
                    //    loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;


                    loTable.AddCell(loCell);

                }
            }

            if (table.Rows.Count == 0)
                document.Add(loTable);
            if (i == table.Rows.Count - 1)
            {
                //bool PrevYr1ShowFlg = Convert.ToBoolean(table.Rows[0][15]);
                //bool PrevYr2ShowFlg = Convert.ToBoolean(table.Rows[0][16]);
                //if (PrevYr1ShowFlg == false && PrevYr2ShowFlg == false)
                //{
                //    loTable.DeleteColumn(2); //after deleting the second column 
                //    loTable.DeleteColumn(2);// the third column comes at second position
                //}
                //else if (PrevYr1ShowFlg == false)
                //    loTable.DeleteColumn(2);
                //else if (PrevYr2ShowFlg == false)
                //    loTable.DeleteColumn(3);
                document.Add(loTable);
                //document.Add(addpageno("", liTotalPage, liCurrentPage, 74, true, ""));
                //     liCurrentPage = liCurrentPage + 1;
                //table.Rows.Count % liPageSize
                if (liTotalPage < liCurrentPage + 1)
                    liTotalPage = liCurrentPage + 1;

                Paragraph PBlank = new Paragraph(" ");
                PdfPTable TabFooter = null;
                PdfPTable TabFooter1 = null;

                if (Ssi_GreshamClientFooter == "3")
                {

                    TabFooter = addFooterMGR(lsDateTime, true, lsFooterText, lsFooterLocation, true, liCurrentPage + 1, liTotalPage, ClientFooterTxt, Ssi_GreshamClientFooter);
                    TabFooter1 = addFooterMGR1(lsDateTime, true, lsFooterText, lsFooterLocation, true, liCurrentPage + 1, liTotalPage, ClientFooterTxt, Ssi_GreshamClientFooter);
                }
                else
                {
                    TabFooter = addFooterMGR(lsDateTime, true, lsFooterText, lsFooterLocation, true, liCurrentPage + 1, liTotalPage, ClientFooterTxt, Ssi_GreshamClientFooter);
                }
                SetTotalPageCount("Marketable Performance");
                TabFooter.WidthPercentage = 100f;
                //  TabFooter.TotalWidth = 100f;
                TabFooter.TotalWidth = 775;

                if (Ssi_GreshamClientFooter == "3")
                {
                    TabFooter1.WidthPercentage = 100f;
                    //  TabFooter.TotalWidth = 100f;
                    TabFooter1.TotalWidth = 775;
                }

                // loTable.SpacingAfter = 12f;
                // lsFooterLocation = "100000001";


                if (Ssi_GreshamClientFooter == "1")
                {
                    if (lsFooterLocation == "100000001")
                    {
                        document.Add(TabFooter);
                    }
                    else
                    {
                        if (lsFooterText.Contains("\n"))
                            // TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                            TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                        /// TabFooter.WriteSelectedRows()
                        else
                            TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                    }
                }

                else if (Ssi_GreshamClientFooter == "2")
                {
                    if (ClientFooterTxt.Contains("\n"))
                        TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                    /// TabFooter.WriteSelectedRows()
                    else
                        TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                }

                else if (Ssi_GreshamClientFooter == "3")
                {
                    document.Add(TabFooter);
                    if (lsFooterText.Contains("\n"))
                    {
                        if (Footerlocation == "100000000")
                            TabFooter1.WriteSelectedRows(0, 4, 30, 83, writer.DirectContent);
                        else
                            TabFooter1.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                    }
                    /// TabFooter.WriteSelectedRows()
                    else
                        TabFooter1.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                }

                else if (Ssi_GreshamClientFooter == "4")
                {
                    if (lsFooterText.Contains("\n"))
                        TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                    /// TabFooter.WriteSelectedRows()
                    else
                        TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                }


                // if (lsFooterText.Contains("\n"))
                //     TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                ///// TabFooter.WriteSelectedRows()
                // else
                //     TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);

                //   document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, 5, true, lsFooterText));


                //Chunk lochunk4 = new Chunk("\n Qualitative Summary", setFontsAllFrutiger(9, 1, 0));
                //Chunk lochunk5 = new Chunk("\n Gresham Advised Equity strategies and Gresham Advised Marketable Equity strategies will not include fixed income at this time.", setFontsAllFrutiger(7, 0, 0));
                //Paragraph p2 = new Paragraph();
                //p2.Add(lochunk4);
                //p2.Add(lochunk5);
                //document.Add(p2);

                //document.Add(addFooter("", liTotalPage, liCurrentPage, table.Rows.Count % liPageSize, true, String.Empty));
            }

        }

        if (table.Rows.Count == 0)
        {
            SetTotalPageCount("Marketable Performance");

        }

        document.Close();
        //if (table.Rows.Count > 0)
        //{
        //    document.Close();

        //    FileInfo loFile = new FileInfo(ls);
        //    try
        //    {
        //        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

        //        HttpContext.Current.Response.Write("<script>");
        //        string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
        //        HttpContext.Current.Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        //        HttpContext.Current.Response.Write("</script>");

        //    }
        //    catch (Exception exc)
        //    {
        //        HttpContext.Current.Response.Write(exc.Message);
        //    }
        //}

        //Commented 8_14_2019(Batch Mixup issue)
        //try
        //{

        //    FileInfo loFile = new FileInfo(ls);
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch
        //{ }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    //added 2_1_2019 Non Marketable (DYNAMO)
    public string NonMarketablePerf()
    {
        liPageSize = 37;// 3 places in current function(33-->37 changed on 6_6_2019)
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;

        String lsSQL = getFinalSp(ReportType.RptNonMarketablePerf);
        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        //DataTable table = newdataset.Tables[1].Copy();

        if (newdataset.Tables[0].Rows.Count < 1 && ReportSource == "AdventInd")
        {
            return "No Record Found";
        }


        Random rand = new Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + rand.Next().ToString();


        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
        //  String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyHHmmss") + System.Guid.NewGuid().ToString() + "NonMarketablePerformance.pdf";
        String fsFinalLocation = TempFolderPath + "\\" + System.Guid.NewGuid().ToString() + "NonMarketablePerformance.pdf";
        PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));

        string lsFooterText = FooterText;//footer text is in below method

        string lsFooterLocation = Footerlocation;
        document.Open();

        // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";
        for (int m = 0; m < newdataset.Tables.Count; m++)
        {
            // string Headertext = newdataset.Tables[m].Rows[0]["Header"].ToString();
            // string FundIRRflg = newdataset.Tables[m].Rows[0]["FundIRRFlg"].ToString();
            string FundIRRflg = "false";
            try
            {
                FundIRRflg = newdataset.Tables[m].Rows[0]["FundIRRFlg"].ToString();
            }
            catch (Exception ex)
            {
                FundIRRflg = "false";
            }
            // string FundIRRflg = newdataset.Tables[m].Rows[0]["FundIRRFlg"].ToString();
            // m++;
            //if (i != 0)
            //{
            //    document.NewPage();
            //}
            DataTable table = newdataset.Tables[m].Copy();

            string HHName = CommitmentReportHeader.ToString().Replace("''", "'");
            //string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
            //if (strTitle != "")
            //    HHName = strTitle;

            string strheader = ReportName;// Headertext;// "Non MARKETABLE PERFORMANCE";
            // string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";

            DateTime asofDT = Convert.ToDateTime(AsOfDate);

            string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);

            int rowsize = table.Rows.Count;
            if (rowsize == 0)
                rowsize = 1;
            //iTextSharp.text.Table loTable = new iTextSharp.text.Table(9, rowsize);   // 2 rows, 2 columns           
            //lsTotalNumberofColumns = "9";

            iTextSharp.text.Table loTable = new iTextSharp.text.Table(10, rowsize);   // 2 rows, 2 columns           
            lsTotalNumberofColumns = "10";


            iTextSharp.text.Cell loCell = new Cell();
            // setTableProperty(loTable);

            #region Table Style
            int[] headerwidths9 = { 32, 7, 9, 9, 8, 9, 9, 9, 9, 9 };

            //  int[] headerwidths9 = { 32, 8, 8, 8, 8, 9, 9, 9, 9 };
            loTable.SetWidths(headerwidths9);
            loTable.Width = 100;

            loTable.Width = 100;

            loTable.Border = 0;
            loTable.Cellspacing = 0;
            loTable.Cellpadding = 3;
            loTable.Locked = false;

            #endregion

            iTextSharp.text.Paragraph lochunk = new Paragraph();
            iTextSharp.text.Chunk lochunknew = new Chunk();


            // int colsize = 9;

            int colsize = 10;

            String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
            double total = (double)table.Rows.Count / liPageSize;
            int liTotalPage = Convert.ToInt32(Math.Ceiling(total));



            if (liTotalPage > 2)
            {
                if (liTotalPage == 3)
                {
                    int temptotal = liPageSize + (liPageSize + 6);
                    if (table.Rows.Count < temptotal)
                        liTotalPage = 2;
                }
                else
                {
                    //  liPageSize = liPageSize + 7;
                    int extra = liTotalPage - 2;
                    // liPageSize = liPageSize + (extra * 7);
                    total = (double)(table.Rows.Count + (extra * 5)) / liPageSize;
                    liTotalPage = Convert.ToInt32(Math.Ceiling(total));
                }
            }
            else if (liTotalPage == 2)
            {
                int temptotal = liPageSize + (liPageSize + 5);
                if (table.Rows.Count <= temptotal)
                    liTotalPage = 1;
            }

            int liCurrentPage = 0;

            //if (table.Rows.Count % liPageSize != 0)
            //{
            //    liTotalPage = liTotalPage + 1;
            //}
            //else
            //{
            //    liPageSize = 30;
            //    liTotalPage = liTotalPage + 1;
            //}
            bool once = true;
            bool once1 = true;

            int liPageFixedSize = 37;//(33-->37 changed on 6_6_2019)

            for (int i = 0; i < rowsize; i++)
            {
                if (liCurrentPage == 0)
                    liPageSize = 37;//(33-->37 changed on 6_6_2019)
                else if (liCurrentPage == liTotalPage - 1)
                {
                    if (once)
                    {
                        liPageSize = liPageFixedSize * (liCurrentPage + 1);
                        liPageSize = liPageSize + (3 * liCurrentPage); // Number of rows used up by header-3
                        once = false;
                    }
                }
                else
                {
                    if (once1)
                    {
                        liPageSize = liPageFixedSize * (liCurrentPage + 1);
                        liPageSize = liPageSize + (3 * liCurrentPage);
                        once1 = false;
                    }
                }

                //if last record of the page is ASSET then it will pushed to next page 
                if (liPageSize < table.Rows.Count)
                {
                    if (Convert.ToString(table.Rows[liPageSize - 1]["Fund"]) == "0")
                        liPageSize = liPageSize - 1;
                }
                //if Calculated Pagesize matches with Total Record and Total is still less than Current page than add 2 records to next page.
                if (liPageSize >= table.Rows.Count && (liCurrentPage + 1) < liTotalPage && !(liPageSize > table.Rows.Count + 15))
                    liPageSize = (liPageSize - (table.Rows.Count - liPageSize)) - 2;


                if (i % liPageSize == 0)
                {

                    document.Add(loTable);

                    if (i != 0)
                    {
                        liCurrentPage = liCurrentPage + 1;
                        // document.Add(addpageno("", liTotalPage, liCurrentPage, liPageSize, true, ""));
                        //  document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                        document.NewPage();
                        SetTotalPageCount("Non Marketable Performance");
                        once1 = true;
                    }

                    // loTable = new iTextSharp.text.Table(9, rowsize);

                    loTable = new iTextSharp.text.Table(10, rowsize);



                    // int[] headerwidths = { 27, 7, 9, 18, 10, 8, 8, 8, 8 };
                    loTable.SetWidths(headerwidths9);
                    loTable.Width = 100;

                    loTable.Width = 100;

                    loTable.Border = 0;
                    loTable.Cellspacing = 0;
                    loTable.Cellpadding = 3;
                    loTable.Locked = false;

                    #region OLD
                    ////lochunk = new Paragraph("Gresham Partners, LLC Private Equity Investments(August 31, 2016)", setFontsAllFrutiger(14, 1, 0));

                    ////loCell = new Cell();
                    ////loCell.Add(lochunk);
                    ////loCell.Colspan = 9;
                    ////loCell.Border = 0;
                    ////loCell.HorizontalAlignment = 1;

                    //lochunk = new Paragraph(Headertext, setFontsAllFrutiger(14, 1, 0));
                    //loCell = new Cell();
                    //loCell.Add(lochunk);
                    //loCell.Border = 0;
                    //loCell.Colspan = 9;
                    //loCell.HorizontalAlignment = 1;
                    //if (liCurrentPage == 0)
                    //{
                    //    loTable.AddCell(loCell);

                    //}
                    //lochunk = new Paragraph(HHName, setFontsAllFrutiger(10, 1, 0));
                    //loCell = new Cell();
                    //loCell.Add(lochunk);
                    //loCell.Colspan = 9;
                    //loCell.Border = 0;
                    //// loCell.BorderWidthTop = 1;
                    //loCell.HorizontalAlignment = 1;
                    //if (liCurrentPage == 0)
                    //{
                    //    loTable.AddCell(loCell);

                    //}
                    #endregion
                    lochunk = new Paragraph(HHName, setFontsAllFrutiger(14, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    //  loCell.Colspan = 9;

                    loCell.Colspan = 10;
                    loCell.HorizontalAlignment = 1;


                    lochunk = new Paragraph(strheader.ToUpper(), setFontsAllFrutiger(10, 0, 0));
                    loCell.Add(lochunk);

                    //    lochunk = new Paragraph(Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                    // loCell.Add(lochunk);

                    lochunk = new Paragraph(_AsOfDate + "\n", setFontsAllFrutiger(10, 0, 1));
                    loCell.Add(lochunk);
                    loCell.Border = 0;


                    //Report Header only for first Page
                    if (liCurrentPage == 0)
                        loTable.AddCell(loCell);

                    //  if (liCurrentPage == liTotalPage - 1)
                    //  document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, 70, true, FooterText, String.Empty));

                    //  AddFooter(document, FooterText);
                    //decimal FeePercent = Convert.ToDecimal(newdataset.Tables[1].Rows[0][0]);
                    //FeePercent = Math.Round(FeePercent, 0);
                    //lochunk = new Chunk("Returns are shown net of all manager fees, but gross of Gresham's fee, which is currently " + FeePercent + " basis points.  Gresham? fee covers a range of interrelated investment, planning and advisory services.", setFontsAllFrutiger(7, 1, 0));
                    //loCell = new Cell();
                    //loCell.Add(lochunk);
                    //loCell.Colspan = 9;
                    //loCell.HorizontalAlignment = 0;
                    //loCell.Border = 0;
                    //loCell.Leading = 7f;
                    //loTable.AddCell(loCell);


                    //Report Header
                    for (int k = 0; k < colsize; k++)
                    {
                        string ColHeader = Convert.ToString(table.Columns[k].ColumnName);

                        if (ColHeader.ToLower() == "fund")
                            ColHeader = "";


                        //                        if (ColHeader.ToLower() == "net irr")
                        if (ColHeader.ToLower() == "irr")
                        {
                            if (FundIRRflg.ToLower() == "true")
                            {
                                ColHeader = ColHeader + "*";
                            }
                        }
                        lochunk = new Paragraph(ColHeader, setFontsAll(8, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        if (k == 0)
                        {
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        }
                        else
                        {
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        }
                        // loCell.BorderWidthTop = 1;
                        // loCell.BorderWidthBottom = 1;
                        //if (k == 0)
                        //{
                        //    loCell.BorderWidthLeft = 1;
                        //}
                        //if (k == colsize - 1)
                        //{
                        //    loCell.BorderWidthRight = 1;
                        //}
                        loTable.AddCell(loCell);
                    }

                    //Gresham Logo only for first page
                    if (liCurrentPage == 0)
                    {
                        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                        png.SetAbsolutePosition(45, 557);//540
                        png.ScalePercent(10);
                        document.Add(png);
                    }
                }

                if (table.Rows.Count > 0)
                {
                    Chunk chunk2;
                    for (int j = 0; j < colsize; j++)
                    {
                        string ColorCode = string.Empty;
                        string ColValue = Convert.ToString(table.Rows[i][j]);
                        ColorCode = "#000000";
                        int BoldFlg = Convert.ToInt32(table.Rows[i]["BoldFlg"]);
                        int totalflg = Convert.ToInt32(table.Rows[i]["Totalflg"]);
                        int ItalicsFlg = 0;
                        // int UnderlineFlg = 0;
                        if (totalflg == 1)
                        {
                            ColorCode = Convert.ToString(table.Rows[i]["ColorCodeTxt"]);
                            // lochunk.IndentationLeft = 20f;
                        }


                        if (ColValue == "N/A" || ColValue == "")
                            ColValue = "";



                        if (j == 0) //component
                        {
                            chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));

                            lochunk = new Paragraph(chunk2);
                            if (BoldFlg == 0)
                            {
                                lochunk.IndentationLeft = 20f;
                            }


                            else if (BoldFlg == 1 && totalflg == 1)
                            {
                                lochunk.IndentationLeft = 20f;
                            }
                        }
                        else
                        {
                            // if ((j == 2 || j == 3 || j == 5 || j == 6) && ColValue != "")

                            if ((j == 2 || j == 3 || j == 5 || j == 7) && ColValue != "")
                            {

                                if (ColValue != "")
                                {
                                    if (ColValue.Contains("-"))
                                    {
                                        ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                        ColValue = ColValue.Replace("(", "($");
                                    }
                                    else
                                        ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                }
                                chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                //if (UnderlineFlg == 1)
                                //    chunk2.SetUnderline(1f, -2f);

                                lochunk = new Paragraph(chunk2);
                                //  if (j == 2 || j == 3 || j == 5 || j == 6)
                                if (j == 2 || j == 3 || j == 5 || j == 7)
                                    lochunk.IndentationRight = 12f;



                                //if (j == 4)
                                //    lochunk.IndentationRight = 20f;
                                //if (j == 5 || j == 6 || j == 2)
                                //    lochunk.IndentationRight = 22f;
                                //if (j == 7 || j == 8)
                                //    lochunk.IndentationRight = 24f;
                            }
                            else if (j == 9 || j == 4 || j == 6)
                            {
                                if (ColValue != "")
                                {
                                    ColValue = ColValue + " %";
                                    chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                    lochunk = new Paragraph(chunk2);

                                    lochunk.IndentationRight = 19f;

                                }
                                else
                                {
                                    chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                    lochunk = new Paragraph(chunk2);
                                }
                            }
                            else
                            {

                                chunk2 = new Chunk(ColValue, setFontsAll(7, BoldFlg, ItalicsFlg, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(ColorCode))));
                                //if (UnderlineFlg == 1)
                                //    chunk2.SetUnderline(1f, -2f);
                                lochunk = new Paragraph(chunk2);

                                lochunk.IndentationRight = 19f;


                                //if (j == 2)
                                //lochunk.IndentationRight = 19f;
                            }
                        }


                        // }

                        loCell = new iTextSharp.text.Cell();


                        loCell.Add(lochunk);


                        if (i == 0)
                            loCell.Leading = 8F;
                        else
                        {
                            if (Convert.ToString(table.Rows[i]["Fund"]) != "True" && Convert.ToString(table.Rows[i]["Fund"]) != "")
                                loCell.Leading = 2F;
                            else
                                loCell.Leading = 4F;
                        }


                        loCell.Border = 0;

                        if (i % liPageSize == 0)
                        {
                            loCell.Leading = 8F;
                        }

                        //loCell.VerticalAlignment = 5;


                        if (j == 1)
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //  else if (j == 2 || j == 3 || j == 5 || j == 6 || j == 4 || j == 7 || j == 8)

                        else if (j == 2 || j == 3 || j == 5 || j == 7 || j == 4 || j == 8 || j == 9 || j == 6)
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        //loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        // else if (j == 1)
                        // loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        //else if (j == 8)
                        //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;

                        loTable.AddCell(loCell);


                    }
                }

                if (table.Rows.Count == 0)
                    document.Add(loTable);
                if (i == table.Rows.Count - 1)
                {

                    document.Add(loTable);

                    if (liTotalPage < liCurrentPage + 1)
                        liTotalPage = liCurrentPage + 1;

                    PdfPTable TabFooter = null;
                    PdfPTable TabFooter1 = null;

                    if (Ssi_GreshamClientFooter == "3")
                    {

                        TabFooter = addFooterMGR(lsDateTime, true, lsFooterText, lsFooterLocation, true, liCurrentPage + 1, liTotalPage, ClientFooterTxt, Ssi_GreshamClientFooter);
                        TabFooter1 = addFooterMGR1(lsDateTime, true, lsFooterText, lsFooterLocation, true, liCurrentPage + 1, liTotalPage, ClientFooterTxt, Ssi_GreshamClientFooter);
                    }
                    else
                    {
                        TabFooter = addFooterMGR(lsDateTime, true, lsFooterText, lsFooterLocation, true, liCurrentPage + 1, liTotalPage, ClientFooterTxt, Ssi_GreshamClientFooter);
                    }
                    SetTotalPageCount("Non Marketable Performance");
                    TabFooter.WidthPercentage = 100f;
                    //  TabFooter.TotalWidth = 100f;
                    TabFooter.TotalWidth = 775;

                    if (Ssi_GreshamClientFooter == "3")
                    {
                        TabFooter1.WidthPercentage = 100f;
                        //  TabFooter.TotalWidth = 100f;
                        TabFooter1.TotalWidth = 775;
                    }

                    // loTable.SpacingAfter = 12f;
                    // lsFooterLocation = "100000001";


                    if (Ssi_GreshamClientFooter == "1")
                    {
                        if (lsFooterLocation == "100000001")
                        {
                            document.Add(TabFooter);
                        }
                        else
                        {
                            if (lsFooterText.Contains("\n"))
                                // TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                                TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                            /// TabFooter.WriteSelectedRows()
                            else
                                TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                        }
                    }

                    else if (Ssi_GreshamClientFooter == "2")
                    {
                        if (ClientFooterTxt.Contains("\n"))
                            TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                        /// TabFooter.WriteSelectedRows()
                        else
                            TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                    }

                    else if (Ssi_GreshamClientFooter == "3")
                    {
                        document.Add(TabFooter);
                        if (lsFooterText.Contains("\n"))
                        {
                            if (Footerlocation == "100000000")
                                TabFooter1.WriteSelectedRows(0, 4, 30, 83, writer.DirectContent);
                            else
                                TabFooter1.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                        }
                        /// TabFooter.WriteSelectedRows()
                        else
                            TabFooter1.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                    }

                    else if (Ssi_GreshamClientFooter == "4")
                    {
                        if (lsFooterText.Contains("\n"))
                            TabFooter.WriteSelectedRows(0, 4, 30, 43, writer.DirectContent);
                        /// TabFooter.WriteSelectedRows()
                        else
                            TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                    }


                }

            }
            if (table.Rows.Count == 0)
            {
                SetTotalPageCount("Marketable Performance");

            }
            document.NewPage();
        }

        //if (table.Rows.Count == 0)
        //{
        //    SetTotalPageCount("Marketable Performance");

        //}

        document.Close();


        //try
        //{

        //    FileInfo loFile = new FileInfo(ls);
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch
        //{ }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }


    private string getAssetsPerformanceSummary()
    {
        //liPageSize = 26; --> Original Value
        liPageSize = 30;

        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        //int liPageSize = 18;
        //int liCurrentPage = 0;
        //lstAssetClass.s

        String lsSQL = getFinalSp(ReportType.RptAssetsPerformanceSummary);
        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        Random rand = new Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + rand.Next().ToString();


        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
                                                                                                                              // String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyHHmmss") + System.Guid.NewGuid().ToString() + "AssetPerformanceReport.pdf";
        String fsFinalLocation = TempFolderPath + "\\" + System.Guid.NewGuid().ToString() + "AssetPerformanceReport.pdf";
        PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));

        string lsFooterText = FooterText;//footer text is in below method
        string lsFooterLocation = Footerlocation;//footer Location is in below method
        document.Open();

        /* Commented footer after requirement from Jeanne --28th June 2016
         * 
        AddRpt5Footer(document, FooterText, writer);
         * 
         */

        // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        string HHName = lsFamiliesName.ToString().Replace("''", "'");
        string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
        if (strTitle != "")
            HHName = strTitle;

        string strheader = "GRESHAM ADVISED ASSETS";
        string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";

        DateTime asofDT = Convert.ToDateTime(AsOfDate);

        string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);


        iTextSharp.text.Table loTable = new iTextSharp.text.Table(9, table.Rows.Count);   // 2 rows, 2 columns           
        lsTotalNumberofColumns = "9";
        iTextSharp.text.Cell loCell = new Cell();


        #region Table Style
        int[] headerwidths9 = { 27, 9, 9, 9, 12, 12, 12, 12, 10 };
        loTable.SetWidths(headerwidths9);
        loTable.Width = 100;
        loTable.Border = 0;
        loTable.Cellspacing = 0;
        loTable.Cellpadding = 3;
        loTable.Locked = false;

        #endregion

        iTextSharp.text.Chunk lochunk = new Chunk();
        iTextSharp.text.Chunk lochunknew = new Chunk();

        int rowsize = table.Rows.Count;
        int colsize = 9;

        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        int liTotalPage = (table.Rows.Count / liPageSize);
        int liCurrentPage = 0;

        if (table.Rows.Count % liPageSize != 0)
        {
            liTotalPage = liTotalPage + 1;
        }
        else
        {
            //liPageSize = 26; --> Original Value
            liPageSize = 30;
            liTotalPage = liTotalPage + 1;
        }

        for (int i = 0; i < rowsize; i++)
        {
            if (i % liPageSize == 0)
            {
                document.Add(loTable);

                if (i != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty, String.Empty, String.Empty, String.Empty));
                    document.NewPage();
                    SetTotalPageCount("Short Term Performance");
                }
                loTable = new iTextSharp.text.Table(9, table.Rows.Count);
                //  int[] headerwidths = { 27, 9, 9, 9, 12, 12, 12, 12, 10 };
                loTable.SetWidths(headerwidths9);
                loTable.Width = 100;

                loTable.Width = 100;

                loTable.Border = 0;
                loTable.Cellspacing = 0;
                loTable.Cellpadding = 3;
                loTable.Locked = false;

                lochunk = new Chunk(HHName, setFontsAllFrutiger(14, 1, 0));
                loCell = new Cell();
                loCell.Add(lochunk);
                loCell.Colspan = 9;
                loCell.HorizontalAlignment = 1;


                lochunk = new Chunk("\n" + strheader, setFontsAllFrutiger(10, 0, 0));
                loCell.Add(lochunk);

                lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                loCell.Add(lochunk);

                lochunk = new Chunk("\n" + _AsOfDate + "\n", setFontsAllFrutiger(10, 0, 1));
                loCell.Add(lochunk);
                loCell.Border = 0;
                loTable.AddCell(loCell);

                decimal FeePercent = Convert.ToDecimal(newdataset.Tables[1].Rows[0][0]);
                FeePercent = Math.Round(FeePercent, 0);
                // lochunk = new Chunk("Returns are shown net of all manager fees, but gross of Gresham’s fee, which is currently " + FeePercent + " basis points.  Gresham’s fee covers a range of interrelated investment, planning and advisory services.", setFontsAllFrutiger(7, 1, 0)); //commented 5_8_2019_ Basecamp Request
                lochunk = new Chunk("Returns are shown net of all manager fees, but gross of Gresham’s fee, which was " + FeePercent + " basis points as of the above date.  Gresham’s fee covers a range of interrelated investment, planning and advisory services.", setFontsAllFrutiger(7, 1, 0));
                loCell = new Cell();
                loCell.Add(lochunk);
                loCell.Colspan = 9;
                loCell.HorizontalAlignment = 0;
                loCell.Border = 0;
                loCell.Leading = 7f;
                loTable.AddCell(loCell);

                for (int k = 0; k < colsize; k++)
                {
                    string ColHeader = Convert.ToString(table.Columns[k].ColumnName);
                    if (ColHeader == "Contributions/Withdrawals")
                        ColHeader = "Contributions/\nWithdrawals";
                    //if (ColHeader.Contains("at"))
                    //    ColHeader = ColHeader.Replace("at", "at\n");

                    lochunk = new Chunk(ColHeader, setFontsAll(8, 1, 0));
                    loCell = new iTextSharp.text.Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 11F;
                    loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                    if (k != 0)
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    loTable.AddCell(loCell);
                }

                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                png.SetAbsolutePosition(45, 557);//540
                png.ScalePercent(10);
                document.Add(png);
            }
            for (int j = 0; j < colsize; j++)
            {
                string cellBackgroundColor = Convert.ToString(table.Rows[i]["ColourCode"]);
                string ColValue = Convert.ToString(table.Rows[i][j]);
                //if (ColValue=="Gresham Advised Strategies (excl. cash, fixed income)")
                //    ColValue = "Gresham Advised Strategies\n(excl. cash, fixed income)";
                //if (ColValue.Contains("(excl."))
                //    ColValue = ColValue.Replace("(", "\n(");
                if (ColValue == "Strategy Benchmark")
                    ColValue = "Weighted Benchmark";
                if (ColValue == "Gresham Advised Values")
                    ColValue = "Gresham Advised";

                if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True" || Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "" || i == 0 || i == 1)
                {
                    if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                    {
                        if (ColValue.Contains("-"))
                        {
                            ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                            ColValue = ColValue.Replace("(", "($");
                        }
                        else
                            ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    }
                    else if (j == 8 && ColValue != "")
                    {
                        ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                        lochunk = new Chunk(ColValue + "%", setFontsAll(7, 1, 0));
                    }
                    else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                    {
                        if (ColValue != "N/A")
                            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    }
                    else
                    {
                        if ((i == 0 || i == 1) && j == 0)
                        {
                            string[] str = ColValue.Split('(');

                            lochunk = new Chunk(str[0], setFontsAll(7, 1, 0));
                            if (str.Length > 1)
                                lochunknew = new Chunk("\n(" + str[1], setFontsAll(6, 1, 0));
                        }
                        else if (j == 0 && cellBackgroundColor != "")
                        {
                            lochunk = new Chunk("   " + ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (j == 0 && (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True"))
                        {
                            lochunk = new Chunk("       " + ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    }
                }
                else
                {
                    if (j == 0) //component
                        lochunk = new Chunk("           " + ColValue, setFontsAll(7, 0, 1));
                    else
                    {
                        if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                ColValue = ColValue.Replace("(", "($");
                            }
                            else
                                ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        }
                        else if (j == 8 && ColValue != "")
                        {
                            lochunk = new Chunk(Convert.ToDecimal(ColValue) + "%", setFontsAll(7, 0, 0));
                        }
                        else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                        {
                            if (ColValue != "N/A")
                                ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        }
                        else
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        }
                    }
                }

                loCell = new iTextSharp.text.Cell();
                loCell.Add(lochunk);
                if ((i == 0 || i == 1) && j == 0)
                    loCell.Add(lochunknew);
                loCell.Border = 0;

                if (i == 0 || i == 1 || i == 4)
                    loCell.Leading = 8F;
                else
                {
                    if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "True" && Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "")
                        loCell.Leading = 0F;
                    else
                        loCell.Leading = 4F;
                }

                //loCell.VerticalAlignment = 5;
                if (j != 0)
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Gresham Advised Values")
                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                else if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")
                    loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;

                if (Convert.ToString(table.Rows[i]["ColourCode"]) != "" || (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components"))
                //if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Growth Strategies" ||
                //    Convert.ToString(table.Rows[i]["Overall Performance"]) == "Economic Hedges" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components" ||
                //    Convert.ToString(table.Rows[i]["Overall Performance"]) == "Risk Reduction Strategies" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Test")
                {
                    loCell.Leading = 7f;
                    loCell.Colspan = 9;
                    //loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                }
                if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components")
                {
                    loCell.BorderColorBottom = iTextSharp.text.Color.WHITE;
                    loCell.BorderWidthBottom = 1.5f;
                }

                if (cellBackgroundColor != "")
                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(cellBackgroundColor));

                if ((Convert.ToString(table.Rows[i]["ColourCode"]) != "" || (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")) && j != 0)
                { }
                else
                {
                    loTable.AddCell(loCell);
                }
            }

            if (i == table.Rows.Count - 1)
            {
                bool PrevYr1ShowFlg = Convert.ToBoolean(table.Rows[0][15]);
                bool PrevYr2ShowFlg = Convert.ToBoolean(table.Rows[0][16]);
                if (PrevYr1ShowFlg == false && PrevYr2ShowFlg == false)
                {
                    loTable.DeleteColumn(2); //after deleting the second column 
                    loTable.DeleteColumn(2);// the third column comes at second position
                }
                else if (PrevYr1ShowFlg == false)
                    loTable.DeleteColumn(2);
                else if (PrevYr2ShowFlg == false)
                    loTable.DeleteColumn(3);

                document.Add(loTable);
                Paragraph PBlank = new Paragraph(" ");
                liCurrentPage = liCurrentPage + 1;

                PdfPTable TabFooter = null;
                PdfPTable TabFooter1 = null;

                if (Ssi_GreshamClientFooter == "3")
                {

                    TabFooter = addFooterAsset(lsDateTime, true, lsFooterText, lsFooterLocation, false, 0, 0, ClientFooterTxt, Ssi_GreshamClientFooter);
                    TabFooter1 = addFooterAsset1(lsDateTime, true, lsFooterText, lsFooterLocation, false, 0, 0, ClientFooterTxt, Ssi_GreshamClientFooter);
                }
                else
                {
                    TabFooter = addFooterAsset(lsDateTime, true, lsFooterText, lsFooterLocation, false, 0, 0, ClientFooterTxt, Ssi_GreshamClientFooter);
                }
                SetTotalPageCount("Short Term Performance");
                TabFooter.WidthPercentage = 100f;
                //  TabFooter.TotalWidth = 100f;
                TabFooter.TotalWidth = 775;

                if (Ssi_GreshamClientFooter == "3")
                {
                    TabFooter1.WidthPercentage = 100f;
                    //  TabFooter.TotalWidth = 100f;
                    TabFooter1.TotalWidth = 775;
                }

                // loTable.SpacingAfter = 12f;
                // lsFooterLocation = "100000001";


                if (Ssi_GreshamClientFooter == "1")
                {
                    if (lsFooterLocation == "100000001")
                    {
                        document.Add(TabFooter);
                    }
                    else
                    {
                        TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                    }
                }

                else if (Ssi_GreshamClientFooter == "2")
                {
                    TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                }

                else if (Ssi_GreshamClientFooter == "3")
                {
                    document.Add(TabFooter);
                    TabFooter1.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);

                    /// TabFooter.WriteSelectedRows()

                }

                else if (Ssi_GreshamClientFooter == "4")
                {
                    TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
                }
                //  document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, 12, true, lsFooterText));
                //Chunk lochunk4 = new Chunk("\n Qualitative Summary", setFontsAllFrutiger(9, 1, 0));
                //Chunk lochunk5 = new Chunk("\n Gresham Advised Equity strategies and Gresham Advised Marketable Strategies Components Equity strategies will not include fixed income at this time.", setFontsAllFrutiger(7, 0, 0));
                //Paragraph p2 = new Paragraph();
                //p2.Add(lochunk4);
                //p2.Add(lochunk5);
                //document.Add(p2);

                //document.Add(addFooter("", liTotalPage, liCurrentPage, table.Rows.Count % liPageSize, true, String.Empty));
            }
        }


        if (table.Rows.Count > 0)
        {
            document.Close();

            // FileInfo loFile = new FileInfo(ls);
            //try
            //{

            //    FileInfo loFile = new FileInfo(ls);
            //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            //}
            //catch
            //{ }

        }
        return fsFinalLocation.Replace(".xls", ".pdf");
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
    private double RoundToMax_AbsoulteReturns(double maxfromdb)
    {
        try
        {
            double maxfromdb_excess = 0.0;
            // bool bMax = false;
            // double maxxx = 0.0;
            double maxx = 0.0;
            maxfromdb = Math.Round(maxfromdb);
            maxfromdb_excess = Math.Round(maxfromdb * 0.12);

            //if (Convert.ToInt64(maxfromdb).ToString().Length == Convert.ToInt64(maxfromdb + maxfromdb_excess).ToString().Length)
            //{
            //    maxx = maxfromdb;
            //    bMax = false;
            //}
            //else
            //{
            //    maxx = maxfromdb + maxfromdb_excess;
            //    bMax = true;
            //}
            maxx = maxfromdb;
            //  maxx = 0;
            // int length = Convert.ToInt64(maxfromdb + maxx).ToString().Length;
            int length = Convert.ToInt64(maxx).ToString().Length; // check the Length of the vakue to decide the Scale value

            int num = 0;
            int percentage = 0;
            if (length == 6)
            {
                num = 25000;

                percentage = 2;
            }
            if (length == 7)
            {
                num = 250000;
                percentage = 4;
            }
            if (length == 8)
            {
                num = 2500000;
                percentage = 6;
            }
            if (length == 9)
            {
                num = 25000000;
                percentage = 8;
            }
            if (length == 10)
            {
                num = 250000000;
                percentage = 10;
            }



            // double i = Math.Ceiling(maxxx / (Double)num) * num;
            double i = Math.Ceiling(maxx / (Double)num) * num; // get the Max Upper Limit for the Y axis

            double percentofvalue = (percentage * maxfromdb) / 100; // the Max and the Limit shouldnt be very close so a percent value is considered according to the Scale Value(num) and checked 
            if (i <= percentofvalue + maxfromdb)
            {
                // i = i + num /2;
                i = i + num; // add a scale value if the max falls in the percentage range of the Limit Max
            }
            else
            {
                //  i = i - num / 2;
            }
            i = Math.Ceiling(i / (Double)5000000) * 5000000; //Ceil to the nearest 5 million

            return i;
        }
        catch (Exception ex)
        {
            return 0.0;
        }
    }
    private double FloorToMinimum_AbsoulteReturns(double minfromdb)
    {
        try
        {
            double minfromdb_excess = 0.0;
            bool bMax = false;

            double minn = 0.0;
            minfromdb = Math.Round(minfromdb);
            minfromdb_excess = Math.Round(minfromdb * 0.12);

            //if (Convert.ToInt64(minfromdb).ToString().Length == Convert.ToInt64(minfromdb - minfromdb_excess).ToString().Length)
            //{
            //    minn = minfromdb;
            //    bMax = false;
            //}
            //else
            //{
            //    minn = minfromdb - minfromdb_excess;
            //    bMax = true;
            //}
            minn = minfromdb;
            //  maxx = 0;
            // int length = Convert.ToInt64(maxfromdb + maxx).ToString().Length;
            int length = Convert.ToInt64(minn).ToString().Length;
            // length++;
            int num = 0;
            int percentage = 0;
            if (length == 6)
            {
                num = 25000;
                percentage = 2;
            }

            if (length == 7)
            {
                num = 250000;
                percentage = 4;
            }

            if (length == 8)
            {
                num = 2500000;

                //percentage = 6;
                //percentage = 8;		  //Changed Basecamp Issue 31_10_2019
                percentage = 12;		  //Changed Basecamp Issue 06_11_2019
            }
            if (length == 9)
            {
                num = 25000000;

                //percentage = 8;
                //percentage = 10;	    //Changed Basecamp Issue 31_10_2019
                percentage = 14;	    //Changed Basecamp Issue 06_11_2019
            }
            if (length == 10)
            {
                num = 250000000;

                //percentage = 10;
                //percentage = 12;		  //Changed Basecamp Issue 31_10_2019
                percentage = 16;		  //Changed Basecamp Issue 06_11_2019
            }

            // double i = Math.Ceiling(maxxx / (Double)num) * num;
            double i = Math.Floor(minn / (Double)num) * num; // fetch the Min Limit

            double percentofvalue = (percentage * minfromdb) / 100;// the Min and the Limit shouldnt be very close so a percent value is considered according to the Scale Value(num) and checked 
            if (i <= minfromdb - percentofvalue)
            {
                //i = i - num / 2;
                // i = i - num -(num/2); //substract the scale by 1.5% i.e [min = min-scale - (scale/2)]
            }
            else
            {
                //  i = i - num / 2;
                i = i - num; //substract the scale by 1
            }

            i = Math.Floor(i / (Double)5000000) * 5000000;//Floor to the Nearest 5 million

            return i;
        }
        catch (Exception ex)
        {
            return 0.0;
        }
    }
    public void AddRpt5Footer(iTextSharp.text.Document document, string FooterText, PdfWriter writer)
    {
        //Table -- Footer
        PdfPTable LoFooter = new PdfPTable(1);

        int[] widthFooter = { 100 };
        LoFooter.SetWidths(widthFooter);

        LoFooter.TotalWidth = 100f;

        LoFooter.WidthPercentage = 100f;

        PdfPCell CellFooterRow1 = new PdfPCell();
        PdfPCell CellFooterRow2 = new PdfPCell();
        PdfPCell CellFooterRow3 = new PdfPCell();
        PdfPCell CellFooterRow4 = new PdfPCell();
        PdfPCell CellFooterRow5 = new PdfPCell();
        PdfPCell CellFooterRow6 = new PdfPCell();
        PdfPCell CellFooterRow7 = new PdfPCell();
        PdfPCell CellFooterRow8 = new PdfPCell();
        PdfPCell CellFooterRow9 = new PdfPCell();

        Phrase PFooterRow1 = new Phrase();
        Phrase PFooterRow2 = new Phrase();
        Phrase PFooterRow3 = new Phrase();
        Phrase PFooterRow4 = new Phrase();
        Phrase PFooterRow5 = new Phrase();
        Phrase PFooterRow6 = new Phrase();
        Phrase PFooterRow7 = new Phrase();
        Phrase PFooterRow8 = new Phrase();
        Phrase PFooterRow9 = new Phrase();

        Chunk PFooterRow1P1 = new Chunk("Gresham Advised Assets (GAA): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150)));
        Chunk PFooterRow1P2 = new Chunk("All Gresham advised investments except cash (includes private strategies - private real assets and private equity).", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        Chunk PFooterRow1P3 = new Chunk("Gresham Advised Marketable Assets (Marketable GAA): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150)));
        Chunk PFooterRow1P4 = new Chunk("All Gresham advised investments except cash and private strategies.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        Chunk PFooterRow1P5 = new Chunk("Weighted Benchmark: ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150)));
        Chunk PFooterRow1P6 = new Chunk("The average of the benchmark return for each asset class, weighted by that asset class' percentage of total marketable GAA each month.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        //Chunk PFooterRow1P7 = new Chunk("Risk Adjusted Performance: ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150)));
        //Chunk PFooterRow1P8 = new Chunk("The annualized return in excess of the return you would expect from your portfolio given its level of market exposure and the market's overall return.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        Chunk PFooterRow1P9 = new Chunk("See notes for this illustration located in the Appendix under Index Definitions for important information.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));


        PFooterRow1.Add(PFooterRow1P1);
        PFooterRow1.Add(PFooterRow1P2);
        PFooterRow2.Add(PFooterRow1P3);
        PFooterRow2.Add(PFooterRow1P4);
        PFooterRow3.Add(PFooterRow1P5);
        PFooterRow3.Add(PFooterRow1P6);
        //PFooterRow4.Add(PFooterRow1P7);
        //PFooterRow4.Add(PFooterRow1P8);
        PFooterRow5.Add(PFooterRow1P9);

        // footPhraseImg.Leading = 8f;
        //        string FooterText = "Gresham Advised Assets (GAA): All Gresham advised investments except cash (includes non-marketable strategies - illiquid real assets and private equity)." +
        //"\nGresham Advised Marketable Assets (Marketable GAA): All Gresham advised investments except cash and non-marketable strategies." +
        //"\nWeighted Benchmark: The average of the benchmark return for each asset class, weighted by that asset class' percentage of total marketable GAA each month." +
        //"\nRisk Adjusted Performance: The annualized return in excess of the return you would expect from your portfolio given its level of market exposure and the market's overall return." +
        //"\nSee notes for this illustration located in the Appendix under Index Definitions for important information.";

        //Phrase footPhraseImg = new Phrase(FooterText, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        ////  HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
        //  footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
        //  footer.Alignment = Element.ALIGN_LEFT; footer.Alignment = Element.ALIGN_TOP;
        //  document.Footer = footer;

        CellFooterRow1.AddElement(PFooterRow1);
        CellFooterRow2.AddElement(PFooterRow2);
        CellFooterRow3.AddElement(PFooterRow3);
        //CellFooterRow4.AddElement(PFooterRow4);
        CellFooterRow5.AddElement(PFooterRow5);

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
        //LoFooter.AddCell(CellFooterRow4);
        // LoFooter.AddCell(CellFooterRow5);


        LoFooter.WidthPercentage = 100f;
        LoFooter.TotalWidth = 100f;
        LoFooter.TotalWidth = 950;
        //LoFooter.WriteSelectedRows(0, 7, 30, 100, writer.DirectContent);  --> Original Values  
        /* To set footer towards the bottom, decrease the value */
        LoFooter.WriteSelectedRows(0, 7, 30, 74, writer.DirectContent);
    }

    //Called From Batch Report,meeting book schedule and single house hold parameter.
    public string MergeReports(string DestinationFileName, string ReportName)
    {

        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss");
        // = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\Gresham_" + strGUID + ".pdf";
        string[] SourceFileName = new string[1];

        if (ReportName == "Portfolio Construction Chart v1") //its old one which renamed as v1.
            SourceFileName[0] = generatePortfolioConstChart();
        else if (ReportName == "Commitment Schedule")
            SourceFileName[0] = generateCommittmentSchReport();
        else if (ReportName == "Investment Objective Chart")
            SourceFileName[0] = generateInvestmentObjectiveChart();
        else if (ReportName == "Allocation Group Pie Chart")
            SourceFileName[0] = generateAllocationGroupPieChart();
        else if (ReportName == "Overall Pie Chart")
            SourceFileName[0] = generateOverAllPieChart();
        else if (ReportName == "Direct Mgr Detail Report")
            SourceFileName[0] = generateDirectMgrDetail();
        else if (ReportName == "Portfolio Construction Chart") //new porthfolio const chart rpt.
            SourceFileName[0] = generatePortfolioConstChartV2();
        else if (ReportName == "Portfolio Construction Chart v2.1") //portfolio construction chart version 2.0
            SourceFileName[0] = generatePortfolioConstChartV2();
        else if (ReportName == "Client Goals") //Perf Analytics - Report 1
            SourceFileName[0] = generatePerfAnalyticsRpt1();
        else if (ReportName == "Absolute Returns") //Perf Analytics - Report 3
            SourceFileName[0] = generatePerfAnalyticsRpt3();
        else if (ReportName == "Capital Protection") //Perf Analytics - Report 4
            SourceFileName[0] = generatePerfAnalyticsRpt4();
        else if (ReportName == "Marketable Performance") //Marketable Performance --MGR
            SourceFileName[0] = getMarketablePerf();
        else if (ReportName == "Short Term Performance") //AssetsPerformanceSummary - (Report 5)
            SourceFileName[0] = getAssetsPerformanceSummary();
        //added 2_1_2019 Non Marketable (DYNAMO)
        else if (ReportingID.ToUpper() == "AFD08C8B-2E25-E911-8106-000D3A1C025B" || ReportingID.ToUpper() == "806E4D33-1D29-E911-8106-000D3A1C025B" || ReportingID.ToUpper() == "90D6C145-1D29-E911-8106-000D3A1C025B" || ReportingID.ToUpper() == "A47E365E-1D29-E911-8106-000D3A1C025B") //Private Equity Performance||Private REal Asset Performance||Outside Private Equity Performance||Outside Private REal Asset Performance
            SourceFileName[0] = NonMarketablePerf();

        if (SourceFileName[0] != "No Record Found")
        {
            PDFMerge PDF = new PDFMerge();
            PDF.MergeFiles(DestinationFileName, SourceFileName);
            if (SourceFileName[0] != "")
            {
                File.Delete(SourceFileName[0]);
            }
        }
        else
        {
            DestinationFileName = "";
        }
        return DestinationFileName;
    }

    public string getFinalSp(ReportType Type)
    {
        // private string getFinalSp(ReportType Type, string ddlHouseholdValue, string ddlHouseholdText, string txtAsofdate, string ddlCashValue, string ddlReportFlgValue, string ddlAllocationGroupValue, string ddlAllocationGroupText, string ddlReport1and2Value, string ddlAllAssetValue)
        string lsSQL = "";
        string houseHold = "";
        string AsOfDates = AsOfDate.Trim() == "" ? "null" : "'" + AsOfDate.Trim() + "'";
        object HouseHold = HouseHoldText == "" ? "null" : "'" + HouseHoldText + "'";
        object AllocationGroup = AllocationGroupText == "" ? "null" : "'" + AllocationGroupText + "'";
        string LegalEntity = LegalEntityId == "" ? "null" : "'" + LegalEntityId + "'";
        string Fund = FundId == "" ? "null" : "'" + FundId + "'";
        string ReportRollUpGrp = ReportRollUpGroupValue == "" ? "null" : "'" + ReportRollUpGroupValue + "'";
        string GreshamAdvFlg = GreshamAdvisedFlag == "" ? "TIA" : GreshamAdvisedFlag;
        object ReportRollupGroupName = ReportRollupGroupIdName == "" ? "null" : "'" + ReportRollupGroupIdName + "'";
        object AssetClass = AssetClassCSV == "" ? "null" : "'" + AssetClassCSV + "'";
        string PriorDates = PriorDate.Trim() == "" ? "null" : "'" + PriorDate.Trim() + "'";

        //added 2_1_2019 Non Marketable (DYNAMO)
        string ReportRollupGroup = ReportRollupGroupId == "" ? "null" : "'" + ReportRollupGroupId + "'";
        string txtHouseholdId = HouseholdId == "" ? "null" : "'" + HouseholdId + "'";
        string HHParameter = HHParameterTxt == "" ? "null" : "'" + HHParameterTxt + "'";
        string txtFundIRR = FundIRR == "" ? "null" : "'" + FundIRR + "'";
        string txtGaOrTIAFlg = _GreshamAdvisedFlag == "" ? "null" : "'" + _GreshamAdvisedFlag + "'";

        if (Type == ReportType.PortfolioConsChart)
        {
            lsSQL = "exec SP_R_CONSTRUCTIONCHART_EXCEL_NEW_GA_BASEDATA " + HouseHold + "," + AsOfDates + ", null, " + AllocationGroup + "";//"exec SP_R_COMMITMENTSCHEDULE 'Kuppenheimer Family', '20110430', null, 0, 0, null, 0, 1";
        }
        else if (Type == ReportType.CommitmentSchedule)
        {
            lsSQL = "exec SP_R_COMMITMENTSCHEDULE_NEW_GA " + HouseHold + "," + AsOfDates + ", null, 0,0, " + AllocationGroup + ",0,1";
        }
        else if (Type == ReportType.InvestmentObjectiveChart)
        {
            lsSQL = "exec GreshamPartners_MSCRM.dbo.SP_R_INVESTMENT_OBJECTIVE_CHART_EXCEL_SMA_NEW_BASEDATA  @HouseholdName  = " + HouseHold + ", @AsofDate = " + AsOfDates + ", @GreshamAdvisedFlagTxt = 'TIA',@AllocGroupName = " + AllocationGroup + "";
        }
        else if (Type == ReportType.AllocationGroupPieChart)
        {
            lsSQL = "exec SP_R_CONSTRUCTIONCHART_EXCEL_NEW_GA_BASEDATA " + HouseHold + "," + AsOfDates + ", null, null";
        }
        else if (Type == ReportType.OverallPieChart)
        {
            lsSQL = "exec SP_R_CONSTRUCTIONCHART_EXCEL_NEW_GA_BASEDATA " + HouseHold + "," + AsOfDates + ", null, " + AllocationGroup + "";
        }
        else if (Type == ReportType.DirectMgrDetail)
        {
            lsSQL = "SP_R_DIRECT_MANAGER @AsofDate =" + AsOfDates + ",@GreshamFundID =" + Fund + ",@LegalEntityID =" + LegalEntity;
        }
        else if (Type == ReportType.PortfolioConsChartNEW)
        {
            if (PortFolioConChartRptVer == "old")
                lsSQL = "SP_R_CONSTRUCTIONCHART_EXCEL @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + ",@GreshamAdvisedFlagTxt = 'TIA'";//old report
            else
                lsSQL = "SP_R_CONSTRUCTIONCHART_EXCEL_NEW_GA_BASEDATA @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + ",@GreshamAdvisedFlagTxt = '" + GreshamAdvFlg + "',@Reportrollupgroupid=" + ReportRollUpGrp + "";
        }
        else if (Type == ReportType.PortfolioConsChartV2)
        {
            lsSQL = "EXEC SP_R_CONSTRUCTIONCHART_EXCEL_NEW_GA_NEW @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + ",@GreshamAdvisedFlagTxt = '" + GreshamAdvFlg + "',@Reportrollupgroupid=" + ReportRollUpGrp + "";
        }
        else if (Type == ReportType.PerfAnalytics)
        {
            //   lsSQL = "SP_S_ReportRollupGroupAllocationName @RRGName =" + Grp + "";
        }
        else if (Type == ReportType.Rpt1LineChart)
        {
            lsSQL = "SP_R_CLIENT_GOALS_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + ", @TrxnGAFlagTxt = '" + GreshamAdvFlg + "',@AsOfDate = " + AsOfDates + "";
        }
        else if (Type == ReportType.Rpt1Table)
        {
            lsSQL = "SP_R_ASSET_VALUE_NEW_GA_BASEDATA @GroupName =" + ReportRollupGroupName + ",@PositionGAFlagTxt='" + GreshamAdvFlg + "' ,@TrxnGAFlagTxt= '" + GreshamAdvFlg + "' ,@AsOfDate = " + AsOfDates + "";
        }
        else if (Type == ReportType.Rpt3LineChart)
        {
            lsSQL = "exec SP_R_WEALTH_CHART_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + " , @PositionGAFlagTxt = '" + GreshamAdvFlg + "', @TrxnGAFlagTxt = '" + GreshamAdvFlg + "' ,@AsOfDate = " + AsOfDates + ",@AssetNameTxt = " + AssetClass + ",@InclFixedIncome = 1";
        }
        else if (Type == ReportType.Rpt3BarChart)
        {
            lsSQL = "exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + ", @PositionGAFlagTxt = '" + GreshamAdvFlg + "' , @TrxnGAFlagTxt = '" + GreshamAdvFlg + "' ,@AsOfDate = " + AsOfDates + ", @AnnPerfFlg = 0 , @HouseHoldName =" + HouseHold + ",@AssetNameTxt = " + AssetClass + ",@InclFixedIncome = 1";
        }
        else if (Type == ReportType.Rpt3Table1)
        {
            lsSQL = "exec SP_R_ANNUAL_PERFORMANCE_NEW_GA_BASEDATA  @GroupName = " + ReportRollupGroupName + ",@PositionGAFlagTxt = '" + GreshamAdvFlg + "',@TrxnGAFlagTxt = '" + GreshamAdvFlg + "',@AsOfDate =" + AsOfDates + " , @AnnPerfFlg = 1 , @HouseHoldName =" + HouseHold + ",@AssetNameTxt = " + AssetClass + ",@InclFixedIncome = 1";
        }
        else if (Type == ReportType.Rpt3Table2)
        {
            lsSQL = "exec SP_R_CPI_PERFORMANCE  @AsOfDate =" + AsOfDates + "  ,@SinceInceptDT ='" + IncDate.ToString() + "',@BenchMarkID	= '75D35570-F8BB-E211-8A81-0019B9E7EE05'";
        }
        else if (Type == ReportType.Rpt4ColumnChart)
        {
            lsSQL = "Exec SP_R_WORST_MONTH_MAXDD_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + ",@PositionGAFlagTxt = '" + GreshamAdvFlg + "',@TrxnGAFlagTxt = '" + GreshamAdvFlg + "',@AsOfDate = " + AsOfDates + ",@BenchMarkName = null,@AssetNameTxt = " + AssetClass + "";
        }
        else if (Type == ReportType.Rpt4ShapeChart)
        {
            lsSQL = "EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + ",@PositionGAFlagTxt = '" + GreshamAdvFlg + "',@TrxnGAFlagTxt = '" + GreshamAdvFlg + "',@AsOfDate = " + AsOfDates + ",@BenchMarkName = null,@AssetNameTxt = " + AssetClass + ",@StartDate = '01-JAN-2011'";
        }
        else if (Type == ReportType.Rpt4TableLT)
        {
            lsSQL = "EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + ",@PositionGAFlagTxt = '" + GreshamAdvFlg + "',@TrxnGAFlagTxt = '" + GreshamAdvFlg + "',@AsOfDate = " + AsOfDates + ",@BenchMarkName = null,@AssetNameTxt = " + AssetClass + "";
        }
        else if (Type == ReportType.Rpt4TableST)
        {
            lsSQL = "EXEC SP_R_RETURN_STD_DEV_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + ",@PositionGAFlagTxt = '" + GreshamAdvFlg + "',@TrxnGAFlagTxt = '" + GreshamAdvFlg + "',@AsOfDate = " + AsOfDates + ",@BenchMarkName = null,@AssetNameTxt = " + AssetClass + ",@StartDate = '01-JAN-2011'";
        }
        else if (Type == ReportType.RptMarketablePerf)
        {
            lsSQL = "EXEC SP_R_MGR_REPORT_NEW_PDF @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + "";
            //lsSQL = "EXEC SP_R_MGR_REPORT_NEW_PDF_TEST @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + "";
        }
        else if (Type == ReportType.RptAssetsPerformanceSummary)
        {
            lsSQL = "EXEC SP_R_PERF_CALCS_NEW_UNIT_STUDY_NEW_GA_BASEDATA @GroupName = " + ReportRollupGroupName + ",@AsofDate =" + AsOfDates + ",@AssetNameTxt = " + AssetClass + ",@PdfFlg=1,@StarDT = " + PriorDates + "";
        }
        else if (Type == ReportType.RptNonMarketablePerf)  //added 2_1_2019 Non Marketable (DYNAMO)
        {
            lsSQL = "EXEC SP_R_DYNAMO_NONMARKETABLE @HouseholdUUID =" + txtHouseholdId + ",@AsofDate =" + AsOfDates + ",@FundIRRFlg =" + txtFundIRR + ",@GaOrTIAFlg =" + txtGaOrTIAFlg + ",@Reportrollupgroupid =" + ReportRollupGroup + ",@AllocGroupName =" + AllocationGroup + ",@HHParameterTxt =" + HHParameter + ",@LegalEntityUUID =" + LegalEntity + ""; //added 4/16/2019 @LegalEntityUUID  
            //lsSQL = "EXEC SP_R_MGR_REPORT_NEW_PDF_TEST @HouseholdName =" + HouseHold + ",@AsofDate =" + AsOfDates + ",@AllocGroupName =" + AllocationGroup + "";
        }


        //HttpContext.Current.Response.Write(lsSQL);
        return lsSQL;

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

    private DataSet AddTotalsInvestmentObjectiveChart(DataSet lodataset)
    {
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            DataRow dr = lodataset.Tables[0].NewRow();

            for (int j = 0; j < lodataset.Tables[0].Rows.Count; j++)
            {
                try
                {
                    for (int k = 0; k < lodataset.Tables[0].Columns.Count; k++)
                    {
                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Current Portfolio %"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;
                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }
                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Current Portfolio Value"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Suggested Allocation"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }
                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Tactical Target"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }
                    }
                }
                catch
                {
                }
            }
            dr["_LineFlg"] = 2;
            dr["Asset Class"] = "TOTAL";
            lodataset.Tables[0].Rows.Add(dr);
            lodataset.AcceptChanges();
        }
        return lodataset;
    }

    private DataSet AddTotalsDirectMgr(DataSet lodataset)
    {
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            DataRow dr = lodataset.Tables[0].NewRow();

            for (int j = 0; j < lodataset.Tables[0].Rows.Count; j++)
            {
                try
                {
                    for (int k = 0; k < lodataset.Tables[0].Columns.Count; k++)
                    {

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("MarketValue"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("PctPortfolio"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }


                        //if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Suggested Allocation"))
                        //{
                        //    if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                        //        dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                        //    dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        //}

                        //if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Tactical Tilt"))
                        //{
                        //    if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                        //        dr[lodataset.Tables[0].Columns[k].ColumnName] = 0;

                        //    dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        //}

                        //if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Tactical Target"))
                        //{
                        //    if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                        //        dr[lodataset.Tables[0].Columns[k].ColumnName] = 0;

                        //    dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        //}


                    }
                }
                catch
                {

                }

            }



            //dr["_LineFlg"] = 2;
            dr["Security"] = "Total Porfolio";
            lodataset.Tables[0].Rows.Add(dr);
            lodataset.AcceptChanges();


        }
        return lodataset;
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
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(6);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk loParagraph = new Chunk();

        if (AsOfDate != "")
            lsDateName = Convert.ToDateTime(AsOfDate).ToString("MMMM dd, yyyy") + "";

        Chunk lochunk = new Chunk(lsFamiliesName, setFontsAll(12, 1, 0));
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + "ASSET ALLOCATION SUMMARY", setFontsAll(11, 0, 0));

        loCell.Add(lochunk);
        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        loCell.HorizontalAlignment = 1;

        lochunk = new Chunk("\n" + lsDateName, setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        loCell.Add(lochunk);
        loCell.Border = 0;
        loTable.AddCell(loCell);

        Boolean lbCheckFoMarket = false;

        lochunk = new Chunk("", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell0 = new Cell();
        loCell0.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell0.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell0.Border = 0;
        loTable.AddCell(loCell0);

        lochunk = new Chunk("Current  Allocation", setFontsAll1(7, 0));
        iTextSharp.text.Chunk lochunk5 = new Chunk("\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "\t " + "$                         %", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk1 = new Chunk(lochunk + "\n ", setFontsAll(7, 1, 0));

        iTextSharp.text.Cell loCell1 = new Cell();

        loCell1.Add(lochunk1);//.SetUnderline(0.8f,-1f)

        loCell1.Add(lochunk5);

        loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell1.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell1.Colspan = 2;
        loCell1.Border = 0;
        loCell1.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        //loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell1.MaxLines = 3;
        loCell1.Leading = 10f;
        //loCell1.EnableBorderSide(2);
        loTable.AddCell(loCell1);

        iTextSharp.text.Chunk lochunk2 = new Chunk("Strategic " + "\n ", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk6 = new Chunk("Allocation", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell2 = new Cell();
        loCell2.Add(lochunk2);
        loCell2.Add(lochunk6);
        loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell2.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell2.Border = 0;
        loCell2.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell2.MaxLines = 3;
        loCell2.Leading = 10f;
        loTable.AddCell(loCell2);

        iTextSharp.text.Chunk lochunk3 = new Chunk("Tactical " + "\n ", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk7 = new Chunk("Tilt", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell3 = new Cell();
        loCell3.Add(lochunk3);
        loCell3.Add(lochunk7);
        loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell3.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell3.Border = 0;
        loCell3.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell3.MaxLines = 3;
        loCell3.Leading = 10f;
        loTable.AddCell(loCell3);

        iTextSharp.text.Chunk lochunk4 = new Chunk("Tactical " + "\n ", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk8 = new Chunk("Target", setFontsAll(7, 1, 0));

        iTextSharp.text.Cell loCell4 = new Cell();
        loCell4.Add(lochunk4);
        loCell4.Add(lochunk8);

        loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell4.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell4.Border = 0;
        loCell4.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell4.MaxLines = 3;
        loCell4.Leading = 10f;
        loTable.AddCell(loCell4);

        foDocument.Add(loTable);
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
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
            string lblError = "Record not found";
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
        //if (ddlHousehold.SelectedValue != "0")
        //{
        //    if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
        //    {
        //        lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
        //    }
        //    else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
        //    {
        //        lsfamilyName = drpHouseHoldReportTitle.SelectedItem.Text;
        //    }
        //    else
        //    {
        //        lsfamilyName = drpAllocationGroupTitle.SelectedItem.Text;
        //    }
        //}




        if (AsOfDate != "")
            lsDateName = Convert.ToDateTime(AsOfDate).ToString("MMMM dd, yyyy") + "";

        /////////////

        //Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));

        Chunk lochunk = new Chunk(CommitmentReportHeader, setFontsAll(12, 1, 0));
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
                    if (Convert.ToString(OldDataset.Tables[0].Rows[1]["_CurrentQuarter"]) != "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsCurrentQuarter", Convert.ToString(OldDataset.Tables[0].Rows[1]["_CurrentQuarter"])), setFontsAll(7, 1, 0));//Convert.ToString(OldDataset.Tables[0].Rows[0]["_CurrentQuarter"])
                    }
                    else if (Convert.ToString(OldDataset.Tables[0].Rows[1]["_CurrentQuarter"]) == "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsCurrentQuarter", "Expected Calls in Current Quarter"), setFontsAll(7, 1, 0));
                    }

                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("ExpectedRemainingCallsNextQuarter"))
                {
                    if (Convert.ToString(OldDataset.Tables[0].Rows[1]["_NextQuarter"]) != "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsNextQuarter", Convert.ToString(OldDataset.Tables[0].Rows[1]["_NextQuarter"])), setFontsAll(7, 1, 0));//Convert.ToString(OldDataset.Tables[0].Rows[0]["_NextQuarter"])
                    }
                    else if (Convert.ToString(OldDataset.Tables[0].Rows[1]["_NextQuarter"]) == "")
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("ExpectedRemainingCallsNextQuarter", "Expected Calls in Next Quarter"), setFontsAll(7, 1, 0));
                    }
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("RemainingCommitment"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("RemainingCommitment", "Remaining Commitment"), setFontsAll(7, 1, 0));
                }
                else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("AverageQuarterlyDist"))
                {
                    lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("AverageQuarterlyDist", "Last Four Qtrs Average Distribution"), setFontsAll(7, 1, 0));
                }
                //else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Investment"))
                //{
                //   lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Investment", "Investment*"), setFontsAll(7, 1, 0));
                //}
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

            loCell.MaxLines = 3;
            loCell.Leading = -2F;
            if (liColumnCount != 0)
            {
                if (liColumnCount == AddDataset.Tables[0].Columns.Count - 1)
                {
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                }
                else
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            }
            else if (liColumnCount == 0)
            {
                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
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
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        foDocument.Add(png);
    }

    public void setHeaderDirectMgrDetails(Document foDocument, DataSet loInsertdataset, DataSet loDatatset)
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
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(3);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk loParagraph = new Chunk();


        //////// set header new addition for pdf
        string lsfamilyName = "";
        //if (ddlHousehold.SelectedValue != "0")
        //{
        //    if (drpAllocationGroupTitle.SelectedValue == "0" && ddlAllocationGroup.SelectedValue != "0")
        //    {
        //        lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
        //    }
        //    else if (ddlHousehold.SelectedValue != "0" && ddlAllocationGroup.SelectedValue == "0")
        //    {
        //        lsfamilyName = drpHouseHoldReportTitle.SelectedItem.Text;
        //    }
        //    else
        //    {
        //        lsfamilyName = drpAllocationGroupTitle.SelectedItem.Text;
        //    }
        //}




        if (AsOfDate != "")
            lsDateName = Convert.ToDateTime(AsOfDate).ToString("MMMM dd, yyyy") + "";

        /////////////

        //Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));

        Chunk lochunk = new Chunk(Convert.ToString(OldDataset.Tables[1].Rows[0]["LegalEntity"]), setFontsAll(12, 1, 0));
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + Convert.ToString(OldDataset.Tables[1].Rows[0]["Fund"]), setFontsAll(12, 1, 0));
        //iTextSharp.text.Cell loCell = new Cell();
        //loCell.Add(lochunk);

        Chunk lochunk12 = new Chunk("\n" + "HOLDINGS", setFontsAll(11, 0, 0));
        //loParagraph.Chunks.Add(lochunk);

        loCell.Add(lochunk);
        loCell.Add(lochunk12);
        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        loCell.HorizontalAlignment = 1;



        lochunk = new Chunk("\n" + lsDateName, setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
        loTable.AddCell(loCell);

        Boolean lbCheckFoMarket = false;
        #region No Use
        //for (int liColumnCount = 0; liColumnCount < AddDataset.Tables[0].Columns.Count; liColumnCount++)
        //{

        //    if (Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName) != "")
        //    {
        //        if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Current Portfolio Value"))
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Current Portfolio Value", "Current \n $"), setFontsAll1(7, 1));
        //        }
        //        else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Current Portfolio %"))
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Current Portfolio %", ""), setFontsAll1(7, 1));
        //        }
        //        else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Suggested Allocation"))
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Suggested Allocation", "Strategic Allocation"), setFontsAll(7, 1, 0));
        //        }
        //        else
        //        {
        //            lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName), setFontsAll(7, 1, 0));
        //        }
        //    }
        //    //}
        //    loCell = new Cell();

        //    loCell.Add(lochunk);
        //    loCell.Border = 0;
        //    loCell.NoWrap = true;//true;

        //    loCell.MaxLines = 2;
        //    loCell.Leading = -2F;
        //    if (liColumnCount == 0 )
        //    {
        //        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        //    }

        //    else if (liColumnCount == 1)
        //    {
        //        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        //        loCell.BorderWidthBottom = 1f;
        //    }
        //    else
        //    {
        //        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        //    }


        //    if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
        //    {
        //        loCell.Leading = 10f;//8
        //        loCell.MaxLines = 5;
        //        //loCell.Leading = 9f;
        //    }
        //    loCell.Leading = 10f;//8

        //    loCell.VerticalAlignment = 1; //5 ,6 bottom : WASTE VALUES - 3,4
        //    loTable.AddCell(loCell);

        //}
        #endregion

        //lochunk = new Chunk("", setFontsAll(7, 1,0));
        //iTextSharp.text.Cell loCell0 = new Cell();
        //loCell0.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //loCell0.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        //loCell0.Border = 0;
        //loTable.AddCell(loCell0);

        //lochunk = new Chunk("Legal Entity", setFontsAll(7, 1, 0));
        //iTextSharp.text.Chunk lochunk5 = new Chunk("", setFontsAll(7, 1, 0));
        //iTextSharp.text.Chunk lochunk1 = new Chunk(lochunk + "\n ", setFontsAll(7, 1, 0));


        //iTextSharp.text.Cell loCell1 = new Cell();

        ////loCell1.Add(lochunk);//.SetUnderline(0.8f,-1f)

        ////loCell1.Add(lochunk5);

        //loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //loCell1.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        ////loCell1.Colspan = 2;
        //loCell1.Border = 0;
        //loCell1.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        ////loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        //loCell1.MaxLines = 3;
        //loCell1.Leading = 10f;
        ////loCell1.EnableBorderSide(2);
        //loTable.AddCell(loCell1);

        iTextSharp.text.Chunk lochunk2 = new Chunk("Security", setFontsAll(7, 1, 0));
        iTextSharp.text.Chunk lochunk6 = new Chunk("", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell2 = new Cell();
        loCell2.Add(lochunk2);
        //loCell2.Add(lochunk6);
        loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
        loCell2.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell2.Border = 0;
        loCell2.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell2.MaxLines = 3;
        loCell2.Leading = 10f;
        loTable.AddCell(loCell2);

        iTextSharp.text.Chunk lochunk3 = new Chunk("Market Value", setFontsAll(7, 1, 0));
        //iTextSharp.text.Chunk lochunk7 = new Chunk("", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell3 = new Cell();
        loCell3.Add(lochunk3);
        //loCell3.Add(lochunk7);
        loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell3.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell3.Border = 0;
        //loCell3.EnableBorderSide(2);
        loCell3.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell3.MaxLines = 3;
        loCell3.Leading = 10f;
        loTable.AddCell(loCell3);

        iTextSharp.text.Chunk lochunk4 = new Chunk("% Portfolio", setFontsAll(7, 1, 0));
        //iTextSharp.text.Chunk lochunk8 = new Chunk("", setFontsAll(7, 1, 0));

        iTextSharp.text.Cell loCell4 = new Cell();
        loCell4.Add(lochunk4);
        //loCell4.Add(lochunk8);

        loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell4.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell4.Border = 0;
        //loCell4.EnableBorderSide(2);
        loCell4.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell4.MaxLines = 3;
        loCell4.Leading = 10f;

        loTable.AddCell(loCell4);


        foDocument.Add(loTable);

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        foDocument.Add(png);
    }

    public iTextSharp.text.Table addpageno(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt)
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

    public void checkTrue(DataSet foDataset, int fiRowCount, String fsField, Cell foCell, iTextSharp.text.Color foColor)
    {

        if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString() == "6")
        {
            foCell.BackgroundColor = foColor;

        }
    }
    public string addPageIndex1(string Rptname, DataTable dtbatch, String TempFolderPath)
    {
        iTextSharp.text.Document document = null;

        try
        {
            int numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
            int numIndexPageSize = 20; // Size of index page 
            int StartRow = 0;
            int EndRow = 0;
            //No of pages in the Coversheet that isappended 
            int coverLetterPages = Convert.ToInt32(HttpContext.Current.Session["pageinCoverLetter"].ToString());
            double total = (double)dtbatch.Rows.Count / numIndexPageSize;
            int liTotalPage = Convert.ToInt32(Math.Ceiling(total));

            numIndexPageCount = numIndexPageCount + liTotalPage;
            //  numIndexPageCount = numIndexPageCount + 1;
            numIndexPageCount = numIndexPageCount + coverLetterPages;// add total pages plus the coverletter pages

            //  clsCombinedReports objCombinedReports = new clsCombinedReports();
            PdfReader reader = new PdfReader(Rptname);
            //  string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\test_" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
            string filename = TempFolderPath + "\\" + "index_" + Guid.NewGuid().ToString() + ".pdf";

            FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
            MemoryStream stream = new MemoryStream();
            //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            var writer = PdfWriter.GetInstance(document, fileStream);
            PdfStamper stamper = new PdfStamper(reader, stream);

            iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.GRAY);
            iTextSharp.text.Font whiteFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.WHITE);

            //  Dictionary<string, int> dicNumFilesCount = (Dictionary<string, int>)Session["BatchDic"];

            //foreach (KeyValuePair<string, int> pair in dicNumFilesCount)
            //{
            //    Response.Write(pair.Key.ToString() + " : " + pair.Value.ToString() + "<br/>");
            //}

            document.Open();

            for (var i = 1; i <= reader.NumberOfPages; i++)
            {
                HttpContext.Current.Session["NumberofPages"] = reader.NumberOfPages;
                if (i <= coverLetterPages)//all the pages inside the coversheet to be set to potrait
                {
                    document.SetPageSize(PageSize.A4);
                }
                else if (i > coverLetterPages)
                {
                    //else

                    document.SetPageSize(PageSize.A4.Rotate());
                }
                document.NewPage();
                string fontpath = HttpContext.Current.Server.MapPath(".");
                var baseFont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
                var importedPage = writer.GetImportedPage(reader, i);

                var contentByte = writer.DirectContent;
                // contentByte.BeginText();
                contentByte.SetFontAndSize(baseFont, 2);

                if (i > coverLetterPages + 1 && i <= numIndexPageCount) //Index Page //coverletter +1 is for all pages 
                {
                    PdfPTable tableindex = new PdfPTable(3);

                    int[] widthtable = { 80, 10, 10 };
                    tableindex.SetWidths(widthtable);
                    // tableindex.TotalWidth = 20f;
                    tableindex.WidthPercentage = 75f;

                    if (i == coverLetterPages + 2)//CoverLettter +2 as to write the index page
                    {
                        // Chapter chapter1 = new Chapter(new Paragraph("This is Chapter 1"), 1);
                        // Section section1 = chapter1.AddSection(20f, "Section 1.1", 2);
                        // chapter1.BookmarkTitle = "Changed Title";
                        //  chapter1.BookmarkOpen = true;

                        //   document.Add(chapter1);

                        StartRow = 0;
                        EndRow = numIndexPageSize;

                        if (dtbatch.Rows.Count < EndRow)
                            EndRow = dtbatch.Rows.Count;
                    }
                    else
                    {
                        StartRow = EndRow;
                        EndRow = numIndexPageSize * (i - 1);
                        if (dtbatch.Rows.Count < EndRow)
                            EndRow = dtbatch.Rows.Count;

                    }
                    HttpContext.Current.Session["Rowcount"] = dtbatch.Rows.Count.ToString();
                    HttpContext.Current.Session["EndRow"] = EndRow;
                    int numRows = 0;
                    for (int x = StartRow; x < EndRow; x++)
                    {
                        if (numRows == 0) //Heading
                        {
                            PdfPCell loRptNameHeading = new PdfPCell();
                            PdfPCell loPageNoHeading = new PdfPCell();
                            PdfPCell loBlank = new PdfPCell();

                            Paragraph lochunk1Heading = new Paragraph(" Report ", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));

                            Paragraph lochunk2Heading = new Paragraph("Page", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                            lochunk2Heading.SetAlignment("center");

                            Paragraph lochunkBlank = new Paragraph(" ", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                            lochunkBlank.SetAlignment("center");

                            Anchor targetindex = new Anchor("Index", setFontsAll(4, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                            //targetindex.Reference = "#Index";

                            string key = "Index";
                            targetindex.Name = key.ToString();


                            PdfOutline Bookmarkindex = writer.RootOutline;
                            PdfOutline mbot1 = new PdfOutline(Bookmarkindex, PdfAction.GotoLocalPage("Index", false), "TABLE OF CONTENTS");

                            lochunkBlank.Add(targetindex);

                            //lochunk1Heading.Leading = 11f;
                            //    lochunk1Heading.SetLeading(15f, 5f);
                            //    lochunk2Heading.SetLeading(15f, 5f);


                            loRptNameHeading.AddElement(lochunk1Heading);
                            loPageNoHeading.AddElement(lochunk2Heading);
                            loBlank.AddElement(lochunkBlank);

                            loRptNameHeading.PaddingBottom = 5f;
                            loPageNoHeading.PaddingBottom = 5f;
                            loBlank.PaddingBottom = 5f;

                            loRptNameHeading.Border = 0;
                            loPageNoHeading.Border = 0;
                            loBlank.Border = 0;

                            loRptNameHeading.BorderWidthBottom = 0.7f;
                            loPageNoHeading.BorderWidthBottom = 0.7f;

                            loRptNameHeading.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"));
                            loPageNoHeading.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"));

                            tableindex.AddCell(loRptNameHeading);
                            tableindex.AddCell(loPageNoHeading);
                            tableindex.AddCell(loBlank);
                        }

                        //Table Content 
                        PdfPCell loRptName = new PdfPCell();    //FirstColumn : Report Name 
                        PdfPCell loPageNo = new PdfPCell();     //SecondColumn : Report Name 
                        PdfPCell loBlankCol = new PdfPCell();

                        string strRptname = getReportName(dtbatch, x);



                        Anchor target = new Anchor(strRptname, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                        target.Reference = "#" + dtbatch.Rows[x]["ssi_greshamreportidname"].ToString().Trim().Replace(" ", "_") + dtbatch.Rows[x]["numPageNo"].ToString();

                        PdfOutline root = writer.RootOutline;
                        PdfOutline mbot = new PdfOutline(root, PdfAction.GotoLocalPage(dtbatch.Rows[x]["ssi_greshamreportidname"].ToString().Trim().Replace(" ", "_") + dtbatch.Rows[x]["numPageNo"].ToString(), false), strRptname);


                        string Title = " ";

                        Paragraph lochunk1 = new Paragraph(" " + Title, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                        lochunk1.Add(target);
                        //lochunk1.SetAlignment("middle");


                        int pagenum = Convert.ToInt32(dtbatch.Rows[x]["numPageNo"]);

                        Paragraph lochunk2 = new Paragraph(Convert.ToString(pagenum + numIndexPageCount), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                        lochunk2.SetAlignment("center");


                        Paragraph lochunkBlank1 = new Paragraph(" ", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                        lochunkBlank1.SetAlignment("center");

                        //lochunk2.SetAlignment("middle");

                        loRptName.AddElement(lochunk1);
                        loPageNo.AddElement(lochunk2);
                        loBlankCol.AddElement(lochunkBlank1);

                        loRptName.Border = 0;
                        loPageNo.Border = 0;
                        loBlankCol.Border = 0;
                        loRptName.PaddingBottom = 5f;
                        loPageNo.PaddingBottom = 5f;



                        loRptName.BorderWidthBottom = 0.7f;
                        loPageNo.BorderWidthBottom = 0.7f;

                        loRptName.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
                        loPageNo.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));

                        loRptName.VerticalAlignment = Element.ALIGN_MIDDLE;
                        loPageNo.VerticalAlignment = Element.ALIGN_MIDDLE;

                        tableindex.AddCell(loRptName);
                        tableindex.AddCell(loPageNo);
                        tableindex.AddCell(loBlankCol);

                        //tableindex.HorizontalAlignment = 0;


                        numRows++;
                    }

                    PdfPTable tableTitle = new PdfPTable(1);
                    int[] width = { 100 };
                    tableTitle.SetWidths(width);

                    //tableTitle.TotalWidth = 100f;
                    tableTitle.WidthPercentage = 60f;
                    Paragraph pTitle = new Paragraph("TABLE OF CONTENTS", setFontsAll(14, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                    pTitle.SetAlignment("center");
                    pTitle.IndentationLeft = 265f;
                    PdfPCell loRTitle = new PdfPCell();
                    loRTitle.Border = 0;
                    loRTitle.PaddingBottom = 5f;
                    loRTitle.PaddingTop = 15f;
                    loRTitle.AddElement(pTitle);
                    tableTitle.AddCell(loRTitle);
                    tableTitle.HorizontalAlignment = 0;
                    string blanktitle = "";
                    Anchor targetind = new Anchor("A ", whiteFont);
                    targetind.Reference = "#Index";
                    //  string key = dicNumFilesCount.FirstOrDefault(x => x.Value == i - 2).Key;

                    //  targetind.Name = "Index";



                    Paragraph lochunkindex = new Paragraph("\n" + blanktitle);


                    lochunkindex.Add(blanktitle);

                    document.Add(lochunkindex); //blank chunk to target table of content 

                    document.Add(tableTitle);
                    document.Add(tableindex);

                    //string Title = " ";
                    //Anchor target = new Anchor("Portfolio Construction Chart v2.1");
                    //target.Reference = "#pc2";

                    //Paragraph lochunk1 = new Paragraph("\n" + Title);
                    //lochunk1.Add(target);

                    //Anchor target1 = new Anchor("Asset Distribution");
                    //target1.Reference = "#AD";

                    //Paragraph lochunk2 = new Paragraph("\n" + Title);
                    //lochunk2.Add(target1);

                    //document.Add(lochunk1);
                    //document.Add(lochunk2);
                }
                else if (i > numIndexPageCount)
                {
                    int numpage = i - numIndexPageCount;
                    DataRow[] rows = dtbatch.Select("numPageNo = " + numpage.ToString() + "");
                    if (rows.Length > 0)
                    {
                        string Title = "";
                        Anchor target = new Anchor("A ", whiteFont);

                        //  string key = dicNumFilesCount.FirstOrDefault(x => x.Value == i - 2).Key;
                        string key = rows[0]["ssi_greshamreportidname"].ToString() + rows[0]["numPageNo"].ToString();
                        target.Name = key.ToString().Trim().Replace(" ", "_");



                        Paragraph lochunk1 = new Paragraph("\n" + Title);


                        lochunk1.Add(target);

                        document.Add(lochunk1);
                    }



                    PdfContentByte cb = writer.DirectContent;
                    ColumnText ct = new ColumnText(cb);
                    ct.SetSimpleColumn(new Phrase(new Chunk("Page " + i + " of " + reader.NumberOfPages, setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))))), 480, 15, 400, 40, 25, Element.ALIGN_CENTER | Element.ALIGN_BOTTOM);
                    ct.Go();
                }
                String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

                PdfContentByte cb1 = writer.DirectContent;
                ColumnText ct1 = new ColumnText(cb1);
                ct1.SetSimpleColumn(new Phrase(new Chunk(lsDateTime, setFontsAll(8, 0, 1, new iTextSharp.text.Color(216, 216, 216)))), 800, 15, 725, 40, 25, Element.ALIGN_RIGHT | Element.ALIGN_BOTTOM);
                ct1.Go();

                // var multiLineString = "" + i.ToString() + "!".Split('\n');

                // foreach (var line in multiLineString)
                //  {
                // contentByte.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Page " + i + " of " + reader.NumberOfPages, 825, 10, 0);
                //  contentByte.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, " ", 825, 10, 0);
                // }

                //contentByte.EndText();
                iTextSharp.text.Rectangle psize = reader.GetPageSizeWithRotation(i);
                HttpContext.Current.Session["VALUE"] = i.ToString() + " Rotation is" + psize.ToString();
                switch (psize.Rotation)
                {
                    case 0:
                        contentByte.AddTemplate(importedPage, 1f, 0, 0, 1f, 0, 0);
                        break;
                    case 90:
                        contentByte.AddTemplate(importedPage, 0, -1f, 1f, 0, 0, psize.Height);
                        break;
                    case 180:
                        contentByte.AddTemplate(importedPage, -1f, 0, 0, -1f, 0, 0);
                        break;
                    case 270:
                        contentByte.AddTemplate(importedPage, 0, 1.0F, -1.0F, 0, psize.Width, 0);
                        break;
                    default:
                        break;
                }
                // contentByte.AddTemplate(importedPage, 1f, 0, 0, 1f, 0, 0);

            }

            document.Close();
            writer.Close();
            return filename;
        }
        catch
        {
            document.Close();
            return "";
        }
    }

    public string addPageIndex(string Rptname, DataTable dtbatch, String TempFolderPath)
    {
        iTextSharp.text.Document document = null;

        try
        {
            int numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
            int numIndexPageSize = 20; // Size of index page 
            int StartRow = 0;
            int EndRow = 0;

            double total = (double)dtbatch.Rows.Count / numIndexPageSize;
            int liTotalPage = Convert.ToInt32(Math.Ceiling(total));

            numIndexPageCount = numIndexPageCount + liTotalPage;

            //  clsCombinedReports objCombinedReports = new clsCombinedReports();
            PdfReader reader = new PdfReader(Rptname);
            //string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\test_" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
            string filename = TempFolderPath + "\\" + "index_" + Guid.NewGuid().ToString() + ".pdf";

            FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
            MemoryStream stream = new MemoryStream();
            //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            //var writer = PdfWriter.GetInstance(document, fileStream);
            var writer = PdfWriter.GetInstance(document, fileStream);
            PdfStamper stamper = new PdfStamper(reader, stream);

            iTextSharp.text.Font blackFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.GRAY);
            iTextSharp.text.Font whiteFont = FontFactory.GetFont("Arial", 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.WHITE);

            //  Dictionary<string, int> dicNumFilesCount = (Dictionary<string, int>)Session["BatchDic"];

            //foreach (KeyValuePair<string, int> pair in dicNumFilesCount)
            //{
            //    Response.Write(pair.Key.ToString() + " : " + pair.Value.ToString() + "<br/>");
            //}

            document.Open();

            for (var i = 1; i <= reader.NumberOfPages; i++)
            {
                document.NewPage();
                string fontpath = HttpContext.Current.Server.MapPath(".");
                var baseFont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
                var importedPage = writer.GetImportedPage(reader, i);

                var contentByte = writer.DirectContent;
                // contentByte.BeginText();
                contentByte.SetFontAndSize(baseFont, 2);

                if (i > 1 && i <= numIndexPageCount) //Index Page 
                {
                    PdfPTable tableindex = new PdfPTable(3);

                    int[] widthtable = { 80, 10, 10 };
                    tableindex.SetWidths(widthtable);
                    // tableindex.TotalWidth = 20f;
                    tableindex.WidthPercentage = 75f;

                    if (i == 2)
                    {
                        // Chapter chapter1 = new Chapter(new Paragraph("This is Chapter 1"), 1);
                        // Section section1 = chapter1.AddSection(20f, "Section 1.1", 2);
                        // chapter1.BookmarkTitle = "Changed Title";
                        //  chapter1.BookmarkOpen = true;

                        //   document.Add(chapter1);

                        StartRow = 0;
                        EndRow = numIndexPageSize;

                        if (dtbatch.Rows.Count < EndRow)
                            EndRow = dtbatch.Rows.Count;
                    }
                    else
                    {
                        StartRow = EndRow;
                        EndRow = numIndexPageSize * (i - 1);
                        if (dtbatch.Rows.Count < EndRow)
                            EndRow = dtbatch.Rows.Count;

                    }


                    int numRows = 0;
                    for (int x = StartRow; x < EndRow; x++)
                    {
                        if (numRows == 0) //Heading
                        {
                            PdfPCell loRptNameHeading = new PdfPCell();
                            PdfPCell loPageNoHeading = new PdfPCell();
                            PdfPCell loBlank = new PdfPCell();

                            Paragraph lochunk1Heading = new Paragraph(" Report ", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));

                            Paragraph lochunk2Heading = new Paragraph("Page", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                            lochunk2Heading.SetAlignment("center");

                            Paragraph lochunkBlank = new Paragraph(" ", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                            lochunkBlank.SetAlignment("center");

                            Anchor targetindex = new Anchor("Index", setFontsAll(4, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                            //targetindex.Reference = "#Index";

                            string key = "Index";
                            targetindex.Name = key.ToString();


                            PdfOutline Bookmarkindex = writer.RootOutline;
                            PdfOutline mbot1 = new PdfOutline(Bookmarkindex, PdfAction.GotoLocalPage("Index", false), "TABLE OF CONTENTS");

                            lochunkBlank.Add(targetindex);

                            //lochunk1Heading.Leading = 11f;
                            //    lochunk1Heading.SetLeading(15f, 5f);
                            //    lochunk2Heading.SetLeading(15f, 5f);


                            loRptNameHeading.AddElement(lochunk1Heading);
                            loPageNoHeading.AddElement(lochunk2Heading);
                            loBlank.AddElement(lochunkBlank);

                            loRptNameHeading.PaddingBottom = 5f;
                            loPageNoHeading.PaddingBottom = 5f;
                            loBlank.PaddingBottom = 5f;

                            loRptNameHeading.Border = 0;
                            loPageNoHeading.Border = 0;
                            loBlank.Border = 0;

                            loRptNameHeading.BorderWidthBottom = 0.7f;
                            loPageNoHeading.BorderWidthBottom = 0.7f;

                            loRptNameHeading.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"));
                            loPageNoHeading.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"));

                            tableindex.AddCell(loRptNameHeading);
                            tableindex.AddCell(loPageNoHeading);
                            tableindex.AddCell(loBlank);
                        }

                        //Table Content 
                        PdfPCell loRptName = new PdfPCell();    //FirstColumn : Report Name 
                        PdfPCell loPageNo = new PdfPCell();     //SecondColumn : Report Name 
                        PdfPCell loBlankCol = new PdfPCell();

                        string strRptname = getReportName(dtbatch, x);



                        Anchor target = new Anchor(strRptname, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                        target.Reference = "#" + dtbatch.Rows[x]["ssi_greshamreportidname"].ToString().Trim().Replace(" ", "_") + dtbatch.Rows[x]["numPageNo"].ToString();

                        PdfOutline root = writer.RootOutline;
                        PdfOutline mbot = new PdfOutline(root, PdfAction.GotoLocalPage(dtbatch.Rows[x]["ssi_greshamreportidname"].ToString().Trim().Replace(" ", "_") + dtbatch.Rows[x]["numPageNo"].ToString(), false), strRptname);


                        string Title = " ";

                        Paragraph lochunk1 = new Paragraph(" " + Title, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                        lochunk1.Add(target);
                        //lochunk1.SetAlignment("middle");


                        int pagenum = Convert.ToInt32(dtbatch.Rows[x]["numPageNo"]);

                        Paragraph lochunk2 = new Paragraph(Convert.ToString(pagenum + numIndexPageCount), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                        lochunk2.SetAlignment("center");


                        Paragraph lochunkBlank1 = new Paragraph(" ", setFontsAll(10, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                        lochunkBlank1.SetAlignment("center");

                        //lochunk2.SetAlignment("middle");

                        loRptName.AddElement(lochunk1);
                        loPageNo.AddElement(lochunk2);
                        loBlankCol.AddElement(lochunkBlank1);

                        loRptName.Border = 0;
                        loPageNo.Border = 0;
                        loBlankCol.Border = 0;
                        loRptName.PaddingBottom = 5f;
                        loPageNo.PaddingBottom = 5f;



                        loRptName.BorderWidthBottom = 0.7f;
                        loPageNo.BorderWidthBottom = 0.7f;

                        loRptName.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));
                        loPageNo.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C0C0C0"));

                        loRptName.VerticalAlignment = Element.ALIGN_MIDDLE;
                        loPageNo.VerticalAlignment = Element.ALIGN_MIDDLE;

                        tableindex.AddCell(loRptName);
                        tableindex.AddCell(loPageNo);
                        tableindex.AddCell(loBlankCol);

                        //tableindex.HorizontalAlignment = 0;


                        numRows++;
                    }

                    PdfPTable tableTitle = new PdfPTable(1);
                    int[] width = { 100 };
                    tableTitle.SetWidths(width);

                    //tableTitle.TotalWidth = 100f;
                    tableTitle.WidthPercentage = 60f;
                    Paragraph pTitle = new Paragraph("TABLE OF CONTENTS", setFontsAll(14, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("000000"))));
                    pTitle.SetAlignment("center");
                    pTitle.IndentationLeft = 265f;
                    PdfPCell loRTitle = new PdfPCell();
                    loRTitle.Border = 0;
                    loRTitle.PaddingBottom = 5f;
                    loRTitle.PaddingTop = 15f;
                    loRTitle.AddElement(pTitle);
                    tableTitle.AddCell(loRTitle);
                    tableTitle.HorizontalAlignment = 0;
                    string blanktitle = "";
                    Anchor targetind = new Anchor("A ", whiteFont);
                    targetind.Reference = "#Index";
                    //  string key = dicNumFilesCount.FirstOrDefault(x => x.Value == i - 2).Key;

                    //  targetind.Name = "Index";



                    Paragraph lochunkindex = new Paragraph("\n" + blanktitle);


                    lochunkindex.Add(blanktitle);

                    document.Add(lochunkindex); //blank chunk to target table of content 

                    document.Add(tableTitle);
                    document.Add(tableindex);

                    //string Title = " ";
                    //Anchor target = new Anchor("Portfolio Construction Chart v2.1");
                    //target.Reference = "#pc2";

                    //Paragraph lochunk1 = new Paragraph("\n" + Title);
                    //lochunk1.Add(target);

                    //Anchor target1 = new Anchor("Asset Distribution");
                    //target1.Reference = "#AD";

                    //Paragraph lochunk2 = new Paragraph("\n" + Title);
                    //lochunk2.Add(target1);

                    //document.Add(lochunk1);
                    //document.Add(lochunk2);
                }
                else if (i > numIndexPageCount)
                {
                    int numpage = i - numIndexPageCount;
                    DataRow[] rows = dtbatch.Select("numPageNo = " + numpage.ToString() + "");
                    if (rows.Length > 0)
                    {
                        string Title = "";
                        Anchor target = new Anchor("A ", whiteFont);

                        //  string key = dicNumFilesCount.FirstOrDefault(x => x.Value == i - 2).Key;
                        string key = rows[0]["ssi_greshamreportidname"].ToString() + rows[0]["numPageNo"].ToString();
                        target.Name = key.ToString().Trim().Replace(" ", "_");



                        Paragraph lochunk1 = new Paragraph("\n" + Title);


                        lochunk1.Add(target);

                        document.Add(lochunk1);
                    }



                    PdfContentByte cb = writer.DirectContent;
                    ColumnText ct = new ColumnText(cb);
                    ct.SetSimpleColumn(new Phrase(new Chunk("Page " + i + " of " + reader.NumberOfPages, setFontsAll(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))))), 480, 15, 400, 40, 25, Element.ALIGN_CENTER | Element.ALIGN_BOTTOM);
                    ct.Go();
                }
                String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

                PdfContentByte cb1 = writer.DirectContent;
                ColumnText ct1 = new ColumnText(cb1);
                ct1.SetSimpleColumn(new Phrase(new Chunk(lsDateTime, setFontsAll(8, 0, 1, new iTextSharp.text.Color(216, 216, 216)))), 800, 15, 725, 40, 25, Element.ALIGN_RIGHT | Element.ALIGN_BOTTOM);
                ct1.Go();

                // var multiLineString = "" + i.ToString() + "!".Split('\n');

                // foreach (var line in multiLineString)
                //  {
                // contentByte.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, "Page " + i + " of " + reader.NumberOfPages, 825, 10, 0);
                //  contentByte.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, " ", 825, 10, 0);
                // }

                // contentByte.EndText();
                iTextSharp.text.Rectangle psize = reader.GetPageSizeWithRotation(i);
                switch (psize.Rotation)
                {
                    case 0:
                        contentByte.AddTemplate(importedPage, 1f, 0, 0, 1f, 0, 0);
                        break;
                    case 90:
                        contentByte.AddTemplate(importedPage, 0, -1f, 1f, 0, 0, psize.Height);
                        break;
                    case 180:
                        contentByte.AddTemplate(importedPage, -1f, 0, 0, -1f, 0, 0);
                        break;
                    case 270:
                        contentByte.AddTemplate(importedPage, 0, 1.0F, -1.0F, 0, psize.Width, 0);
                        break;
                    default:
                        break;
                }


            }

            document.Close();
            writer.Close();
            return filename;
        }
        catch (Exception ex)
        {
            document.Close();
            //writer.Close();
            return "";

        }

    }

    protected string getReportName(DataTable foTable, int j)
    {
        string ReportingID = Convert.ToString(foTable.Rows[j]["ssi_GreshamReportId"]);
        string strRptName = "";
        String lsAllocationGroupNEW = Convert.ToString(foTable.Rows[j]["Ssi_AllocationGroup"]);
        string TempFilePath = Convert.ToString(foTable.Rows[j]["ssi_TemplateFilePath"]);

        String lsFinalTitleAfterChange = String.Empty;
        if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"])))
            lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"]);

        if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"])))
            lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"]);

        String ReportName = Convert.ToString(foTable.Rows[j]["ssi_GreshamReportIdName"]);
        if (ReportName == "Client Goals" || ReportName == "Absolute Returns" || ReportName == "Capital Protection")
        {
            if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["Ssi_HouseholdIdName"])))
            {
                lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["Ssi_HouseholdIdName"]);
            }
        }
        //added 5_20_2019 -- LegalEntity -- Title
        else if (ReportingID.ToUpper() == "AFD08C8B-2E25-E911-8106-000D3A1C025B" || ReportingID.ToUpper() == "806E4D33-1D29-E911-8106-000D3A1C025B" || ReportingID.ToUpper() == "90D6C145-1D29-E911-8106-000D3A1C025B" || ReportingID.ToUpper() == "A47E365E-1D29-E911-8106-000D3A1C025B") //Private Equity Performance||Private REal Asset Performance||Outside Private Equity Performance||Outside Private REal Asset Performance
        {
            if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["Ssi_LegalEntityIdName"])))
            {
                lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["Ssi_LegalEntityIdName"]);
            }
        }

        string fsGAorTIAflag = Convert.ToString(foTable.Rows[j]["ssi_gaortia"]);
        string fsDiscretionaryFlg = Convert.ToString(foTable.Rows[j]["Discretionary Flag"]);

        if (fsGAorTIAflag == "GA")
        {
            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                strRptName = "GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + " - Discretionary: " + lsFinalTitleAfterChange;
            else
                strRptName = "GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + ": " + lsFinalTitleAfterChange;
        }
        else
        {
            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                strRptName = Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + " - Discretionary: " + lsFinalTitleAfterChange;
            else
                strRptName = Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + ": " + lsFinalTitleAfterChange;
        }

        if (TempFilePath != "")
            strRptName = Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "");

        return strRptName;
    }

    public void SetTotalPageCount(string RptName)
    {
        try
        {


            if (HttpContext.Current.Session["CurPageInBatch"] != null)
            {
                HttpContext.Current.Session["CurPageInBatch"] = (int)HttpContext.Current.Session["CurPageInBatch"] + 1;
                //if (HttpContext.Current.Session["BatchDic"] != null)
                //{
                //    Dictionary<string, int> dicbatchFilesCount = (Dictionary<string, int>)HttpContext.Current.Session["BatchDic"];
                //    if (!dicbatchFilesCount.ContainsKey(RptName))
                //    {
                //        dicbatchFilesCount.Add(RptName, (int)HttpContext.Current.Session["CurPageInBatch"]);
                //    }
                //}

            }
            else
            {
                HttpContext.Current.Session["CurPageInBatch"] = 1;
                //Dictionary<string, int> dicbatchFilesCount = new Dictionary<string, int>();
                //dicbatchFilesCount.Add(RptName, (int)HttpContext.Current.Session["CurPageInBatch"]);
                //HttpContext.Current.Session["BatchDic"] = dicbatchFilesCount;
            }
        }
        catch (Exception ex)
        {

        }
    }

    public int GetPageCountFromPDF(string filePath)
    {
        try
        {
            PdfReader pdfReader = new PdfReader(filePath);
            int numberOfPages = pdfReader.NumberOfPages;
            return numberOfPages;
        }
        catch (Exception exe)
        {
            return 0;
        }
    }
    public iTextSharp.text.Table addFooter(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt, String FooterLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        iTextSharp.text.Table fotable = new iTextSharp.text.Table(2, 1);
        fotable.Width = 100;
        fotable.Border = 0;
        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        fotable.Cellpadding = 0;
        Cell loCell = new Cell();
        Chunk loChunk = new Chunk();
        int EndOfReportPageCnt = 3;
        bool flagClientFooter = false;
        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                FooterLocation = "100000000";
                if (FooterLocation == "100000001")
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
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt - 1; i++)
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
                    //loCell.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.Add(loChunk);
                    fotable.AddCell(loCell);

                    int pageFootercnt = liPageSize - 2 - liLastPageData;

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
                        flagClientFooter = true;
                    }


                    if (!flagClientFooter)
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
                            flagClientFooter = true;
                        }
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
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt - 1; i++)
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
        //loChunk = new Chunk(lsDateTime + "                     ", Font8GreyItalic());//Commented -- FooterLogic
        loChunk = new Chunk("                     ", Font8GreyItalic());
        loCell.Add(loChunk);
        loCell.Leading = 8f;//25f
        loCell.BorderWidth = 0;
        loCell.Colspan = 2;
        loCell.HorizontalAlignment = 2;// iTextSharp.text.Cell.ALIGN_RIGHT;
        fotable.AddCell(loCell);
        //fotable.TableFitsPage = true;

        return fotable;
    }

    public PdfPTable addFooterAsset(String lsDateTime, Boolean footerflg, String FooterTxt, String FooterLocation, bool isPageNo, int liCurrentPage, int liTotalPages, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);

        string Footer_Location = FooterLocation;
        fotable.TotalWidth = 100f;

        fotable.WidthPercentage = 100f;

        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        //  fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();
        if (isPageNo)
        {
            loCell = new PdfPCell();
            //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
            loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
            loChunk.Leading = 15f;//25f
            loChunk.SetAlignment("center");
            loCell.HorizontalAlignment = 2;
            loCell.BorderWidth = 0;
            loCell.Colspan = 2;
            loCell.PaddingBottom = 15f;
            loCell.AddElement(loChunk);
            fotable.AddCell(loCell);
        }


        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                Footer_Location = "100000000";
            }


            if (Footer_Location == "100000001")
            {
                #region End Of Report
                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                //loCell = new PdfPCell();
                ////loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                //loChunk = new Paragraph("");
                ////loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                ////loCell.PaddingTop = -8f;
                ////loChunk.Leading = 15f;//25f
                //loCell.HorizontalAlignment = 0;
                //loCell.Colspan = 2;
                //loCell.BorderWidth = 0;
                //fotable.AddCell(loCell);

                //loCell = new PdfPCell();
                ////loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                //loChunk = new Paragraph("");
                ////loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                ////loCell.PaddingTop = -8f;
                ////loChunk.Leading = 15f;//25f
                //loCell.HorizontalAlignment = 0;
                //loCell.Colspan = 2;
                //loCell.BorderWidth = 0;
                //fotable.AddCell(loCell);

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));

                if (Ssi_GreshamClientFooter == "2")
                    loCell.PaddingTop = -15f;
                //   loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion
            }

            else if (Footer_Location == "100000000" && Ssi_GreshamClientFooter != "3")
            {
                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //   loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
            }
        }

        if (Ssi_GreshamClientFooter != "3")
        {

            loCell = new PdfPCell();
            //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
            loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());

            loCell.AddElement(loChunk);
            loChunk.SetAlignment("right");
            loCell.PaddingTop = -8f;
            loCell.BorderWidth = 0;
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 2;// iTextSharp.text.Cell.ALIGN_RIGHT;
                                           // fotable.AddCell(loCell);//Commented -- FooterLogic
                                           //fotable.TableFitsPage = true;
        }
        return fotable;
    }


    public PdfPTable addFooterAsset1(String lsDateTime, Boolean footerflg, String FooterTxt, String FooterLocation, bool isPageNo, int liCurrentPage, int liTotalPages, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);


        fotable.TotalWidth = 100f;

        fotable.WidthPercentage = 100f;

        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        //  fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();
        if (isPageNo)
        {
            loCell = new PdfPCell();
            //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
            loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
            loChunk.Leading = 15f;//25f
            loChunk.SetAlignment("center");
            loCell.HorizontalAlignment = 2;
            loCell.BorderWidth = 0;
            loCell.Colspan = 2;
            loCell.PaddingBottom = 15f;
            loCell.AddElement(loChunk);
            fotable.AddCell(loCell);
        }


        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "3")
            {
                if (FooterLocation == "100000000")
                {
                    FooterText = ClientFooterTxt + "\n" + FooterText;
                }
            }




            if (FooterLocation == "100000000")
            {
                #region Default Location

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterText, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion
            }

            else
            {
                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

            }
        }

        loCell = new PdfPCell();
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        if (isPageNo)
            loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages + "                     ", Font7Normal());
        else
            loChunk = new Paragraph(" " + liTotalPages, Font7Normal());
        loChunk.Leading = 15f;//25f
        loChunk.SetAlignment("right");

        // loCell.HorizontalAlignment = 2;
        loCell.BorderWidth = 0;

        //loCell.Colspan = 2;
        // loCell.PaddingBottom = 15f;
        loCell.AddElement(loChunk);
        // fotable.AddCell(loCell); //Commented -- FooterLogic


        loCell = new PdfPCell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());
        loChunk.Leading = 15f;//25f
        loCell.AddElement(loChunk);
        loChunk.SetAlignment("right");
        loCell.PaddingTop = -8f;
        loCell.BorderWidth = 0;
        return fotable;
    }


    public PdfPTable addFooterClientGoal(String lsDateTime, Boolean footerflg, String FooterTxt, String FooterLocation, bool isPageNo, int liCurrentPage, int liTotalPages, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);

        string Footer_Location = FooterLocation;
        //  fotable.TotalWidth = 100f;

        // fotable.WidthPercentage = 100f;

        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        //  fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();
        if (isPageNo)
        {
            loCell = new PdfPCell();
            //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
            loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
            loChunk.Leading = 15f;//25f
            loChunk.SetAlignment("center");
            loCell.HorizontalAlignment = 2;
            loCell.BorderWidth = 0;
            loCell.Colspan = 2;
            loCell.PaddingBottom = 15f;
            loCell.AddElement(loChunk);
            fotable.AddCell(loCell);
        }


        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                Footer_Location = "100000000";
            }


            if (Footer_Location == "100000001")
            {
                #region End Of Report
                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                //loCell = new PdfPCell();
                ////loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                //loChunk = new Paragraph("");
                ////loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                ////loCell.PaddingTop = -8f;
                ////loChunk.Leading = 15f;//25f
                //loCell.HorizontalAlignment = 0;
                //loCell.Colspan = 2;
                //loCell.BorderWidth = 0;
                //fotable.AddCell(loCell);

                //loCell = new PdfPCell();
                ////loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                //loChunk = new Paragraph("");
                ////loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                ////loCell.PaddingTop = -8f;
                ////loChunk.Leading = 15f;//25f
                //loCell.HorizontalAlignment = 0;
                //loCell.Colspan = 2;
                //loCell.BorderWidth = 0;
                //fotable.AddCell(loCell);

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));

                if (Ssi_GreshamClientFooter == "2")
                    loCell.PaddingTop = -15f;
                //   loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion
            }

            else if (Footer_Location == "100000000" && Ssi_GreshamClientFooter != "3")
            {
                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //   loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
            }
        }

        if (Ssi_GreshamClientFooter != "3")
        {

            loCell = new PdfPCell();
            //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
            loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());

            loCell.AddElement(loChunk);
            loChunk.SetAlignment("right");
            loCell.PaddingTop = -8f;
            loCell.BorderWidth = 0;
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 2;// iTextSharp.text.Cell.ALIGN_RIGHT;
                                           // fotable.AddCell(loCell);//Commented -- FooterLogic
                                           //fotable.TableFitsPage = true;
        }
        return fotable;
    }


    public PdfPTable addFooterClientGoal1(String lsDateTime, Boolean footerflg, String FooterTxt, String FooterLocation, bool isPageNo, int liCurrentPage, int liTotalPages, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);


        fotable.TotalWidth = 100f;

        fotable.WidthPercentage = 100f;

        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        //  fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();
        if (isPageNo)
        {
            loCell = new PdfPCell();
            //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
            loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
            loChunk.Leading = 15f;//25f
            loChunk.SetAlignment("center");
            loCell.HorizontalAlignment = 2;
            loCell.BorderWidth = 0;
            loCell.Colspan = 2;
            loCell.PaddingBottom = 15f;
            loCell.AddElement(loChunk);
            fotable.AddCell(loCell);
        }


        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "3")
            {
                if (FooterLocation == "100000000")
                {
                    FooterText = ClientFooterTxt + "\n" + FooterText;
                }
            }




            if (FooterLocation == "100000000")
            {
                #region Default Location

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterText, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion
            }

            else
            {
                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

            }
        }

        loCell = new PdfPCell();
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        if (isPageNo)
            loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages + "                     ", Font7Normal());
        else
            loChunk = new Paragraph(" " + liTotalPages, Font7Normal());
        loChunk.Leading = 15f;//25f
        loChunk.SetAlignment("right");

        // loCell.HorizontalAlignment = 2;
        loCell.BorderWidth = 0;

        //loCell.Colspan = 2;
        // loCell.PaddingBottom = 15f;
        loCell.AddElement(loChunk);
        // fotable.AddCell(loCell); //Commented -- FooterLogic


        loCell = new PdfPCell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());
        loChunk.Leading = 15f;//25f
        loCell.AddElement(loChunk);
        loChunk.SetAlignment("right");
        loCell.PaddingTop = -8f;
        loCell.BorderWidth = 0;
        return fotable;
    }

    public PdfPTable addFooterMGR(String lsDateTime, Boolean footerflg, String FooterTxt, String FooterLocation, bool isPageNo, int liCurrentPage, int liTotalPages, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);
        string Footer_Location = FooterLocation;

        fotable.TotalWidth = 100f;

        fotable.WidthPercentage = 100f;

        int[] headerwidths = { 57, 43 };
        fotable.SetWidths(headerwidths);
        //  fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();

        Paragraph PBlank = new Paragraph(" ");
        //FooterLocation = "100000001";
        if (footerflg)
        {
            //if (Ssi_GreshamClientFooter == "2")
            //{
            //    loCell = new PdfPCell();
            //    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
            //    loChunk = new Paragraph(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            //    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            //    loCell.PaddingTop = -12f;
            //    //loChunk.Leading = 15f;//25f
            //    loCell.HorizontalAlignment = 0;
            //    loCell.Colspan = 2;
            //    loCell.BorderWidth = 0;

            //    loCell.AddElement(loChunk);
            //    fotable.AddCell(loCell);



            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                Footer_Location = "100000000";
            }


            if (Footer_Location == "100000001")
            {
                #region End Of Report
                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

                loCell = new PdfPCell();
                loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                loCell.HorizontalAlignment = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                //loCell = new PdfPCell();
                ////loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                //loChunk = new Paragraph("");
                ////loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                ////loCell.PaddingTop = -8f;
                ////loChunk.Leading = 15f;//25f
                //loCell.HorizontalAlignment = 0;
                //loCell.Colspan = 2;
                //loCell.BorderWidth = 0;
                //fotable.AddCell(loCell);

                //loCell = new PdfPCell();
                ////loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                //loChunk = new Paragraph("");
                ////loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                ////loCell.PaddingTop = -8f;
                ////loChunk.Leading = 15f;//25f
                //loCell.HorizontalAlignment = 0;
                //loCell.Colspan = 2;
                //loCell.BorderWidth = 0;
                //fotable.AddCell(loCell);

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));

                if (Ssi_GreshamClientFooter == "2")
                    loCell.PaddingTop = -15f;
                //   loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion
            }

            else if (Footer_Location == "100000000" && Ssi_GreshamClientFooter != "3")
            {
                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //   loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
            }






        }


        if (Ssi_GreshamClientFooter != "3")
        {
            loCell = new PdfPCell();
            //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
            if (isPageNo)
                loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages + "                     ", Font7Normal());
            else
                loChunk = new Paragraph(" " + liTotalPages, Font7Normal());
            loChunk.Leading = 15f;//25f
            loChunk.SetAlignment("right");

            // loCell.HorizontalAlignment = 2;
            loCell.BorderWidth = 0;

            //loCell.Colspan = 2;
            // loCell.PaddingBottom = 15f;
            loCell.AddElement(loChunk);
            // fotable.AddCell(loCell); //Commented -- FooterLogic


            loCell = new PdfPCell();
            //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
            loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());
            loChunk.Leading = 15f;//25f
            loCell.AddElement(loChunk);
            loChunk.SetAlignment("right");
            loCell.PaddingTop = -8f;
            loCell.BorderWidth = 0;
        }
        //loCell.Colspan = 2;
        // loCell.HorizontalAlignment = 2;// iTextSharp.text.Cell.ALIGN_RIGHT;
        // fotable.AddCell(loCell); //Commented -- FooterLogic
        //fotable.TableFitsPage = true;

        return fotable;
    }

    public PdfPTable addFooterMGR1(String lsDateTime, Boolean footerflg, String FooterTxt, String FooterLocation, bool isPageNo, int liCurrentPage, int liTotalPages, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);


        fotable.TotalWidth = 100f;

        fotable.WidthPercentage = 100f;

        int[] headerwidths = { 57, 43 };
        fotable.SetWidths(headerwidths);
        //  fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();

        Paragraph PBlank = new Paragraph(" ");
        //FooterLocation = "100000001";
        if (footerflg)
        {
            //if (Ssi_GreshamClientFooter == "2")
            //{
            //    loCell = new PdfPCell();
            //    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
            //    loChunk = new Paragraph(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            //    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            //    loCell.PaddingTop = -12f;
            //    //loChunk.Leading = 15f;//25f
            //    loCell.HorizontalAlignment = 0;
            //    loCell.Colspan = 2;
            //    loCell.BorderWidth = 0;

            //    loCell.AddElement(loChunk);
            //    fotable.AddCell(loCell);

            if (Ssi_GreshamClientFooter == "3")
            {
                if (FooterLocation == "100000000")
                {
                    FooterText = ClientFooterTxt + "\n" + FooterText;
                }
            }




            if (FooterLocation == "100000000")
            {
                #region Default Location

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterText, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion
            }

            else
            {
                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loCell.PaddingTop = -12f;
                //loChunk.Leading = 15f;//25f
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;

                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);

            }






        }



        loCell = new PdfPCell();
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        if (isPageNo)
            loChunk = new Paragraph("Page " + liCurrentPage + " of " + liTotalPages + "                     ", Font7Normal());
        else
            loChunk = new Paragraph(" " + liTotalPages, Font7Normal());
        loChunk.Leading = 15f;//25f
        loChunk.SetAlignment("right");

        // loCell.HorizontalAlignment = 2;
        loCell.BorderWidth = 0;

        //loCell.Colspan = 2;
        // loCell.PaddingBottom = 15f;
        loCell.AddElement(loChunk);
        // fotable.AddCell(loCell); //Commented -- FooterLogic


        loCell = new PdfPCell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());
        loChunk.Leading = 15f;//25f
        loCell.AddElement(loChunk);
        loChunk.SetAlignment("right");
        loCell.PaddingTop = -8f;
        loCell.BorderWidth = 0;
        //loCell.Colspan = 2;
        // loCell.HorizontalAlignment = 2;// iTextSharp.text.Cell.ALIGN_RIGHT;
        // fotable.AddCell(loCell); //Commented -- FooterLogic
        //fotable.TableFitsPage = true;

        return fotable;
    }


    public PdfPTable addFooter(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt, string strNew, String FooterLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);
        //fotable.Width = 97;
        //fotable.Border = 0;
        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        //   fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();
        int EndOfReportPageCnt = 2;

        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                FooterLocation = "100000000";
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }





                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    for (int i = 0; i < EndOfReportPageCnt + 2; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);


                    #endregion
                }
            }
            else if (Ssi_GreshamClientFooter == "3")
            {
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt - 1; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }





                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 15f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    FooterTxt = ClientFooterTxt + "\n" + FooterTxt;
                    for (int i = 0; i < EndOfReportPageCnt + 2; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);

                    #endregion
                }
            }


            else if (Ssi_GreshamClientFooter == "1")
            {
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt - 1; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }





                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    for (int i = 0; i < EndOfReportPageCnt + 2; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);


                    #endregion
                }
            }

            else if (Ssi_GreshamClientFooter == "4")
            {

                #region For NONE
                for (int i = 0; i < EndOfReportPageCnt + 1; i++)
                {
                    loCell = new PdfPCell();
                    loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                    loCell.HorizontalAlignment = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);
                    loCell = new PdfPCell();
                    loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                    loCell.HorizontalAlignment = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);
                }

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loChunk.Leading = 8f;
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion

            }
        }


        loCell = new PdfPCell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        //loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());//Commented -- FooterLogic
        loChunk = new Paragraph("                     ", Font8GreyItalic());
        loChunk.SetAlignment("right");
        loCell.AddElement(loChunk);
        loChunk.Leading = 8f;//25f
        loCell.BorderWidth = 0;
        loCell.PaddingTop = 8f;
        // loCell.Colspan = 2;
        loCell.HorizontalAlignment = 2;// iTextSharp.text.Cell.ALIGN_RIGHT;
        fotable.AddCell(loCell);
        //fotable.TableFitsPage = true;

        return fotable;
    }

    public PdfPTable addFooterAbsoluteReturn(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt, string strNew, String FooterLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);
        //fotable.Width = 97;
        //fotable.Border = 0;
        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        //   fotable.Cellpadding = 0;
        PdfPCell loCell = new PdfPCell();
        Paragraph loChunk = new Paragraph();
        int EndOfReportPageCnt = 1;

        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                FooterLocation = "100000000";
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);

                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }





                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    // loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    for (int i = 0; i < EndOfReportPageCnt + 1; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);

                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);


                    #endregion
                }
            }
            else if (Ssi_GreshamClientFooter == "3")
            {
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report


                    //for (int i = 0; i < EndOfReportPageCnt; i++)
                    //{
                    //    loCell = new PdfPCell();
                    //    loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                    //    loCell.HorizontalAlignment = 2;
                    //    loCell.BorderWidth = 0;
                    //    loCell.AddElement(loChunk);
                    //    fotable.AddCell(loCell);
                    //    loCell = new PdfPCell();
                    //    loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                    //    loCell.HorizontalAlignment = 2;
                    //    loCell.BorderWidth = 0;
                    //    loCell.AddElement(loChunk);
                    //    fotable.AddCell(loCell);
                    //}





                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 14f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 12f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    FooterTxt = ClientFooterTxt + "\n" + FooterTxt;
                    for (int i = 0; i < EndOfReportPageCnt + 1; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);

                    #endregion
                }
            }


            else if (Ssi_GreshamClientFooter == "1")
            {
                if (FooterLocation == "100000001")
                {
                    #region Footer on End Report

                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 14f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    for (int i = 0; i < EndOfReportPageCnt + 1; i++)
                    {
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new PdfPCell();
                        loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.AddElement(loChunk);
                        fotable.AddCell(loCell);
                    }


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);


                    #endregion
                }
            }

            else if (Ssi_GreshamClientFooter == "4")
            {

                #region For NONE
                for (int i = 0; i < EndOfReportPageCnt + 1; i++)
                {
                    loCell = new PdfPCell();
                    loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                    loCell.HorizontalAlignment = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);
                    loCell = new PdfPCell();
                    loChunk = new Paragraph("dev", Font8Whitecheck("test"));
                    loCell.HorizontalAlignment = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);
                }

                loCell = new PdfPCell();
                //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                loChunk = new Paragraph(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                loChunk.Leading = 8f;
                loCell.HorizontalAlignment = 0;
                loCell.Colspan = 2;
                loCell.BorderWidth = 0;
                loCell.AddElement(loChunk);
                fotable.AddCell(loCell);
                #endregion

            }
        }


        loCell = new PdfPCell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        //loChunk = new Paragraph(lsDateTime + "                     ", Font8GreyItalic());//Commented -- FooterLogic
        loChunk = new Paragraph("                     ", Font8GreyItalic());
        loChunk.SetAlignment("right");
        loCell.AddElement(loChunk);
        loChunk.Leading = 8f;//25f
        loCell.BorderWidth = 0;
        loCell.PaddingTop = 8f;
        // loCell.Colspan = 2;
        loCell.HorizontalAlignment = 2;// iTextSharp.text.Cell.ALIGN_RIGHT;
        fotable.AddCell(loCell);
        //fotable.TableFitsPage = true;

        return fotable;
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


    #endregion

    #region Formatting Related Code like font ,border and colors

    private void SetBorder(Cell foCell, bool IsTop, bool IsBottom, bool IsLeft, bool IsRight)
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
        string fontpath = HttpContext.Current.Server.MapPath(".");
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
    // forth folio const chart old or v1
    public void setTableProperty1(iTextSharp.text.Table fotable)
    {
        //int[] headerwidths = { 28, 9, 9, 9, 9, 9, 9, 9, 7 };

        setWidthsoftable1(fotable);

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
                fotable.Width = 53;
                break;
            case "5":
                int[] headerwidths5 = { 30, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths5);
                fotable.Width = 67;
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

    public void setWidthsoftable1(iTextSharp.text.Table fotable)
    {

        switch (lsTotalNumberofColumns)
        {
            case "22":
                int[] headerwidths2 = { 30, 9 };
                fotable.SetWidths(headerwidths2);
                fotable.Width = 40;
                break;
            case "3":
                int[] headerwidths3 = { 31, 31, 31 };
                fotable.SetWidths(headerwidths3);
                fotable.Width = 94;
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
                int[] headerwidths9 = { 33, 8, 8, 8, 8, 8, 8, 8, 8 };
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
            case "13":
                int[] headerwidths13 = { 12, 2, 13, 2, 18, 18, 2, 15, 2, 13, 13, 2, 13 };
                fotable.SetWidths(headerwidths13);
                fotable.Width = 96; break;
            case "12":
                int[] headerwidths22 = { 18, 2, 13, 2, 14, 14, 2, 15, 0, 13, 13, 2, 18 };
                fotable.SetWidths(headerwidths22);
                fotable.Width = 80; break;
            case "21":
                int[] headerwidths12 = { 7, 7, 12, 7, 7, 7, 7, 35, 7, 7, 7, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                fotable.SetWidths(headerwidths12);
                fotable.Width = 150; break;
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

    #region Watermark Code
    public void WatermarkPdf(String src, String strFileLocationOut, string Image)
    {
        iTextSharp.text.Image img = null;
        float w = 0;
        float h = 0;

        PdfReader reader = new PdfReader(src);
        int n = reader.NumberOfPages;
        PdfStamper stamper = new PdfStamper(reader, new FileStream(strFileLocationOut, FileMode.Create, FileAccess.Write, FileShare.None));

        // text watermark

        // iTextSharp.text.Font f = new iTextSharp.text.Font(FontFamily.HELVETICA, 30);
        //iTextSharp.text.Font f = FontFactory.GetFont("Arial", 65, iTextSharp.text.Font.NORMAL, iTextSharp.text.Color.GRAY);
        iTextSharp.text.Font f = FontFactory.GetFont("Arial", 65, iTextSharp.text.Font.NORMAL, new iTextSharp.text.Color(145, 145, 145));
        Phrase p = new Phrase("Internal Use Only", f); //String to be set

        // transparency
        PdfGState gs1 = new PdfGState();
        gs1.FillOpacity = 0.5f;

        // properties
        PdfContentByte over;
        iTextSharp.text.Rectangle pagesize;
        float x = 0, y = 0;
        // loop over every page
        for (int i = 1; i <= n; i++)
        {
            pagesize = reader.GetPageSizeWithRotation(i);
            if (Image != "")//Image to be set
            {
                img = iTextSharp.text.Image.GetInstance(Image);
                over = stamper.GetUnderContent(i);
                over.SaveState();
                if (pagesize.Rotation == 90 || pagesize.Rotation == 360)//Landscape
                {
                    w = img.ScaledWidth / 2;
                    h = img.ScaledHeight / 3;

                    x = pagesize.Height / 2;
                    y = pagesize.Width / 2;
                    x = x + 150;
                    y = y - 125;
                }
                else if (pagesize.Rotation == 0 || pagesize.Rotation == 180)//Potrait
                {
                    w = img.ScaledWidth / 3;
                    h = img.ScaledHeight / 3;

                    x = pagesize.Height / 3;
                    y = pagesize.Width / 1;
                    x = x + 45;
                    y = y - 185;
                }
                over.SetGState(gs1);
                over.AddImage(img, w, 0, 0, h, x - (w / 2), y - (h / 2));
            }
            else//String to be set
            {
                if (pagesize.Rotation == 0 || pagesize.Rotation == 180)//Potrait
                {
                    //x = iTextSharp.text.PageSize.A4.Width / 2;
                    // y = iTextSharp.text.PageSize.A4.Height / 3;
                    x = pagesize.Height / 2;
                    y = pagesize.Width / 1;
                    x = x - 50f;
                    y = y - 200f;
                }
                else if (pagesize.Rotation == 90 || pagesize.Rotation == 360)//Landscape
                {
                    //x = iTextSharp.text.PageSize.A4.Width / 2;
                    // y = iTextSharp.text.PageSize.A4.Height / 3;
                    x = pagesize.Height / 2;
                    y = pagesize.Width / 3;
                    x = x + 125f;
                }

                over = stamper.GetOverContent(i);//.GetOverContent(i);
                over.SaveState();

                over.SetGState(gs1);
                ColumnText.ShowTextAligned(over, Element.ALIGN_CENTER, p, x, y, 45);
            }

            //if (i % 2 == 1)
            //    ColumnText.ShowTextAligned(over, Element.ALIGN_CENTER, p, x, y, 0);
            //else
            //    over.AddImage(img, w, 0, 0, h, x - (w / 2), y - (h / 2));
            over.RestoreState();
        }
        stamper.Close();
        reader.Close();
    }
    #endregion
}

