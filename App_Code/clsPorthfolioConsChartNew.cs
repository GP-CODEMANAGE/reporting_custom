using System;
using System.Data;
using System.Configuration;
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
using org.jfree.ui;
using org.jfree.chart.labels;

/// <summary>
/// Summary description for clsPorthfolioConsChartNew
/// </summary>
public class clsPorthfolioConsChartNew
{
    #region General Declaration
    //public string lsSQL = "";
    Boolean fbCheckExcel = false;
    public StreamWriter sw = null;
    public string strDescription = string.Empty;
    //bool bProceed = true;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    public int AllocationColumnSize = 22;
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName;

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
    private string _CommitmentReportHeader = string.Empty;
    private string _GreshamAdvisedFlag = string.Empty;


    private string _Footerlocation = string.Empty;

    private string _ClientFooterTxt = string.Empty;
    private string _Ssi_GreshamClientFooter = string.Empty;

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
    #endregion

    public string generatePDFFinal(DataSet newdataset, string TempFolderPath)
    {
        liPageSize = 29;

        DB clsDB = new DB();

        String lsFooterTxt = String.Empty;

        DataTable table = newdataset.Tables[1];
        Random rnd = new Random();
        string strRndNumber = Convert.ToString(rnd.Next(5));
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + strRndNumber;

        //   String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        int topmargin = 30;
        if (Convert.ToBoolean(newdataset.Tables[3].Rows[0]["ShowTargetFlg"]) == true && FooterText != "")
            topmargin = 25;

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 50, 50, topmargin, 6);//10,10
        // String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));

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
            string International_Equity_UUid = "B513B878-2071-E311-BDD3-0019B9E7EE05";
            string Global_Opportunistic_UUid = "8413896B-4925-DF11-B686-001D09665E8F";
            string Liquid_Real_Assets_UUid = "0332530A-1AD3-DF11-9789-0019B9E7EE05";
            /***********************************************************************************/
            string PrivEqty_UUID = "02FFE912-D704-DE11-A38C-001D09665E8F";
            string ConHold_UUID = "E2465B5C-40A7-4A35-B5EA-50A2C74CF6F5";
            string Illiquid_Real_Assets_UUID = "C2A2D71C-D704-DE11-A38C-001D09665E8F";

            string ColName = "sas_assetClassID";

            string MarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 1))));
            //string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2)))));
            string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2))));
            string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 1))));
            //string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal((GetFilteredValue(table, ConHold_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2)))));
            string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2))));

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
                        string UUID = "B513B878-2071-E311-BDD3-0019B9E7EE05";
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

            string StarategicPurposeChart = generateOverAllPieChart(newdataset.Tables[2], "1", TempFolderPath);
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(StarategicPurposeChart);

            string VolatileChart = generateOverAllPieChart(newdataset.Tables[4], "2", TempFolderPath);
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
            dashjpg.SetAbsolutePosition(138f, 308f);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

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

            Paragraph pSpace = new Paragraph();
            pSpace.Add("\n");
            document.Add(loTable);
            //document.Add(pSpace);
            document.Add(Table);
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
        else
        {
            //lblError.Text = "Record not found";
            return "Record not found";
        }
    }

    public string generatePDFFinal_New(DataSet newdataset, string TempFolderPath)
    {
        liPageSize = 29;

        DB clsDB = new DB();

        String lsFooterTxt = String.Empty;

        DataTable table = newdataset.Tables[1];
        Random rnd = new Random();
        string strRndNumber = Convert.ToString(rnd.Next());

        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + "_" + strRndNumber;

        // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".xls";

        int topmargin = 20;//25//30
        int bottommargin = 6;
        float dashImgYaxis = 304f;//299f
        if (Convert.ToBoolean(newdataset.Tables[3].Rows[0]["ShowTargetFlg"]) == true && FooterText != "")
        {
            if (newdataset.Tables[3].Rows.Count >= 17)
            {
                topmargin = 14;
                bottommargin = 3;
                dashImgYaxis = 310f;
            }
        }

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 50, 50, topmargin, bottommargin);//10,10
                                                                                                                                                //   String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        String ls = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
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
            //string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2)))));
            //string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2))));
            string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["Marketable"])));
            string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 1))));
            //string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal((GetFilteredValue(table, ConHold_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2)))));
            //string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2))));
            string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["NonMarketable"])));

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

                        if (GreshamAdvisedFlag == "GA")
                        {
                            lochunk = new Chunk("\nGRESHAM ADVISED ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell.Add(lochunk);
                        }
                        else
                        {
                            lochunk = new Chunk("\nTOTAL INVESTMENT ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell.Add(lochunk);
                        }

                        string Title = "Portfolio Construction";

                        lochunk = new Chunk("\n" + Title, setFontsAllFrutiger(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        //loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        loCell.Add(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.Leading = 13F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 1 && j == 0)
                    {
                        //the text of this line adjusted in above line due to space between header text.
                        //so this line is not adding in pdf
                        string Title = "";

                        lochunk = new Chunk("" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        //lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        //loCell.Add(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
                        loCell.Leading = 0F;
                        //loTable.AddCell(loCell);

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

            string StarategicPurposeChart = generateOverAllPieChart(newdataset.Tables[2], "1", TempFolderPath);
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(StarategicPurposeChart);

            string VolatileChart = generateOverAllPieChart(newdataset.Tables[4], "2", TempFolderPath);
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
            dashjpg.SetAbsolutePosition(138f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(421f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(703f, dashImgYaxis);
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
            clsCombinedReports clscombinedreports = new clsCombinedReports();
            clscombinedreports.SetTotalPageCount("Portfolio Construction Chart v2.1");


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

    public string generatePDFFinalV2OLD(DataSet newdataset, string TempFolderPath)
    {
        liPageSize = 29;

        DB clsDB = new DB();

        String lsFooterTxt = String.Empty;

        DataTable table = newdataset.Tables[1];
        DataTable dtSubAsset = newdataset.Tables[0];

        Random rnd = new Random();
        string strRndNumber = Convert.ToString(rnd.Next(5));
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + strRndNumber;

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        int topmargin = 20;//25//30
        int bottommargin = 6;
        float dashImgYaxis = 304f;//299f
        if (Convert.ToBoolean(Convert.ToInt16(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == true && FooterText != "")
        {
            if (newdataset.Tables[3].Rows.Count >= 17)
            {
                topmargin = 14;
                bottommargin = 1;
                dashImgYaxis = 310f;
            }
        }

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 50, 50, topmargin, bottommargin);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));



        AddFooter(document, FooterText);

        document.Open();
        PdfContentByte cb = writer.DirectContent;
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

            string Cash_UUid = "A11B900A-2071-E311-BDD3-0019B9E7EE05";
            string Fixed_Income_UUid = "F41C2839-2071-E311-BDD3-0019B9E7EE05";
            string Hedged_Strategies_UUid = "D5A0EEA3-2071-E311-BDD3-0019B9E7EE05";
            string Domestic_Equity_UUid = "00000000-0000-0000-0000-000000000000";
            // string International_Equity_UUid = "B513B878-2071-E311-BDD3-0019B9E7EE05";
            string Global_Opportunistic_UUid = "6D4F7558-2071-E311-BDD3-0019B9E7EE05";
            string Liquid_Real_Assets_UUid = "435F2391-2071-E311-BDD3-0019B9E7EE05";
            /***********************************************************************************/
            string PrivEqty_UUID = "3C2101F2-2071-E311-BDD3-0019B9E7EE05";
            string ConHold_UUID = "E298DD16-2071-E311-BDD3-0019B9E7EE05";
            string Illiquid_Real_Assets_UUID = "AF4FEE66-2071-E311-BDD3-0019B9E7EE05";

            //string ColName = "sas_assetClassID";
            string ColName = "ssi_subassetclassid";
            string MarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 1))));
            //string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2)))));
            //string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2))));
            string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["Marketable"])));
            string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 1))));
            //string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal((GetFilteredValue(table, ConHold_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2)))));
            //string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2))));
            string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["NonMarketable"])));

            double MarketPerc = Convert.ToDouble(MarketablePercent);
            double nonMarketPerc = Convert.ToDouble(NonMarketablePercent);
            if (MarketPerc > 100.0)
                MarketablePercent = "100.0";
            if (nonMarketPerc > 100.0)
                NonMarketablePercent = "100.0";

            cb.MoveTo(428, 450);
            cb.LineTo(603, 450);

            cb.MoveTo(470, 450);
            cb.LineTo(470, 409);

            cb.MoveTo(514, 450);
            cb.LineTo(514, 409);

            cb.MoveTo(560, 450);
            cb.LineTo(560, 409);

            //   cb.MoveTo(430, 435);
            // cb.LineTo(70, 300);
            //Path closed and stroked
            cb.SetColorStroke(new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F")));
            cb.SetLineWidth(1.5f);
            cb.ClosePathStroke();

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

                        if (GreshamAdvisedFlag == "GA")
                        {
                            lochunk = new Chunk("\nGRESHAM ADVISED ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell.Add(lochunk);
                        }
                        else
                        {
                            lochunk = new Chunk("\nTOTAL INVESTMENT ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell.Add(lochunk);
                        }

                        string Title = "Portfolio Construction";

                        lochunk = new Chunk("\n" + Title, setFontsAllFrutiger(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        //loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        loCell.Add(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.Leading = 13F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 1 && j == 0)
                    {
                        //the text of this line adjusted in above line due to space between header text.
                        //so this line is not adding in pdf
                        string Title = "";

                        lochunk = new Chunk("" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        //lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        //loCell.Add(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
                        loCell.Leading = 0F;
                        //loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 1) //Heading -- Left Boxes
                    {
                        string UUID = "A11B900A-2071-E311-BDD3-0019B9E7EE05";
                        string Hdr1 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        //lochunk = new Chunk(Hdr1 + Text + "", setFontsAll(9, 1, 0));
                        lochunk = new Chunk("HEAD", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                        loCell = new iTextSharp.text.Cell();

                        loCell.Colspan = 7;
                        j = j + 6;
                        loCell.Border = 0;
                        loCell.Add(lochunk);
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 7) //Dash Lines After Heading Left
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
                        string UUID = "E298DD16-2071-E311-BDD3-0019B9E7EE05";
                        string Hdr2 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        // lochunk = new Chunk(Hdr2 + Text + "", setFontsAll(9, 1, 0));
                        lochunk = new Chunk("HEAD", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
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
                        string UUID = "AF4FEE66-2071-E311-BDD3-0019B9E7EE05";
                        string Hdr3 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        //lochunk = new Chunk(Hdr3 + Text + "", setFontsAll(9, 1, 0));
                        lochunk = new Chunk("HEAD", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
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
                        string UUID = "A11B900A-2071-E311-BDD3-0019B9E7EE05";
                        string Cash = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Cash + GetSeprator(Cash) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        //loCell.BorderWidthBottom = 2F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 4)
                    {
                        string UUID = "F41C2839-2071-E311-BDD3-0019B9E7EE05";
                        string Fixed_Income = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Fixed_Income + GetSeprator(Fixed_Income) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 6) //Hedged Strategies Heading
                    {
                        string UUID = "D5A0EEA3-2071-E311-BDD3-0019B9E7EE05";
                        string Hedged_Strategies = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string HedgedStr = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Hedged_Strategies + GetSeprator(Hedged_Strategies) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + HedgedStr + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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
                        string UUID = "00000000-0000-0000-0000-000000000000";
                        string Global_UUID = "2394047C-7D1D-E411-8A68-0019B9E7EE05";
                        string US_UUID = "77A63F24-2071-E311-BDD3-0019B9E7EE05";
                        string International_UUID = "B513B878-2071-E311-BDD3-0019B9E7EE05";
                        string Emrging_UUID = "0C080B96-6A52-E511-940C-005056A0099E";


                        string strSubAssetClassName1 = GetFilteredValue(dtSubAsset, Global_UUID, "ssi_subassetclassid", 2);
                        string strSubAssetClassName2 = GetFilteredValue(dtSubAsset, US_UUID, "ssi_subassetclassid", 2);
                        string strSubAssetClassName3 = GetFilteredValue(dtSubAsset, International_UUID, "ssi_subassetclassid", 2);
                        string strSubAssetClassName4 = GetFilteredValue(dtSubAsset, Emrging_UUID, "ssi_subassetclassid", 2);

                        string strCurrentAllocation1 = RoundUp(GetFilteredValue(dtSubAsset, Global_UUID, "ssi_subassetclassid", 7));
                        string strCurrentAllocation2 = RoundUp(GetFilteredValue(dtSubAsset, US_UUID, "ssi_subassetclassid", 7));
                        string strCurrentAllocation3 = RoundUp(GetFilteredValue(dtSubAsset, International_UUID, "ssi_subassetclassid", 7));
                        string strCurrentAllocation4 = RoundUp(GetFilteredValue(dtSubAsset, Emrging_UUID, "ssi_subassetclassid", 7));

                        if (Convert.ToDecimal(strCurrentAllocation1) < 10)
                            strCurrentAllocation1 = strCurrentAllocation1 + "% ";
                        else
                            strCurrentAllocation1 = strCurrentAllocation1 + "%";

                        if (Convert.ToDecimal(strCurrentAllocation2) < 10)
                            strCurrentAllocation2 = strCurrentAllocation2 + "% ";
                        else
                            strCurrentAllocation2 = strCurrentAllocation2 + "%";

                        if (Convert.ToDecimal(strCurrentAllocation3) < 10)
                            strCurrentAllocation3 = strCurrentAllocation3 + "% ";
                        else
                            strCurrentAllocation3 = strCurrentAllocation3 + "%";
                        if (Convert.ToDecimal(strCurrentAllocation4) < 10)
                            strCurrentAllocation4 = strCurrentAllocation4 + "% ";
                        else
                            strCurrentAllocation4 = strCurrentAllocation4 + "%";


                        string strCurrentValue1 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, Global_UUID, "ssi_subassetclassid", 3)));
                        string strCurrentValue2 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, US_UUID, "ssi_subassetclassid", 3)));
                        string strCurrentValue3 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, International_UUID, "ssi_subassetclassid", 3)));
                        string strCurrentValue4 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, Emrging_UUID, "ssi_subassetclassid", 3)));

                        if (Convert.ToDecimal(strCurrentValue1) == 0)
                            strCurrentValue1 = "     $" + strCurrentValue1 + "        ";
                        else if (Convert.ToDecimal(strCurrentValue1) < 100000)
                            strCurrentValue1 = "   $" + strCurrentValue1 + "";
                        else if (Convert.ToDecimal(strCurrentValue1) < 1000000)
                            strCurrentValue1 = "  $" + strCurrentValue1 + "";
                        else if (Convert.ToDecimal(strCurrentValue1) < 10000000)
                            strCurrentValue1 = " $" + strCurrentValue1;
                        else
                            strCurrentValue1 = " $" + strCurrentValue1;

                        if (Convert.ToDecimal(strCurrentValue2) == 0)
                            strCurrentValue2 = "    $" + strCurrentValue2 + "         ";
                        else if (Convert.ToDecimal(strCurrentValue2) < 1000000)
                            strCurrentValue2 = "   $" + strCurrentValue2;
                        else if (Convert.ToDecimal(strCurrentValue2) < 10000000)
                            strCurrentValue2 = "   $" + strCurrentValue2;
                        else
                            strCurrentValue2 = "   $" + strCurrentValue2;

                        if (Convert.ToDecimal(strCurrentValue3) == 0)
                            strCurrentValue3 = "    $" + strCurrentValue3 + "         ";
                        else if (Convert.ToDecimal(strCurrentValue3) < 1000000)
                            strCurrentValue3 = "   $" + strCurrentValue3;
                        else if (Convert.ToDecimal(strCurrentValue3) < 10000000)
                            strCurrentValue3 = "    $" + strCurrentValue3;
                        else
                            strCurrentValue3 = "  $" + strCurrentValue3;

                        if (Convert.ToDecimal(strCurrentValue4) == 0)
                            strCurrentValue4 = "          $" + strCurrentValue4 + " ";
                        else if (Convert.ToDecimal(strCurrentValue4) < 1000000)
                            strCurrentValue4 = "       $" + strCurrentValue4;
                        else if (Convert.ToDecimal(strCurrentValue4) < 10000000)
                            strCurrentValue4 = "    $" + strCurrentValue4;
                        else
                            strCurrentValue4 = "$" + strCurrentValue4;

                        string Domestic_Equity = GetFilteredValue(table, UUID, ColName);

                        string Domequity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        //lochunk = new Chunk("\t\t" + Domestic_Equity +"\n"+ RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Domequity + "", setFontsAll(7, 0, 0));
                        Paragraph p = new Paragraph("" + Domestic_Equity + "\n       " + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n   $" + Domequity + "", setFontsAll(7, 0, 0));

                        Paragraph p1 = new Paragraph("", setFontsAll(6, 0, 0));
                        Chunk Para1 = new Chunk("\n     " + strSubAssetClassName1, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para2 = new Chunk("           " + strSubAssetClassName2, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para3 = new Chunk("           " + strSubAssetClassName3, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para4 = new Chunk("     " + strSubAssetClassName4, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));

                        Paragraph p2 = new Paragraph("", setFontsAll(6, 0, 0));
                        Chunk Para21 = new Chunk("     " + strCurrentAllocation1 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para22 = new Chunk("           " + strCurrentAllocation2 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para23 = new Chunk("            " + strCurrentAllocation3 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para24 = new Chunk("              " + strCurrentAllocation4 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));

                        Paragraph p3 = new Paragraph("", setFontsAll(6, 0, 0));
                        Chunk Para31 = new Chunk(" " + strCurrentValue1, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para32 = new Chunk("" + strCurrentValue2, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para33 = new Chunk("" + strCurrentValue3, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para34 = new Chunk("" + strCurrentValue4, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));

                        p1.Add(Para1);
                        p1.Add(Para2);
                        p1.Add(Para3);
                        p1.Add(Para4);

                        p2.Add(Para21);
                        p2.Add(Para22);
                        p2.Add(Para23);
                        p2.Add(Para24);

                        p3.Add(Para31);
                        p3.Add(Para32);
                        p3.Add(Para33);
                        p3.Add(Para34);

                        p.IndentationLeft = 62f;

                        loCell = new iTextSharp.text.Cell();
                        p.Leading = 8;
                        p2.Leading = 8;
                        p3.Leading = 8;
                        loCell.Colspan = 3;
                        j = j + 2;
                        loCell.Add(p);
                        loCell.Add(p1);
                        loCell.Add(p2);
                        loCell.Add(p3);
                        // loCell.Add(Para2);
                        // loCell.Add(Para3);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
                        SetBorder(loCell, true, true, true, true);

                        //loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 10) //International Equity Heading
                    {
                        string UUID = "B513B878-2071-E311-BDD3-0019B9E7EE05";
                        string International_Equity = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string Intquity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + International_Equity + GetSeprator(International_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Intquity + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        // loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 12)
                    {
                        string UUID = "6D4F7558-2071-E311-BDD3-0019B9E7EE05";
                        string Global_Opportunistic = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string GlobOpp = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Global_Opportunistic + GetSeprator(Global_Opportunistic) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + GlobOpp + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));

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
                        string UUID = "435F2391-2071-E311-BDD3-0019B9E7EE05";
                        string Liquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Liquid_Real_Assets + GetSeprator(Liquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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
                        string ConHoldUUID = "E298DD16-2071-E311-BDD3-0019B9E7EE05";
                        string PrivEqtyUUID = "3C2101F2-2071-E311-BDD3-0019B9E7EE05";
                        ConcentratedHolding = Convert.ToDouble(GetFilteredValue(table, ConHoldUUID, ColName, 1));
                        string Backcolor = "";
                        if (ConcentratedHolding == 0)
                        {
                            Backcolor = GetFilteredValue(table, PrivEqtyUUID, ColName, 16);
                            //string Private_Equity = Convert.ToString(table.Rows[10][0]);
                            string Private_Equity = GetFilteredValue(table, PrivEqtyUUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, PrivEqtyUUID, ColName, 1)));
                            lochunk = new Chunk("" + Private_Equity + GetSeprator(Private_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, PrivEqtyUUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }
                        else
                        {
                            Backcolor = GetFilteredValue(table, ConHoldUUID, ColName, 16);
                            //string Concentrated_Holdings = Convert.ToString(table.Rows[1][0]);
                            string Concentrated_Holdings = GetFilteredValue(table, ConHoldUUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, ConHoldUUID, ColName, 1)));
                            lochunk = new Chunk("" + Concentrated_Holdings + GetSeprator(Concentrated_Holdings) + RoundUp(Convert.ToString(GetFilteredValue(table, ConHoldUUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }

                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 10)
                    {
                        if (ConcentratedHolding != 0)
                        {
                            string UUID = "3C2101F2-2071-E311-BDD3-0019B9E7EE05";
                            string Private_Equity = GetFilteredValue(table, UUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                            lochunk = new Chunk(Private_Equity + GetSeprator(Private_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);
                            string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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
                        string UUID = "AF4FEE66-2071-E311-BDD3-0019B9E7EE05";
                        string Illiquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk(Illiquid_Real_Assets + GetSeprator(Illiquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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

            string StarategicPurposeChart = generateOverAllPieChartV2(newdataset.Tables[2], "1", TempFolderPath);
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(StarategicPurposeChart);

            string VolatileChart = generateOverAllPieChartV2(newdataset.Tables[4], "2", TempFolderPath);
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
            cell.Add(GetCenterTableV2(newdataset.Tables[3]));
            // Table.AddCell(cell);

            chunk = new Chunk("\n\n         Volatility Profile", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
            cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            // Table.AddCell(cell);

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
            dashjpg.SetAbsolutePosition(138f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(421f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(703f, dashImgYaxis);
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
            cell.Add(GetCenterTableV2(newdataset.Tables[3]));
            // cell.Rowspan = 2;
            Table.AddCell(cell);

            cell = new iTextSharp.text.Cell();
            cell.Add(volatilejpg);
            cell.Border = 0;
            Table.AddCell(cell);
            */

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


    public string generatePDFFinalV2(DataSet newdataset, string TempFolderPath)
    {
        liPageSize = 29;

        DB clsDB = new DB();

        String lsFooterTxt = String.Empty;

        DataTable table = newdataset.Tables[1];
        DataTable dtSubAsset = newdataset.Tables[0];

        Random rnd = new Random();
        string strRndNumber = Convert.ToString(rnd.Next(5));
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + strRndNumber;

        // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        int topmargin = 20;//25//30
        int bottommargin = 6;
        float dashImgYaxis = 304f;//299f
        if (Convert.ToBoolean(Convert.ToInt16(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == true && FooterText != "")
        {
            if (newdataset.Tables[3].Rows.Count >= 17)
            {
                topmargin = 20;
                bottommargin = 1;
                dashImgYaxis = 310f;
            }
        }

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 50, 50, topmargin, bottommargin);//10,10
                                                                                                                                                //  String fsFinalLocation = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));

        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

        //  AddFooter(document, FooterText);

        PdfPTable gTable = addFooterAbsoluteReturn(lsDateTime, 1, 1, liPageSize - 3, true, FooterText, "1", Footerlocation, ClientFooterTxt, Ssi_GreshamClientFooter);

        document.Open();
        PdfContentByte cb = writer.DirectContent;
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);

        //png.SetAbsolutePosition(40, 810);//540
        //png.ScalePercent(10);
        //document.Add(png);

        lsTotalNumberofColumns = 15 + "";

        iTextSharp.text.Table Table = new iTextSharp.text.Table(3, 1);
        iTextSharp.text.Cell cell = new Cell();

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(15, 7);   // 2 rows, 2 columns           
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

            string Cash_UUid = "A11B900A-2071-E311-BDD3-0019B9E7EE05";
            string Fixed_Income_UUid = "F41C2839-2071-E311-BDD3-0019B9E7EE05";
            string Hedged_Strategies_UUid = "D5A0EEA3-2071-E311-BDD3-0019B9E7EE05";
            string Domestic_Equity_UUid = "00000000-0000-0000-0000-000000000000";
            // string International_Equity_UUid = "B513B878-2071-E311-BDD3-0019B9E7EE05";
            string Global_Opportunistic_UUid = "6D4F7558-2071-E311-BDD3-0019B9E7EE05";
            string Liquid_Real_Assets_UUid = "435F2391-2071-E311-BDD3-0019B9E7EE05";
            /***********************************************************************************/
            string PrivEqty_UUID = "3C2101F2-2071-E311-BDD3-0019B9E7EE05";
            string ConHold_UUID = "E298DD16-2071-E311-BDD3-0019B9E7EE05";
            string Illiquid_Real_Assets_UUID = "AF4FEE66-2071-E311-BDD3-0019B9E7EE05";

            //string ColName = "sas_assetClassID";
            string ColName = "ssi_subassetclassid";

            //string MarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 1))));
            string MarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(table.Rows[0]["MarketableTotal"])));
            string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["Marketable"])));
            //string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 1))));
            string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(Convert.ToDecimal(table.Rows[0]["NonMarketableTotal"]))));
            string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["NonMarketable"])));

            double MarketPerc = Convert.ToDouble(MarketablePercent);
            double nonMarketPerc = Convert.ToDouble(NonMarketablePercent);
            if (MarketPerc > 100.0)
                MarketablePercent = "100.0";
            if (nonMarketPerc > 100.0)
                NonMarketablePercent = "100.0";

            cb.MoveTo(428, 435);
            cb.LineTo(603, 435);

            cb.MoveTo(470, 435);
            cb.LineTo(470, 398);

            cb.MoveTo(515, 435);
            cb.LineTo(515, 398);

            cb.MoveTo(560, 435);
            cb.LineTo(560, 398);

            ////   cb.MoveTo(430, 435);
            //// cb.LineTo(70, 300);
            ////Path closed and stroked
            cb.SetColorStroke(new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F")));
            cb.SetLineWidth(1.5f);
            cb.ClosePathStroke();

            //Rectangle rect = new Rectangle(100, 150, 220, 200);
            //String text = "test2\nggggg";
            //// try to get max font size that fit in rectangle
            //BaseFont bf = BaseFont.CreateFont();
            //int textHeightInGlyphSpace = bf.GetAscent(text) - bf.GetDescent(text);
            //float fontSize = 8f;
            //Phrase phrase = new Phrase("test1\n aaaabbbb", new Font(bf, fontSize));
            //ColumnText.ShowTextAligned(cb, Element.ALIGN_CENTER, phrase,
            //    // center horizontally
            //        (rect.Left + rect.Height) / 2,
            //    // shift baseline based on descent
            //       150,
            //        0);

            //// draw the rect
            //cb.SaveState();
            ////cb.setColorStroke(BaseColor.BLUE);
            //cb.Rectangle(rect.Left, rect.Bottom, rect.Width, rect.Height);
            //cb.Stroke();
            //cb.RestoreState();


            // cb.BeginText();
            // BaseFont f_cn = BaseFont.CreateFont("c:\\windows\\fonts\\calibri.ttf", BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            // cb.SetFontAndSize(f_cn, 6);
            // cb.SetTextMatrix(475, 15);  //(xPos, yPos)
            //// cb.ShowText("Some text here and the Date: " + DateTime.Now.ToShortDateString());
            // cb.ShowTextAligned(Element.ALIGN_CENTER, "Some text here and the Date: \r\n ssss " + DateTime.Now.ToShortDateString(), 475, 150, 0);
            // cb.EndText();




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

                        if (GreshamAdvisedFlag == "GA")
                        {
                            lochunk = new Chunk("\nGRESHAM ADVISED ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell.Add(lochunk);
                        }
                        else
                        {
                            lochunk = new Chunk("\nTOTAL INVESTMENT ASSETS" + Text + "", setFontsAll(10, 0, 0));
                            loCell.Add(lochunk);
                        }

                        string Title = "Portfolio Construction";

                        lochunk = new Chunk("\n" + Title, setFontsAllFrutiger(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        //loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        loCell.Add(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
                        loCell.Leading = 13F;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 1 && j == 0)
                    {
                        //the text of this line adjusted in above line due to space between header text.
                        //so this line is not adding in pdf
                        string Title = "";

                        lochunk = new Chunk("" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        //lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        //loCell.Add(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
                        loCell.Leading = 0F;
                        //loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 1) //Heading -- Left Boxes
                    {
                        string UUID = "A11B900A-2071-E311-BDD3-0019B9E7EE05";
                        string Hdr1 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        lochunk = new Chunk(Hdr1 + Text + "", setFontsAll(9, 1, 0));
                        //lochunk = new Chunk("HEAD", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                        loCell = new iTextSharp.text.Cell();

                        loCell.Colspan = 7;
                        j = j + 6;
                        loCell.Border = 0;
                        loCell.Add(lochunk);
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 7) //Dash Lines After Heading Left
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
                        string UUID = "00000000-0000-0000-0000-000000000000";
                        string Hdr2 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        if (Hdr2 == "0")
                            Hdr2 = "";
                        lochunk = new Chunk(Hdr2 + Text + "", setFontsAll(9, 1, 0));
                        //lochunk = new Chunk("HEAD", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
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
                        string UUID = "AF4FEE66-2071-E311-BDD3-0019B9E7EE05";
                        string Hdr3 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        lochunk = new Chunk(Hdr3 + Text + "", setFontsAll(9, 1, 0));
                        //lochunk = new Chunk("HEAD", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 0)
                    {
                        lochunk = new Chunk("Total Marketable \n Strategies\n" + MarketablePercent + Text + "%\n$" + MarketableValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 2)
                    {
                        string UUID = "A11B900A-2071-E311-BDD3-0019B9E7EE05";
                        string Cash = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Cash + GetSeprator(Cash) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                        //loCell.BorderWidthBottom = 2F;
                        //loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 4)
                    {
                        string UUID = "F41C2839-2071-E311-BDD3-0019B9E7EE05";
                        string Fixed_Income = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Fixed_Income + GetSeprator(Fixed_Income) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 6) //Hedged Strategies Heading
                    {
                        string UUID = "D5A0EEA3-2071-E311-BDD3-0019B9E7EE05";
                        string Hedged_Strategies = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string HedgedStr = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Hedged_Strategies + GetSeprator(Hedged_Strategies) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + HedgedStr + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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
                        string UUID = "00000000-0000-0000-0000-000000000000";
                        string Global_UUID = "2394047C-7D1D-E411-8A68-0019B9E7EE05";
                        string US_UUID = "77A63F24-2071-E311-BDD3-0019B9E7EE05";
                        string International_UUID = "B513B878-2071-E311-BDD3-0019B9E7EE05";
                        //string Emrging_UUID = "0C080B96-6A52-E511-940C-005056A0099E"; //fpr TEST
                        string Emrging_UUID = "F99F0D51-0382-E511-9418-005056A0567E"; //for PROD


                        string strSubAssetClassName1 = GetFilteredValue(dtSubAsset, Global_UUID, "ssi_subassetclassid", 2);
                        string strSubAssetClassName2 = GetFilteredValue(dtSubAsset, US_UUID, "ssi_subassetclassid", 2);
                        string strSubAssetClassName3 = GetFilteredValue(dtSubAsset, International_UUID, "ssi_subassetclassid", 2);
                        string strSubAssetClassName4 = GetFilteredValue(dtSubAsset, Emrging_UUID, "ssi_subassetclassid", 2);

                        string strCurrentAllocation1 = RoundUp(GetFilteredValue(dtSubAsset, Global_UUID, "ssi_subassetclassid", 7));
                        string strCurrentAllocation2 = RoundUp(GetFilteredValue(dtSubAsset, US_UUID, "ssi_subassetclassid", 7));
                        string strCurrentAllocation3 = RoundUp(GetFilteredValue(dtSubAsset, International_UUID, "ssi_subassetclassid", 7));
                        string strCurrentAllocation4 = RoundUp(GetFilteredValue(dtSubAsset, Emrging_UUID, "ssi_subassetclassid", 7));

                        //if (Convert.ToDecimal(strCurrentAllocation1) < 10)
                        //    strCurrentAllocation1 = strCurrentAllocation1 + "% ";
                        //else
                        //    strCurrentAllocation1 = strCurrentAllocation1 + "%";

                        //if (Convert.ToDecimal(strCurrentAllocation2) < 10)
                        //    strCurrentAllocation2 = strCurrentAllocation2 + "% ";
                        //else
                        //    strCurrentAllocation2 = strCurrentAllocation2 + "%";

                        //if (Convert.ToDecimal(strCurrentAllocation3) < 10)
                        //    strCurrentAllocation3 = strCurrentAllocation3 + "% ";
                        //else
                        //    strCurrentAllocation3 = strCurrentAllocation3 + "%";
                        //if (Convert.ToDecimal(strCurrentAllocation4) < 10)
                        //    strCurrentAllocation4 = strCurrentAllocation4 + "% ";
                        //else
                        //    strCurrentAllocation4 = strCurrentAllocation4 + "%";


                        string strCurrentValue1 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, Global_UUID, "ssi_subassetclassid", 3)));
                        string strCurrentValue2 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, US_UUID, "ssi_subassetclassid", 3)));
                        string strCurrentValue3 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, International_UUID, "ssi_subassetclassid", 3)));
                        string strCurrentValue4 = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(dtSubAsset, Emrging_UUID, "ssi_subassetclassid", 3)));

                        //if (Convert.ToDecimal(strCurrentValue1) == 0)
                        //    strCurrentValue1 = "     $" + strCurrentValue1 + "        ";
                        //else if (Convert.ToDecimal(strCurrentValue1) < 100000)
                        //    strCurrentValue1 = "   $" + strCurrentValue1 + "";
                        //else if (Convert.ToDecimal(strCurrentValue1) < 1000000)
                        //    strCurrentValue1 = "  $" + strCurrentValue1 + "";
                        //else if (Convert.ToDecimal(strCurrentValue1) < 10000000)
                        //    strCurrentValue1 = " $" + strCurrentValue1;
                        //else
                        //    strCurrentValue1 = " $" + strCurrentValue1;

                        //if (Convert.ToDecimal(strCurrentValue2) == 0)
                        //    strCurrentValue2 = "    $" + strCurrentValue2 + "         ";
                        //else if (Convert.ToDecimal(strCurrentValue2) < 1000000)
                        //    strCurrentValue2 = "   $" + strCurrentValue2;
                        //else if (Convert.ToDecimal(strCurrentValue2) < 10000000)
                        //    strCurrentValue2 = "   $" + strCurrentValue2;
                        //else
                        //    strCurrentValue2 = "   $" + strCurrentValue2;

                        //if (Convert.ToDecimal(strCurrentValue3) == 0)
                        //    strCurrentValue3 = "    $" + strCurrentValue3 + "         ";
                        //else if (Convert.ToDecimal(strCurrentValue3) < 1000000)
                        //    strCurrentValue3 = "   $" + strCurrentValue3;
                        //else if (Convert.ToDecimal(strCurrentValue3) < 10000000)
                        //    strCurrentValue3 = "    $" + strCurrentValue3;
                        //else
                        //    strCurrentValue3 = "  $" + strCurrentValue3;

                        //if (Convert.ToDecimal(strCurrentValue4) == 0)
                        //    strCurrentValue4 = "          $" + strCurrentValue4 + " ";
                        //else if (Convert.ToDecimal(strCurrentValue4) < 1000000)
                        //    strCurrentValue4 = "       $" + strCurrentValue4;
                        //else if (Convert.ToDecimal(strCurrentValue4) < 10000000)
                        //    strCurrentValue4 = "    $" + strCurrentValue4;
                        //else
                        //    strCurrentValue4 = "$" + strCurrentValue4;

                        string Domestic_Equity = GetFilteredValue(table, UUID, ColName);

                        string Domequity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        //lochunk = new Chunk("\t\t" + Domestic_Equity +"\n"+ RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Domequity + "", setFontsAll(7, 0, 0));
                        Paragraph p = new Paragraph("      " + Domestic_Equity + "\n       " + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n   $" + Domequity + "", setFontsAll(7, 0, 0));

                        Paragraph p1 = new Paragraph("", setFontsAll(6, 0, 0));
                        Chunk Para1 = new Chunk("\n     " + strSubAssetClassName1, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para2 = new Chunk("           " + strSubAssetClassName2, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para3 = new Chunk("           " + strSubAssetClassName3, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para4 = new Chunk("     " + strSubAssetClassName4, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));

                        Paragraph p2 = new Paragraph("", setFontsAll(6, 0, 0));
                        Chunk Para21 = new Chunk("     " + strCurrentAllocation1 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para22 = new Chunk("           " + strCurrentAllocation2 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para23 = new Chunk("            " + strCurrentAllocation3 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para24 = new Chunk("              " + strCurrentAllocation4 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));

                        Paragraph p3 = new Paragraph("", setFontsAll(6, 0, 0));
                        Chunk Para31 = new Chunk(" " + strCurrentValue1, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para32 = new Chunk("" + strCurrentValue2, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para33 = new Chunk("" + strCurrentValue3, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        Chunk Para34 = new Chunk("" + strCurrentValue4, setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));

                        p1.Add(Para1);
                        p1.Add(Para2);
                        p1.Add(Para3);
                        p1.Add(Para4);

                        p2.Add(Para21);
                        p2.Add(Para22);
                        p2.Add(Para23);
                        p2.Add(Para24);

                        p3.Add(Para31);
                        p3.Add(Para32);
                        p3.Add(Para33);
                        p3.Add(Para34);

                        //   p.IndentationLeft = 62f;

                        loCell = new iTextSharp.text.Cell();
                        p.Leading = 8;
                        p2.Leading = 8;
                        p3.Leading = 8;

                        ColumnText ct = new ColumnText(cb);
                        //ct.SetSimpleColumn(475, 60, 17,500, 5, Element.ALIGN_UNDEFINED);
                        //ct.AddElement(new Paragraph("This is the text added  in the \n rectangle"));
                        Phrase myText = new Phrase("" + strSubAssetClassName1 + "\n" + strCurrentAllocation1 + "%\n$" + strCurrentValue1 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        ct.SetSimpleColumn(myText, 880, 170, 15, 431, 8, Element.ALIGN_CENTER);
                        ct.Go();


                        ColumnText ct1 = new ColumnText(cb);
                        Phrase myText1 = new Phrase("" + strSubAssetClassName2 + "\n" + strCurrentAllocation2 + "%\n$" + strCurrentValue2 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        ct1.SetSimpleColumn(myText1, 960, 170, 26, 431, 8, Element.ALIGN_CENTER);
                        ct1.Go();

                        ColumnText ct2 = new ColumnText(cb);
                        Phrase myText2 = new Phrase("" + strSubAssetClassName3 + "\n" + strCurrentAllocation3 + "%\n$" + strCurrentValue3 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        ct2.SetSimpleColumn(myText2, 1042, 170, 33, 431, 8, Element.ALIGN_CENTER);
                        ct2.Go();

                        ColumnText ct3 = new ColumnText(cb);
                        Phrase myText3 = new Phrase("" + strSubAssetClassName4 + "\n" + strCurrentAllocation4 + "%\n$" + strCurrentValue4 + "", setFontsAllFrutiger(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#7F7F7F"))));
                        ct3.SetSimpleColumn(myText3, 1122, 170, 42, 431, 8, Element.ALIGN_CENTER);
                        ct3.Go();

                        loCell.Colspan = 3;
                        j = j + 2;
                        loCell.Add(p);
                        //loCell.Add(p1);
                        //loCell.Add(p2);
                        //loCell.Add(p3);
                        // loCell.Add(Para2);
                        // loCell.Add(Para3);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 10) //International Equity Heading
                    {
                        string UUID = "B513B878-2071-E311-BDD3-0019B9E7EE05";
                        string International_Equity = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string Intquity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + International_Equity + GetSeprator(International_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Intquity + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        // loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 12)
                    {
                        string UUID = "6D4F7558-2071-E311-BDD3-0019B9E7EE05";
                        string Global_Opportunistic = GetFilteredValue(table, UUID, ColName);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        string GlobOpp = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Global_Opportunistic + GetSeprator(Global_Opportunistic) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + GlobOpp + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));

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
                        string UUID = "435F2391-2071-E311-BDD3-0019B9E7EE05";
                        string Liquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Liquid_Real_Assets + GetSeprator(Liquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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
                        lochunk = new Chunk("Total Private\n Strategies\n" + NonMarketablePercent + Text + "%\n$" + NonMarketableValue + "", setFontsAll(7, 0, 0));
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
                        string ConHoldUUID = "E298DD16-2071-E311-BDD3-0019B9E7EE05";
                        string PrivEqtyUUID = "3C2101F2-2071-E311-BDD3-0019B9E7EE05";
                        ConcentratedHolding = Convert.ToDouble(GetFilteredValue(table, ConHoldUUID, ColName, 1));
                        string Backcolor = "";
                        if (ConcentratedHolding == 0)
                        {
                            Backcolor = GetFilteredValue(table, PrivEqtyUUID, ColName, 16);
                            //string Private_Equity = Convert.ToString(table.Rows[10][0]);
                            string Private_Equity = GetFilteredValue(table, PrivEqtyUUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, PrivEqtyUUID, ColName, 1)));
                            lochunk = new Chunk("" + Private_Equity + GetSeprator(Private_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, PrivEqtyUUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }
                        else
                        {
                            Backcolor = GetFilteredValue(table, ConHoldUUID, ColName, 16);
                            //string Concentrated_Holdings = Convert.ToString(table.Rows[1][0]);
                            string Concentrated_Holdings = GetFilteredValue(table, ConHoldUUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, ConHoldUUID, ColName, 1)));
                            lochunk = new Chunk("" + Concentrated_Holdings + GetSeprator(Concentrated_Holdings) + RoundUp(Convert.ToString(GetFilteredValue(table, ConHoldUUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        }

                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);

                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 10)
                    {
                        if (ConcentratedHolding != 0)
                        {
                            string UUID = "3C2101F2-2071-E311-BDD3-0019B9E7EE05";
                            string Private_Equity = GetFilteredValue(table, UUID, ColName);
                            string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                            lochunk = new Chunk(Private_Equity + GetSeprator(Private_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Add(lochunk);
                            string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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
                        string UUID = "AF4FEE66-2071-E311-BDD3-0019B9E7EE05";
                        string Illiquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk(Illiquid_Real_Assets + GetSeprator(Illiquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        string Backcolor = GetFilteredValue(table, UUID, ColName, 16);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Backcolor));
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
            #region Set with for center table  
            //if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) != true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) != true)
            //{
            //    setWidthsoftable(Table); // if Discrection columns are not to be shown
            //}
            if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["DiscretionaryFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) == true)
            {
                int[] headerwidths3 = { 30, 100, 30 }; // if aall columns are not to be shown
                Table.SetWidths(headerwidths3);
                Table.Width = 100;
            }
            else if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == false && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["DiscretionaryFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) == true)
            {
                int[] headerwidths3 = { 30, 70, 30 }; // if aall columns are not to be shown
                Table.SetWidths(headerwidths3);
                Table.Width = 100;
            }
            else if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["DiscretionaryFlg"])) == false && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) == true)
            {
                int[] headerwidths3 = { 30, 70, 30 }; // if aall columns are not to be shown
                Table.SetWidths(headerwidths3);
                Table.Width = 100;
            }
            else if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["DiscretionaryFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) == false)
            {
                int[] headerwidths3 = { 30, 70, 30 }; // if aall columns are not to be shown
                Table.SetWidths(headerwidths3);
                Table.Width = 100;
            }
            else if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == false && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["DiscretionaryFlg"])) == false && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) == true)
            {
                int[] headerwidths3 = { 30, 50, 30 }; // if aall columns are not to be shown
                Table.SetWidths(headerwidths3);
                Table.Width = 100;
            }
            else if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["DiscretionaryFlg"])) == false && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) == false)
            {
                int[] headerwidths3 = { 30, 50, 30 }; // if aall columns are not to be shown
                Table.SetWidths(headerwidths3);
                Table.Width = 100;
            }
            else if (Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["ShowTargetFlg"])) == false && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["DiscretionaryFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(newdataset.Tables[3].Rows[0]["InterimFlg"])) == false)
            {
                int[] headerwidths3 = { 30, 50, 30 }; // if aall columns are not to be shown
                Table.SetWidths(headerwidths3);
                Table.Width = 100;
            }
            else
            {
                setWidthsoftable(Table); // if Discrection columns are not to be shown
            }


            #endregion


            // setTableProperty(Table);
            // setWidthsoftable(Table);
            //fotable.Width = 100;
            Table.Alignment = 1;
            Table.Border = 0;
            Table.Cellspacing = 0;
            Table.Cellpadding = 2.8f;
            Table.Locked = false;

            string StarategicPurposeChart = generateOverAllPieChartV2(newdataset.Tables[2], "1", TempFolderPath);
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(StarategicPurposeChart);

            string VolatileChart = generateOverAllPieChartV2(newdataset.Tables[4], "2", TempFolderPath, true);
            iTextSharp.text.Image volatilejpg = iTextSharp.text.Image.GetInstance(VolatileChart);

            iTextSharp.text.Image solidjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\solid.png");
            iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\dashbar.png");
            //document.Add(jpg); //add an image to the created pdf document

            chunk = new Chunk("\n Strategic Purpose", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
            //cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
            Table.AddCell(cell);


            cell = new iTextSharp.text.Cell();
            cell.Add(GetCenterTableV2(newdataset.Tables[3]));
            Table.AddCell(cell);

            chunk = new Chunk("\n         Volatility Profile", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new iTextSharp.text.Cell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_TOP;
            cell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
            Table.AddCell(cell);

            jpg.Border = 0;
            jpg.ScaleToFit(225f, 205f);
            jpg.SetAbsolutePosition(15f, 50f);
            //jpg.Alignment = Image.TEXTWRAP | Image.ALIGN_RIGHT;
            //  jpg.IndentationLeft = 19f;
            jpg.SpacingAfter = 9f;
            //jpg.BorderWidthTop = 36f;
            //jpg.BorderColorTop = Color.WHITE;
            document.Add(jpg);

            volatilejpg.ScaleToFit(225f, 205f);
            volatilejpg.SetAbsolutePosition(665f, 50f);
            // jpg.Alignment = Image.TEXTWRAP | Image.ALIGN_RIGHT;
            volatilejpg.IndentationLeft = 19f;
            // volatilejpg.IndentationRight = 19f;
            volatilejpg.SpacingAfter = 9f;
            //volatilejpg.BorderWidthTop = 36f;
            // volatilejpg.BorderColorTop = Color.WHITE;
            document.Add(volatilejpg);

            solidjpg.ScaleToFit(50f, 190f);
            solidjpg.SetAbsolutePosition(138f, dashImgYaxis);
            solidjpg.IndentationLeft = 9f;
            solidjpg.SpacingAfter = 9f;
            document.Add(solidjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(421f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(703f, dashImgYaxis);
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
            cell.Add(GetCenterTableV2(newdataset.Tables[3]));
            // cell.Rowspan = 2;
            Table.AddCell(cell);

            cell = new iTextSharp.text.Cell();
            cell.Add(volatilejpg);
            cell.Border = 0;
            Table.AddCell(cell);
            */


            document.Add(loTable);
            document.Add(Table);
            document.Add(gTable);

            document.Close();

            //added 18-05-2018 (Sasmit- Cleanup JUNKFILES)
            try
            {
                if (StarategicPurposeChart != "")
                {
                    File.Delete(StarategicPurposeChart);
                }
                if (VolatileChart != "")
                {
                    File.Delete(VolatileChart);
                }
            }
            catch (Exception ex)
            {
            }
            clsCombinedReports clscombinedreports = new clsCombinedReports();
            clscombinedreports.SetTotalPageCount("Portfolio Construction Chart v2.1");


            //try
            //{
            //    FileInfo loFile = new FileInfo(ls);
            //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            //}
            //catch
            //{ }
            return fsFinalLocation.Replace(".xls", ".pdf");
        }
        else
        {
            //lblError.Text = "Record not found";
            return "Record not found";
        }
    }
    public PdfPTable addFooterAbsoluteReturn(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt, string strNew, String FooterLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        PdfPTable fotable = new PdfPTable(2);
        //fotable.Width = 97;
        //fotable.Border = 0;
        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        fotable.WidthPercentage = 100f;
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


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    // loChunk.Leading = 8f;
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
                    loChunk = new Paragraph(FooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
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


                    loCell = new PdfPCell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Paragraph(FooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.AddElement(loChunk);
                    fotable.AddCell(loCell);



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
                    loChunk = new Paragraph(ClientFooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 11f;
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
                    loChunk = new Paragraph(FooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
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
                    loChunk = new Paragraph(FooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loChunk.Leading = 8f;
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
                    loChunk = new Paragraph(FooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
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
                for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData; liCounter++)
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
                loChunk = new Paragraph(FooterTxt, setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
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
        loChunk = new Paragraph("                     ", Font8Whitecheck("test"));
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


    public string generatePDFFinalV2_TEMP(DataSet newdataset, string TempFolderPath)
    {
        liPageSize = 29;

        DB clsDB = new DB();

        String lsFooterTxt = String.Empty;

        DataTable table = newdataset.Tables[1];
        Random rnd = new Random();
        string strRndNumber = Convert.ToString(rnd.Next(5));
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss") + strRndNumber;

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        int topmargin = 20;//25//30
        int bottommargin = 6;
        float dashImgYaxis = 304f;//299f
        if (Convert.ToBoolean(newdataset.Tables[3].Rows[0]["ShowTargetFlg"]) == true && FooterText != "")
        {
            if (newdataset.Tables[3].Rows.Count >= 17)
            {
                topmargin = 14;
                bottommargin = 3;
                dashImgYaxis = 310f;
            }
        }

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 50, 50, topmargin, bottommargin);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
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

        PdfPTable Table = new PdfPTable(3);
        PdfPCell cell = new PdfPCell();

        PdfPTable loTable = new PdfPTable(15);   // 2 rows, 2 columns           
        PdfPCell loCell = new PdfPCell();
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
            //string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2)))));
            //string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Cash_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Fixed_Income_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Hedged_Strategies_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Domestic_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, International_Equity_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Global_Opportunistic_UUid, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Liquid_Real_Assets_UUid, ColName, 2))));
            string MarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["Marketable"])));
            string NonMarketableValue = String.Format("{0:#,###0;(#,###0)}", RoundValue(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 1))) + RoundValue(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 1))));
            //string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal((GetFilteredValue(table, ConHold_UUID, ColName, 2))) + RoundPercent(Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2)))));
            //string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(GetFilteredValue(table, PrivEqty_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, ConHold_UUID, ColName, 2)) + Convert.ToDecimal(GetFilteredValue(table, Illiquid_Real_Assets_UUID, ColName, 2))));
            string NonMarketablePercent = String.Format("{0:#,###0.0;(#,###0.0)}", RoundPercent(Convert.ToDecimal(table.Rows[0]["NonMarketable"])));

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
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);

                        if (GreshamAdvisedFlag == "GA")
                        {
                            lochunk = new Chunk("\nGRESHAM ADVISED ASSETS" + Text + "", setFontsAll(10, 0, 0));

                            loCell.AddElement(lochunk);
                        }
                        else
                        {
                            lochunk = new Chunk("\nTOTAL INVESTMENT ASSETS" + Text + "", setFontsAll(10, 0, 0));

                            loCell.AddElement(lochunk);
                        }

                        string Title = "Portfolio Construction";

                        lochunk = new Chunk("\n" + Title, setFontsAllFrutiger(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        //loCell = new PdfPCell();
                        loCell.AddElement(lochunk);

                        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        loCell.AddElement(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        loCell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
                        //loCell.Leading = 13F;//TEMP
                        loTable.AddCell(loCell);

                    }
                    else if (i == 1 && j == 0)
                    {
                        //the text of this line adjusted in above line due to space between header text.
                        //so this line is not adding in pdf
                        string Title = "";

                        lochunk = new Chunk("" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);

                        //lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
                        //loCell.AddElement(lochunk);

                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.Border = 0;
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loCell.VerticalAlignment = PdfPCell.ALIGN_TOP;
                        // loCell.Leading = 0F;//TEMP
                        //loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 1) //Heading Above boxes 
                    {
                        string UUID = "9776259D-0392-4DE0-8A12-0399724ABF8D";
                        string Hdr1 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        // lochunk = new Chunk(Hdr1 + Text + "", setFontsAll(9, 1, 0)); //Commented As per new logic in version 2.0
                        lochunk = new Chunk("   ", setFontsAll(9, 1, 0));
                        loCell = new PdfPCell();

                        loCell.Colspan = 7;
                        j = j + 6;
                        loCell.Border = 0;
                        loCell.AddElement(lochunk);
                        //loCell.Leading = 11F;////TEMP
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 7)
                    {
                        lochunk = new Chunk(UpperDashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        // loCell.Leading = 10f;//TEMP
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loCell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 8)
                    {
                        string UUID = "E2465B5C-40A7-4A35-B5EA-50A2C74CF6F5";
                        string Hdr2 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        //lochunk = new Chunk(Hdr2 + Text + "", setFontsAll(9, 1, 0));  //Commented As per new logic in version 2.0
                        lochunk = new Chunk(" .", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"))));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Colspan = 5;
                        j = j + 4;
                        loCell.Border = 0;
                        //loCell.Leading = 11F;//TEMP
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 13)
                    {
                        lochunk = new Chunk(UpperDashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        //loCell.Leading = 10f;//TEMP
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loCell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 2 && j == 14)
                    {
                        string UUID = "C2A2D71C-D704-DE11-A38C-001D09665E8F";
                        string Hdr3 = Convert.ToString(GetFilteredValue(table, UUID, ColName, 3));
                        //lochunk = new Chunk(Hdr3 + Text + "", setFontsAll(9, 1, 0));    //Commented As per new logic in version 2.0
                        lochunk = new Chunk("  ", setFontsAll(9, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        // loCell.Leading = 11F;//TEMP
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 0)
                    {
                        lochunk = new Chunk("Marketable Strategies\n\n" + MarketablePercent + Text + "%\n$" + MarketableValue + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 3 && j == 2)
                    {
                        string UUID = "9776259D-0392-4DE0-8A12-0399724ABF8D";
                        string Cash = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Cash + GetSeprator(Cash) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;

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
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 6) //Hedged Strategies Heading
                    {
                        string UUID = "2287692A-D704-DE11-A38C-001D09665E8F";
                        string Hedged_Strategies = GetFilteredValue(table, UUID, ColName);
                        string HedgedStr = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Hedged_Strategies + GetSeprator(Hedged_Strategies) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + HedgedStr + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B6DDE8"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 7)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        // loCell.Leading = 10f;//TEMP
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loCell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 8)
                    {

                        loCell = new PdfPCell();

                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        SetBorder(loCell, true, true, true, true);
                        loCell.Colspan = 3;
                        j = j + 2;
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;

                        PdfPTable loinnertable = new PdfPTable(2);

                        PdfPCell locelltop1 = new PdfPCell();
                        PdfPCell locellbottom1 = new PdfPCell();
                        PdfPCell locellbottom2 = new PdfPCell();
                        PdfPCell locellbottom3 = new PdfPCell();
                        PdfPCell locellbottom4 = new PdfPCell();

                        int[] widths = { 50, 50 };
                        loinnertable.SetWidths(widths);
                        loinnertable.WidthPercentage = 100;


                        string UUID = "E2A78BEB-D604-DE11-A38C-001D09665E8F";
                        string Domestic_Equity = GetFilteredValue(table, UUID, ColName);
                        string Domequity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Domestic_Equity + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Domequity + "", setFontsAll(7, 0, 0));
                        locelltop1.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        // locelltop1.Colspan = 3;
                        locelltop1.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        locelltop1.AddElement(lochunk);

                        lochunk = new Chunk("AAA");
                        locellbottom1.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        locellbottom1.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        locellbottom1.AddElement(lochunk);

                        lochunk = new Chunk("BBB");
                        locellbottom2.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        locellbottom2.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        locellbottom2.AddElement(lochunk);

                        lochunk = new Chunk("CCC");
                        locellbottom3.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        locellbottom3.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        locellbottom3.AddElement(lochunk);

                        lochunk = new Chunk("DDD");
                        locellbottom4.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                        locellbottom4.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        locellbottom4.AddElement(lochunk);


                        loinnertable.AddCell(locelltop1);
                        loinnertable.AddCell(locellbottom1);
                        //   loinnertable.AddCell(locellbottom2);
                        //   loinnertable.AddCell(locellbottom3);
                        //   loinnertable.AddCell(locellbottom4);

                        //     loCell.AddElement(locelltop1);
                        loTable.AddCell(locelltop1);

                    }
                    else if (i == 3 && j == 10) //International Equity Heading
                    {
                        string UUID = "42B39247-D704-DE11-A38C-001D09665E8F";
                        string International_Equity = GetFilteredValue(table, UUID, ColName);
                        string Intquity = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + International_Equity + GetSeprator(International_Equity) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + Intquity + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        // loTable.AddCell(loCell); //Commented As per new logic in version 2

                    }
                    else if (i == 3 && j == 12)
                    {
                        string UUID = "8413896B-4925-DF11-B686-001D09665E8F";
                        string Global_Opportunistic = GetFilteredValue(table, UUID, ColName);
                        string GlobOpp = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Global_Opportunistic + GetSeprator(Global_Opportunistic) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + GlobOpp + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        //SetBorder(loCell, false);
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 13)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        //  loCell.Leading = 10f;//TEMP
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loCell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 3 && j == 14)
                    {
                        string UUID = "0332530A-1AD3-DF11-9789-0019B9E7EE05";
                        string Liquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk("" + Liquid_Real_Assets + GetSeprator(Liquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 4 && j == 0)
                    {
                        lochunk = new Chunk("");
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        loCell.Colspan = 15;
                        j = j + 14;
                        loCell.BorderWidthBottom = 1F;
                        loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        //loCell.Height = 5;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 0)
                    {
                        lochunk = new Chunk("Non-Marketable Strategies\n" + NonMarketablePercent + Text + "%\n$" + NonMarketableValue + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        SetBorder(loCell, true, true, true, true);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.Color.White);
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }
                    else if (i == 6 && j == 7)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(9, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        //   loCell.Leading = 10f;//TEMP
                        loCell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
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

                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));

                        SetBorder(loCell, true, true, true, true);
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
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
                            loCell = new PdfPCell();
                            loCell.AddElement(lochunk);
                            loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#C2D69A"));
                            SetBorder(loCell, true, true, true, true);

                            loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                            loTable.AddCell(loCell);
                        }
                        else
                        {
                            lochunk = new Chunk("", setFontsAll(7, 0, 0));
                            loCell = new PdfPCell();
                            loCell.AddElement(lochunk);
                            loCell.Border = 0;
                            loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                            loTable.AddCell(loCell);
                        }

                    }
                    else if (i == 6 && j == 13)
                    {
                        lochunk = new Chunk(DashLines + Text + "", setFontsAll(7, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        //  loCell.Leading = 10f; //TEMP
                        loCell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else if (i == 6 && j == 14)
                    {
                        string UUID = "C2A2D71C-D704-DE11-A38C-001D09665E8F";
                        string Illiquid_Real_Assets = GetFilteredValue(table, UUID, ColName);
                        string strValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(GetFilteredValue(table, UUID, ColName, 1)));
                        lochunk = new Chunk(Illiquid_Real_Assets + GetSeprator(Illiquid_Real_Assets) + RoundUp(Convert.ToString(GetFilteredValue(table, UUID, ColName, 2))) + Text + "%\n$" + strValue + "", setFontsAll(7, 0, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#CCC0DA"));
                        SetBorder(loCell, true, true, true, true);

                        loCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        loTable.AddCell(loCell);

                    }
                    else
                    {
                        lochunk = new Chunk(Text, setFontsAll(10, 1, 0));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);

                        loCell.Border = 0;
                        loCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        loTable.AddCell(loCell);
                    }
                }
            }
            #endregion

            lsTotalNumberofColumns = 3 + "";
            setTableProperty(Table);

            string StarategicPurposeChart = generateOverAllPieChart(newdataset.Tables[2], "1", TempFolderPath);
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(StarategicPurposeChart);

            string VolatileChart = generateOverAllPieChart(newdataset.Tables[4], "2", TempFolderPath);
            iTextSharp.text.Image volatilejpg = iTextSharp.text.Image.GetInstance(VolatileChart);

            iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\dashbar.png");
            //document.Add(jpg); //add an image to the created pdf document

            chunk = new Chunk("\n\nStrategic Purpose", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new PdfPCell();
            cell.AddElement(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            Table.AddCell(cell);


            cell = new PdfPCell();
            cell.AddElement(GetCenterTable1(newdataset.Tables[3]));
            Table.AddCell(cell);

            chunk = new Chunk("\n\n         Volatility Profile", setFontsAll(9, 1, 0));
            //chunk = new Chunk("", setFontsAll(7, 1, 0));
            cell = new PdfPCell();
            cell.AddElement(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
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
            dashjpg.SetAbsolutePosition(138f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(421f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            dashjpg.ScaleToFit(50f, 190f);
            dashjpg.SetAbsolutePosition(703f, dashImgYaxis);
            dashjpg.IndentationLeft = 9f;
            dashjpg.SpacingAfter = 9f;
            document.Add(dashjpg);

            /*
            chunk = new Chunk("Strategic Purpose \n\n\n", setFontsAll(7, 1, 0));
            cell = new PdfPCell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
            Table.AddCell(cell);

            chunk = new Chunk("");
            cell = new PdfPCell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
            Table.AddCell(cell);


            chunk = new Chunk("Volatile Profile", setFontsAll(7, 1, 0));
            cell = new PdfPCell();
            cell.Add(chunk);
            cell.Border = 0;
            cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
            Table.AddCell(cell);

            chunk = new Chunk("");
            cell = new PdfPCell();
            cell.Add(jpg);
            cell.Border = 0;
            cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
            Table.AddCell(cell);

            cell = new PdfPCell();
            cell.Add(GetCenterTable(newdataset.Tables[3]));
            // cell.Rowspan = 2;
            Table.AddCell(cell);

            cell = new PdfPCell();
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

    public iTextSharp.text.Table GetCenterTableV2(DataTable dt)
    {
        int rowsize = dt.Rows.Count + 1;

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(8, rowsize);//18   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        iTextSharp.text.Chunk lochunk = new Chunk();

        //  lsTotalNumberofColumns = 5 + "";


        if (Convert.ToBoolean(Convert.ToInt32(dt.Rows[0]["ShowTargetFlg"])) == true && Convert.ToBoolean(Convert.ToInt32(dt.Rows[0]["InterimFlg"])) != true)
        {
            lsTotalNumberofColumns = 5 + "";
            AllocationColumnSize = 27;// to maintain the width of first column.
            int[] headerwidths5 = { AllocationColumnSize, 12, 12, 12, 12, 12, 12, 12 };
            loTable.SetWidths(headerwidths5);
            //  setTableProperty(loTable);
        }
        else
        {
            AllocationColumnSize = 27;// to maintain the width of first column.
            int[] headerwidths5 = { AllocationColumnSize, 12, 12, 12, 12, 12, 12, 12 };
            loTable.SetWidths(headerwidths5);

        }
        // setTableProperty(loTable);
        loTable.DefaultVerticalAlignment = Element.ALIGN_TOP;

        loTable.Alignment = 1;
        loTable.Border = 1;
        loTable.Cellspacing = 0;
        loTable.Cellpadding = 3f;
        loTable.Locked = false;

        loTable.Width = 100;


        for (int i = 0; i < rowsize; i++)
        {
            for (int j = 0; j < 8; j++)
            {
                if (i == 0 && j == 0)
                {
                    lochunk = new Chunk("", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 7F;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 1)
                {
                    lochunk = new Chunk("Current Allocation", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 7F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;

                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 2)
                {
                    lochunk = new Chunk("Current Allocation", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 7F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 3)
                {
                    lochunk = new Chunk("Interim Allocation", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 7F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 4)
                {
                    lochunk = new Chunk("Target Allocation", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 7F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 5)
                {
                    lochunk = new Chunk("Discretionary \n Range", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Colspan = 2;
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 7F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 7)
                {
                    lochunk = new Chunk("Volatility Profile", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 7F;
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
                    {
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                        if (ColValue != "")
                            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(ColValue));
                    }
                    else if (j == 5)
                    {
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                        if (ColValue != "")
                            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(ColValue));
                    }
                    else if (j == 6)
                    {
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                        if (ColValue != "")
                            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}%", Convert.ToDecimal(ColValue));
                    }
                    else if (j == 7)
                    {
                        ColValue = Convert.ToString(dt.Rows[i - 1][j]);
                        // ColValue = ColValue + "      ";
                    }




                    if (Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)
                        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    else if (Convert.ToString(dt.Rows[i - 1]["TabFlg"]) == "True")
                    {
                        if (j == 0)
                            lochunk = new Chunk("       " + ColValue, setFontsAllFrutiger(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#808080"))));
                        else
                            lochunk = new Chunk(ColValue, setFontsAllFrutiger(7, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#808080"))));
                    }
                    else
                        lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));

                    loCell = new iTextSharp.text.Cell();
                    loCell.Add(lochunk);

                    if (dt.Rows.Count >= 17)
                        loCell.Leading = 2.6F;
                    else
                        loCell.Leading = 3F;

                    if (Convert.ToString(dt.Rows[i - 1]["TabFlg"]) == "True")
                        loCell.Leading = 2F;

                    if (ColUUID.ToLower() == "fffb5207-6075-e211-aa29-0019b9e7ee05" && Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)//Risk Reduction Strategies
                    {
                        // string Color = Convert.ToString(dt.Rows[i - 1]["ColourCode"]);
                        string Color = (Convert.ToString(dt.Rows[i - 1]["ColourCode"]) == "" ? "#B6DDE8" : Convert.ToString(dt.Rows[i - 1]["ColourCode"]));
                        loCell.Leading = 2.7F;
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Color));
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_MIDDLE;

                    }
                    else if (ColUUID.ToLower() == "277c510e-6075-e211-aa29-0019b9e7ee05" && Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)//Growth Strategies
                    {
                        loCell.Leading = 2.7F;
                        string Color = (Convert.ToString(dt.Rows[i - 1]["ColourCode"]) == "" ? "#C2D69A" : Convert.ToString(dt.Rows[i - 1]["ColourCode"]));
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Color));
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_MIDDLE;

                    }
                    else if (ColUUID.ToLower() == "07f0821c-6075-e211-aa29-0019b9e7ee05" && Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)//Economic Hedges
                    {
                        loCell.Leading = 2.7F;
                        string Color = (Convert.ToString(dt.Rows[i - 1]["ColourCode"]) == "" ? "#CCC0DA" : Convert.ToString(dt.Rows[i - 1]["ColourCode"]));
                        loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(Color));
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_MIDDLE;
                    }
                    loCell.Border = 0;
                    if (i == rowsize - 1)
                    {
                        loCell.BorderWidthTop = 0.5F;
                        loCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.Leading = 2.5f;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_MIDDLE;
                    }
                    if (j == 0)
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    else if (j == 7)
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;

                    if ((j == 1 || j == 2 || j == 3 || j == 4 || j == 5 || j == 6) && Convert.ToString(dt.Rows[i - 1]["TotalFlg"]) == "True")
                    {
                        loCell.BorderWidthTop = 0.5F;
                        loCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
                        loCell.Leading = 2.5f;
                        loCell.VerticalAlignment = iTextSharp.text.Cell.ALIGN_MIDDLE;
                    }
                    loTable.AddCell(loCell);
                }
            }
        }
        if (Convert.ToBoolean(Convert.ToInt32(dt.Rows[0]["DiscretionaryFlg"])) != true)
        {
            loTable.DeleteColumn(6); // Delete Discretionary Min and Discretionary Max Columns from Table
            loTable.DeleteColumn(5);
        }

        if (Convert.ToBoolean(Convert.ToInt32(dt.Rows[0]["ShowTargetFlg"])) != true)
            loTable.DeleteColumn(4); // Delete Target Allocation from the Table

        if (Convert.ToBoolean(Convert.ToInt32(dt.Rows[0]["InterimFlg"])) != true)
            loTable.DeleteColumn(3); // Delete the Interim Allocation from the Table






        return loTable;
    }

    public iTextSharp.text.Table GetCenterTable(DataTable dt)
    {
        int rowsize = dt.Rows.Count + 1;

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(5, rowsize);//18   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        iTextSharp.text.Chunk lochunk = new Chunk();

        lsTotalNumberofColumns = 5 + "";

        if (Convert.ToBoolean(dt.Rows[0]["ShowTargetFlg"]) == true)
            AllocationColumnSize = 27;// to maintain the width of first column.

        setTableProperty(loTable);

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
                    loCell.Leading = 9F;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 1)
                {
                    lochunk = new Chunk("Current Allocation", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 9F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 2)
                {
                    lochunk = new Chunk("Current Allocation (%)", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 9F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 3)
                {
                    lochunk = new Chunk("Target Allocation (%)", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 9F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 4)
                {
                    lochunk = new Chunk("Volatility Profile", setFontsAllFrutiger(7, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loCell.Leading = 9F;
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

        if (Convert.ToBoolean(dt.Rows[0]["ShowTargetFlg"]) != true)
            loTable.DeleteColumn(3);

        return loTable;
    }

    public PdfPTable GetCenterTable1(DataTable dt)
    {
        int rowsize = dt.Rows.Count + 1;

        PdfPTable loTable = new PdfPTable(5);//18   // 2 rows, 2 columns           
        PdfPCell loCell = new PdfPCell();
        iTextSharp.text.Chunk lochunk = new Chunk();

        lsTotalNumberofColumns = 5 + "";

        if (Convert.ToBoolean(dt.Rows[0]["ShowTargetFlg"]) == true)
            AllocationColumnSize = 27;// to maintain the width of first column.

        setTableProperty(loTable);

        for (int i = 0; i < rowsize; i++)
        {
            for (int j = 0; j < 5; j++)
            {
                if (i == 0 && j == 0)
                {
                    lochunk = new Chunk("", setFontsAllFrutiger(7, 1, 0));
                    loCell = new PdfPCell();
                    loCell.AddElement(lochunk);
                    loCell.Border = 0;
                    //loCell.Leading = 9F;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 1)
                {
                    lochunk = new Chunk("Current Allocation", setFontsAllFrutiger(7, 1, 0));
                    loCell = new PdfPCell();
                    loCell.AddElement(lochunk);
                    loCell.Border = 0;
                    //loCell.Leading = 9F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 2)
                {
                    lochunk = new Chunk("Current Allocation (%)", setFontsAllFrutiger(7, 1, 0));
                    loCell = new PdfPCell();
                    loCell.AddElement(lochunk);
                    loCell.Border = 0;
                    //loCell.Leading = 9F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 3)
                {
                    lochunk = new Chunk("Target Allocation (%)", setFontsAllFrutiger(7, 1, 0));
                    loCell = new PdfPCell();
                    loCell.AddElement(lochunk);
                    loCell.Border = 0;
                    //loCell.Leading = 9F;
                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                    loTable.AddCell(loCell);
                }
                if (i == 0 && j == 4)
                {
                    lochunk = new Chunk("Volatility Profile", setFontsAllFrutiger(7, 1, 0));
                    loCell = new PdfPCell();
                    loCell.AddElement(lochunk);
                    loCell.Border = 0;
                    //loCell.Leading = 9F;
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

                    loCell = new PdfPCell();
                    loCell.AddElement(lochunk);
                    //loCell.Leading = 4F;

                    if (ColUUID.ToLower() == "fffb5207-6075-e211-aa29-0019b9e7ee05" && Convert.ToString(dt.Rows[i - 1]["StrategicFlg"]) == "True" && j == 0)//Risk Reduction Strategies
                    {
                        //  loCell.Leading = 4F;
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

        /**  if (Convert.ToBoolean(dt.Rows[0]["ShowTargetFlg"]) != true)
              loTable.DeleteColumn(3);**/
        //TEMP

        return loTable;
    }

    public string generateOverAllPieChartV2(DataTable table, string ChartType, string TempFolderPath, bool Volatile = false)
    {
        /*  Chart Type
         * Strategic Purpose == 1;
         * Volatile Profile == 2; */

        string Header = string.Empty;
        if (ChartType == "1")
            Header = "";//"Strategic Purpose"
        else
            Header = "";//Volatility Profile
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss");

        //  String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";
        String fsFinalLocation = TempFolderPath + "\\" + "OP_" + Guid.NewGuid().ToString() + ".xls";

        AppDomain.CurrentDomain.Load("JCommon");
        org.jfree.data.general.DefaultPieDataset myDataSet = new org.jfree.data.general.DefaultPieDataset();
        if (table.Rows.Count > 0)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (Convert.ToString(table.Rows[i][2]) != "")
                {
                    string ColValue = String.Format("{0:#,###0.0;-#,###0.0}", Convert.ToDecimal(table.Rows[i][2]));
                    //myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[i][2]), 1)));
                    //if (ChartType == "1")
                    //    myDataSet.setValue(Convert.ToString(table.Rows[i][3]), Convert.ToDouble(ColValue));
                    //else

                    string KeyName = Convert.ToString(table.Rows[i][0]);
                    int Digits = 0;
                    int index = ColValue.IndexOf('.');
                    if (index > 0)
                    {
                        Digits = ColValue.Substring(0, index).Length;
                    }


                    if (KeyName.ToLower().Contains("cash"))
                    {
                        if (Digits == 2)
                        {
                            KeyName = KeyName + "            ";
                        }
                        else
                        {
                            KeyName = KeyName + "              ";
                        }
                    }

                    if (KeyName.ToLower().Contains("high"))
                    {
                        if (Digits == 2)
                        {
                            KeyName = KeyName + "             ";
                        }
                        else
                        {
                            KeyName = KeyName + "              ";

                        }
                    }
                    if (KeyName.ToLower().Contains("low"))
                    {
                        if (Digits == 2) // 
                        {
                            KeyName = KeyName + "              ";
                        }
                        else
                        {
                            KeyName = KeyName + "                ";
                        }
                    }
                    if (KeyName.ToLower().Contains("moderate"))
                    {
                        if (Digits == 2) // 
                        {
                            KeyName = KeyName + "    ";
                        }
                        else
                        {
                            KeyName = KeyName + "       ";
                        }
                    }
                    if (KeyName.ToLower().Contains("risk"))
                    {
                        if (Digits == 2) // 
                        {
                            KeyName = KeyName + "  ";
                        }
                        else { KeyName = KeyName + "   "; }
                    }
                    if (KeyName.ToLower().Contains("growth"))
                    {
                        if (Digits == 2) // 
                        {
                            KeyName = KeyName + "                 ";
                        }
                        else
                        {
                            KeyName = KeyName + "                 ";
                        }
                    }
                    if (KeyName.ToLower().Contains("economic"))
                    {
                        if (Digits == 2) // 
                        {
                            KeyName = KeyName + "                ";
                        }
                        else
                        {
                            KeyName = KeyName + "                 ";
                        }
                    }


                    int KeyNameCount = KeyName.Length;
                    string Space = "";
                    //if (Volatile)
                    //{
                    //    for (int k = KeyNameCount; k < 18; k++)
                    //    {

                    //        Space = Space + " ";

                    //    }
                    //    if (Space != "")
                    //    {
                    //        KeyName = KeyName + Space;
                    //    }
                    //}
                    //else
                    //{
                    //    for (int k = KeyNameCount; k < 30; k++)
                    //    {

                    //        Space = Space + " ";

                    //    }
                    //    if (Space != "")
                    //    {
                    //        KeyName = KeyName + Space;
                    //    }
                    //}
                    int FinalCount = KeyName.Length;

                    //if (ColValue == "0.0")
                    //    myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble("0.0001"));
                    //else
                    //    myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(ColValue));

                    if (ColValue == "0.0")
                        myDataSet.setValue(KeyName, Convert.ToDouble("0.0001"));
                    else
                        myDataSet.setValue(KeyName, Convert.ToDouble(ColValue));
                    //if (ColValue != "0.0")
                    //    myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(ColValue));
                }
            }
        }
        JFreeChart pieChart = ChartFactory.createPieChart("", myDataSet, true, true, false);
        pieChart.setPadding(new RectangleInsets(0, 0, 30, 0));

        pieChart.setBackgroundPaint(java.awt.Color.white);
        pieChart.setBorderVisible(false);

        pieChart.setTitle(new org.jfree.chart.title.TextTitle(Header, new java.awt.Font("Frutiger55", java.awt.Font.BOLD, 14)));

        PiePlot ColorConfigurator = (PiePlot)pieChart.getPlot();

        ColorConfigurator.setLabelGenerator(null);

        java.text.NumberFormat df = java.text.NumberFormat.getNumberInstance();
        df.setMinimumFractionDigits(1);
        ColorConfigurator.setLegendLabelGenerator(new StandardPieSectionLabelGenerator("{0} {1}%", df, df));
        // ColorConfigurator.
        LegendTitle legend = pieChart.getLegend(0);
        //legend.setPadding(new RectangleInsets(0, 50, 0, 50));
        legend.setItemFont(new System.Drawing.Font("Frutiger-Roman", 12));
        if (Volatile)
        {
            legend.setItemLabelPadding(new RectangleInsets(0, 0, 0, 0));
        }
         // legend.setPosition(RectangleEdge.BOTTOM);
         /* legend.setBackgroundPaint(java.awt.Color.white)*/
         ;
        legend.setBorder(BlockBorder.NONE);


        if (Volatile)
        {
            //  legend.setItemLabelPadding(new RectangleInsets(0, 20, 0, 0));
            legend.setPadding(0, 80, 0, 25);
        }
        else
        {
            legend.setPadding(0, 35, 0, 0);
        }

        //legend.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // legend.setVerticalAlignment(VerticalAlignment.CENTER);
        //  legend.setBorder(0, 0, 0, 0);

        ColorConfigurator.setLabelBackgroundPaint(System.Drawing.Color.White);// ColorConfigurator.getLabelPaint()
        ColorConfigurator.setLabelOutlinePaint(System.Drawing.Color.White);
        ColorConfigurator.setLabelShadowPaint(System.Drawing.Color.White);
        //ColorConfigurator.setBackgroundPaint(System.Drawing.Color.White);
        ColorConfigurator.setOutlinePaint(System.Drawing.Color.White);
        //ColorConfigurator.setOutlineStroke setOutlinePaint(System.Drawing.Brush.);
        //  ColorConfigurator.setLabelFont(new System.Drawing.Font("Frutiger-Roman", 14));//, System.Drawing.FontStyle.Bold
        ColorConfigurator.setLabelGenerator(null);
        ColorConfigurator.setCircular(true);
        //ColorConfigurator.setOutlineVisible(false);  fix to remove border
        //ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardPieSectionLabelGenerator("{0} {1}%"));
        //ColorConfigurator.setLabelGenerator(null);
        //ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardCategoryItemLabelGenerator("{0}"));
        //ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardCategoryItemLabelGenerator("{0} =  {1}%"));
        //ColorConfigurator.setInteriorGap(0.30);

        //ColorConfigurator.setInteriorGap(0.15);
        //  ColorConfigurator.setInteriorGap(0.15);
        // ColorConfigurator.setLabelGap(0);

        java.util.List keys = myDataSet.getKeys();

        for (int i = 0; i < keys.size(); i++)
        {
            string RiskColor = "#CCECFF";
            string GrowthColor = "#E6D5F3";
            string EcoColor = "#C5E0B4";
            string CashColor = "#D9D9D9";
            string HighColor = "#C4BD97";
            string LowColor = "#E6B9B8";
            string ModerateColor = "#FAC090";

            if (keys.get(i).ToString().Contains("Risk") || keys.get(i).ToString().Contains("Growth") || keys.get(i).ToString().Contains("Economic"))
            {
                RiskColor = GetFilteredValue(table, "FFFB5207-6075-E211-AA29-0019B9E7EE05", "StrategicUID", 5);
                GrowthColor = GetFilteredValue(table, "277C510E-6075-E211-AA29-0019B9E7EE05", "StrategicUID", 5);
                EcoColor = GetFilteredValue(table, "07F0821C-6075-E211-AA29-0019B9E7EE05", "StrategicUID", 5);
            }

            if (keys.get(i).ToString().Contains("Cash") || keys.get(i).ToString().Contains("High") || keys.get(i).ToString().Contains("Low") || keys.get(i).ToString().Contains("Moderate"))
            {
                //TEST
                //CashColor = GetFilteredValue(table, "BDC4D744-9379-E511-9414-005056A0099E", "ssi_volitalityprofileId", 3);
                //                HighColor = GetFilteredValue(table, "195F7271-9379-E511-9414-005056A0099E", "ssi_volitalityprofileId", 3);
                //              LowColor = GetFilteredValue(table, "D533A953-9379-E511-9414-005056A0099E", "ssi_volitalityprofileId", 3);
                //            ModerateColor = GetFilteredValue(table, "78D30464-9379-E511-9414-005056A0099E", "ssi_volitalityprofileId", 3);

                //PROD
                CashColor = GetFilteredValue(table, "228B5C51-0082-E511-9418-005056A0567E", "ssi_volitalityprofileId", 3);
                HighColor = GetFilteredValue(table, "D60E645F-0082-E511-9418-005056A0567E", "ssi_volitalityprofileId", 3);
                LowColor = GetFilteredValue(table, "5BD9DC6A-0082-E511-9418-005056A0567E", "ssi_volitalityprofileId", 3);
                ModerateColor = GetFilteredValue(table, "6FF68E77-0082-E511-9418-005056A0567E", "ssi_volitalityprofileId", 3);
            }

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
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml(RiskColor));
            if (keys.get(i).ToString().Contains("Growth"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml(GrowthColor));
            if (keys.get(i).ToString().Contains("Economic"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml(EcoColor));
            if (keys.get(i).ToString().Contains("Cash"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml(CashColor));
            if (keys.get(i).ToString().Contains("High"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml(HighColor));
            if (keys.get(i).ToString().Contains("Low"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml(LowColor));
            if (keys.get(i).ToString().Contains("Moderate"))
                ColorConfigurator.setSectionPaint(i, System.Drawing.ColorTranslator.FromHtml(ModerateColor));
        }
        ChartRenderingInfo thisImageMapInfo = new ChartRenderingInfo();
        java.io.OutputStream jos = new java.io.FileOutputStream(fsFinalLocation.Replace(".xls", ".png"));
        //ChartUtilities.writeChartAsPNG(jos, pieChart, 200, 330);
        //if (Volatile)
        //{
        //    ChartUtilities.writeChartAsPNG(jos, pieChart, 150, 300);
        //}
        //else
        //{
        //    ChartUtilities.writeChartAsPNG(jos, pieChart, 205, 300);
        //}
        //  ChartUtilities.writeChartAsPNG(jos, pieChart, 225, 330);
        ChartUtilities.writeChartAsPNG(jos, pieChart, 280, 330);

        return fsFinalLocation.Replace(".xls", ".png");
    }

    public string generateOverAllPieChart(DataTable table, string ChartType, string TempFolderPath)
    {
        /*  Chart Type
         * Strategic Purpose == 1;
         * Volatile Profile == 2; */

        string Header = string.Empty;
        if (ChartType == "1")
            Header = "";//"Strategic Purpose"
        else
            Header = "";//Volatility Profile
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmss");

        //  String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";
        String fsFinalLocation = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".xls";

        AppDomain.CurrentDomain.Load("JCommon");
        org.jfree.data.general.DefaultPieDataset myDataSet = new org.jfree.data.general.DefaultPieDataset();
        if (table.Rows.Count > 0)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                if (Convert.ToString(table.Rows[i][2]) != "")
                {
                    string ColValue = String.Format("{0:#,###0.0;-#,###0.0}", Convert.ToDecimal(table.Rows[i][2]));
                    //myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble(Math.Round(Convert.ToDecimal(table.Rows[i][2]), 1)));
                    //if (ChartType == "1")
                    //    myDataSet.setValue(Convert.ToString(table.Rows[i][3]), Convert.ToDouble(ColValue));
                    //else
                    if (ColValue == "0.0")
                        myDataSet.setValue(Convert.ToString(table.Rows[i][0]), Convert.ToDouble("0.0001"));
                    else
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
        ColorConfigurator.setLabelFont(new System.Drawing.Font("Frutiger-Roman", 14));//, System.Drawing.FontStyle.Bold

        ColorConfigurator.setCircular(false);
        //ColorConfigurator.setOutlineVisible(false);  fix to remove border
        ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardPieSectionLabelGenerator("{0} {1}%"));
        //ColorConfigurator.setLabelGenerator(null);
        //ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardCategoryItemLabelGenerator("{0}"));
        //ColorConfigurator.setLabelGenerator(new org.jfree.chart.labels.StandardCategoryItemLabelGenerator("{0} =  {1}%"));
        //ColorConfigurator.setInteriorGap(0.30);

        //ColorConfigurator.setInteriorGap(0.15);
        ColorConfigurator.setInteriorGap(0.25);
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
        if (dt.Rows.Count > 0)
            retVal = Convert.ToString(dt.Rows[0][GetDataFromColumn]);
        else
            retVal = "0";
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
        //strText = "Gresham Advisors, LLC | 333 W. Wacker Dr. Suite 700 | Chicago, IL 60606 | P 312.960.0200 | F 312.960.0204 | www.greshampartners.com";
        Phrase footPhraseImg = new Phrase("test", setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter footer = new HeaderFooter(footPhraseImg, footPhraseImg);
        // HeaderFooter footer1 = new HeaderFooter(footPhraseImg, false);
        //footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
        footer.Border = 0;
        footer.Alignment = Element.ALIGN_LEFT;
        document.Footer = footer;

        //    document.Footer = footer1;
    }

    private decimal RoundPercent(decimal Value)
    {
        Value = Convert.ToDecimal(String.Format("{0:#,###0.0;-#,###0.0}", Value));
        return Value;
    }
    private decimal RoundValue(decimal Value)
    {
        Value = Convert.ToDecimal(String.Format("{0:#,###0;-#,###0}", Value));
        return Value;
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

    public void SetBorder(PdfPCell foCell, bool IsTop, bool IsBottom, bool IsLeft, bool IsRight)
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
        if (lsFormatedString != "")
        {
            lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
            return lsFormatedString;
        }
        else
            return "";
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
    public iTextSharp.text.Font setFontsAllFrutiger(int size, int bold, int italic, iTextSharp.text.Color foColor)
    {
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
        fotable.Cellpadding = 3f;
        fotable.Locked = false;

    }

    public void setTableProperty(PdfPTable fotable)
    {
        //int[] headerwidths = { 28, 9, 9, 9, 9, 9, 9, 9, 7 };

        setWidthsoftable(fotable);

        //fotable.Width = 100;

        fotable.HorizontalAlignment = 1;
        fotable.WidthPercentage = 100;



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
                int[] headerwidths3 = { 40, 60, 40 };
                fotable.SetWidths(headerwidths3);
                fotable.Width = 100;
                break;
            case "4":
                int[] headerwidths4 = { 45, 5, 15, 15 };
                fotable.SetWidths(headerwidths4);
                fotable.Width = 80;
                break;
            case "5":
                int[] headerwidths5 = { AllocationColumnSize, 12, 12, 12, 12 };
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

    public void setWidthsoftable(PdfPTable fotable)
    {

        switch (lsTotalNumberofColumns)
        {
            case "2":
                int[] headerwidths2 = { 70, 18 };
                fotable.SetWidths(headerwidths2);
                fotable.TotalWidth = 88;
                break;
            case "3":
                int[] headerwidths3 = { 41, 60, 40 };
                fotable.SetWidths(headerwidths3);
                fotable.TotalWidth = 100;
                break;
            case "4":
                int[] headerwidths4 = { 45, 5, 15, 15 };
                fotable.SetWidths(headerwidths4);
                fotable.TotalWidth = 80;
                break;
            case "5":
                int[] headerwidths5 = { AllocationColumnSize, 12, 12, 12, 12 };
                fotable.SetWidths(headerwidths5);
                fotable.TotalWidth = 100;
                break;
            case "6":
                int[] headerwidths6 = { 27, 11, 11, 8, 5, 7 };
                fotable.SetWidths(headerwidths6);
                fotable.TotalWidth = 70;
                break;
            case "7":
                int[] headerwidths7 = { 30, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths7);
                fotable.TotalWidth = 85;
                break;
            case "8":
                int[] headerwidths8 = { 30, 9, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths8);
                fotable.TotalWidth = 94;
                break;
            case "9":
                int[] headerwidths9 = { 27, 9, 9, 9, 9, 9, 9, 9, 7 };
                fotable.SetWidths(headerwidths9);
                fotable.TotalWidth = 97;
                break;

            case "10":
                int[] headerwidths10 = { 25, 8, 8, 8, 8, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths10);
                fotable.TotalWidth = 97; break;
            case "11":
                //int[] headerwidths11 = { 25, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7 };
                int[] headerwidths11 = { 25, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths11);
                fotable.TotalWidth = 95; break;
            case "21":
                int[] headerwidths12 = { 7, 7, 12, 7, 7, 7, 7, 35, 7, 7, 7, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                fotable.SetWidths(headerwidths12);
                fotable.TotalWidth = 150; break;
            case "13":
                int[] headerwidths13 = { 12, 2, 15, 2, 20, 20, 2, 15, 5, 15, 15, 2, 15 };
                fotable.SetWidths(headerwidths13);
                fotable.TotalWidth = 100; break;
            case "14":
                int[] headerwidths14 = { 30, 9 };
                fotable.SetWidths(headerwidths14);
                fotable.TotalWidth = 100;
                break;
            case "15":
                int[] headerwidths15 = { 15, 2, 15, 2, 15, 2, 15, 2, 15, 2, 15, 2, 15, 2, 15 };
                fotable.SetWidths(headerwidths15);
                fotable.TotalWidth = 100; break;
                break;
            case "16":
                int[] headerwidths16 = { 30, 9 };
                fotable.SetWidths(headerwidths16);
                fotable.TotalWidth = 39;
                break;
            case "17":
                int[] headerwidths17 = { 30, 9 };
                fotable.SetWidths(headerwidths17);
                fotable.TotalWidth = 39;
                break;
            case "18":
                int[] headerwidths18 = { 30, 9 };
                fotable.SetWidths(headerwidths18);
                fotable.TotalWidth = 39;
                break;
            case "19":
                int[] headerwidths19 = { 30, 9 };
                fotable.SetWidths(headerwidths19);
                fotable.TotalWidth = 39;
                break;
            case "29":
                int[] headerwidths20 = { 30, 9 };
                fotable.SetWidths(headerwidths20);
                fotable.TotalWidth = 39;
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
