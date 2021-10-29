using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Globalization;
using System.Text.RegularExpressions;

public partial class BillingRpt : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        string EndDate = null;

        DateTime dtAsOfDate1 = DateTime.Now;
        DateTime lastDay1 = new DateTime(dtAsOfDate1.Year, dtAsOfDate1.Month, 1); //1st Day of Current Month
        lastDay1 = lastDay1.AddDays(-1);  //last date of previous month

        DateTime date1 = DateTime.Now;
        DateTime quarterEn1d = NearestQuarterEnd(date1);
        EndDate = quarterEn1d.ToShortDateString();
        EndDate = "'" + EndDate + "'";



        if (!IsPostBack)
        {
            DateTime dtAsOfDate = DateTime.Now;
            DateTime lastDay = new DateTime(dtAsOfDate.Year, dtAsOfDate.Month, 1); //1st Day of Current Month
            lastDay = lastDay.AddDays(-1);  //last date of previous month

            txtAUMDate.Text = lastDay.ToString("MM/dd/yyyy");
            DateTime date = DateTime.Now;
            DateTime quarterEnd = NearestQuarterEnd(date);
            txtAUMDate.Text = quarterEnd.ToShortDateString();

            BindAdvisors();
            BindHouseHold();
            BindBillingName();
        }
    }
    public DateTime NearestQuarterEnd(DateTime date)
    {
        IEnumerable<DateTime> candidates =
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
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        
        string BillingUUID = "";
        object HHValue = ddlHH.SelectedValue == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlHH.SelectedValue + "'";
        string strBillFor = ddlBillFor.SelectedValue;
        char[] delimiterChars = { '|' };
        string[] words = strBillFor.ToString().Split(delimiterChars);
        int len = words.Length;
        if (len > 0)
        {
            strBillFor = words[0];
            BillingUUID = words[0];
        }
        List<string> ls = new List<string>();
        DB clsDB = new DB();

        string SQLQuery = "SP_S_BILLINGINVOICE_CHECK @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ";

        DataSet dsBillingChk = clsDB.getDataSet(SQLQuery);
        for (int i = 0; i < dsBillingChk.Tables[1].Rows.Count; i++)
        {
            string data = dsBillingChk.Tables[1].Rows[i]["ssi_billingId"].ToString();
            // SourceFileArray[i] = data;
            ls.Add(data);
        }

       DataTable dtBillingforList = dsBillingChk.Tables[1];

       if (ls.Count == 0)
       {
           SQLQuery = "SP_S_HOUSEHOLD_BILLING @HouseholdID=" + HHValue+"";
           DataSet dsBillings = clsDB.getDataSet(SQLQuery);

           for (int i = 0; i < dsBillings.Tables[0].Rows.Count; i++)
           {
               string data = dsBillings.Tables[0].Rows[i]["ssi_billingId"].ToString();
              
               ls.Add(data);
           }

       }
       //else
       //    lblMessage.Text = "No Data Found";

       PDFBillingWorksheetAndInvoice(ls, dtBillingforList);

       // GenratePortfolioOld();
    }

    public void PDFBillingWorksheetAndInvoice(List<string> list, DataTable dtBillingforList)
    {

        string[] SourceFileArray;
        //List<string> list = new List<string>();
        //list = (List<string>)ViewState["sourceLsit"];
        //DataTable dtBillingforList = (DataTable)ViewState["BillingForList"];

        int rowCount = dtBillingforList.Rows.Count;
        string sql = "EXEC SP_R_BILLING @ReportFlg = 1,@PdfFlg=1, @BillingForUUID = '";

        //foreach (string UID in SourceFileArray)
        //{
        //    sql = sql + UID + ",";
        //}

        for (int i = 0; i < list.Count(); i++)
        {
            sql = sql + list[i].ToString() + ",";
        }

        sql = sql.Substring(0, sql.LastIndexOf(","));
        sql = sql + "',@AsOfDate = '" + txtAUMDate.Text + "'";
        try
        {
            String DestinationPath = null;
            String Path1 = null;
            DataSet newdataset;
            DB clsDB = new DB();
            newdataset = null;
            //  string lsSql = "EXEC SP_R_BILLING @BillingForUUID = '801DE384-D1A4-E511-9418-005056A0567E,841DE384-D1A4-E511-9418-005056A0567E,821DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'";
            newdataset = clsDB.getDataSet(sql);

            int noOfPdf = newdataset.Tables.Count;

           
                SourceFileArray = new string[noOfPdf + 1];
                if (rowCount > 0)
                {
                string vFileInvoice = PDFInvoic();
                int cnt = SourceFileArray.Length;
                if (vFileInvoice != null)
                    SourceFileArray[0] = vFileInvoice;

            }
            else
            {
                
            }
            string aod = txtAUMDate.Text;
            DateTime dAsofDate = DateTime.Now;
            try
            {
                dAsofDate = DateTime.ParseExact(aod, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            }
            catch
            {
                dAsofDate = DateTime.ParseExact(aod, "M/dd/yyyy", CultureInfo.InvariantCulture);
            }

            if (noOfPdf > 0)
            {


                for (int i = 0; i < noOfPdf; i++)
                {
                    if (newdataset.Tables[i].Rows.Count > 0)
                    {
                        DataTable dt;
                        dt = newdataset.Tables[i];
                        DataRow dr = dt.Rows[0];
                        string HHName = dr["BillingName"].ToString();

                        string path = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + HHName + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString();
                        string filename = GeneratePdf(dt, i);
                        SourceFileArray[i + 1] = filename;
                    }
                }
                // String DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "MergedReport.pdf";
                if (rowCount == 1)
                {
                    //String DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "MergedReport.pdf";
                    DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + ddlHH.SelectedItem.Text + "-" + ddlBillFor.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd") + ".pdf";
                    Path1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + ddlHH.SelectedItem.Text + "-" + ddlBillFor.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd");
                }
                else
                {
                    DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + ddlHH.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd") + ".pdf";
                    Path1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + ddlHH.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd");
                }
                if (System.IO.File.Exists(DestinationPath))
                {
                    System.IO.File.Delete(DestinationPath);
                }



                if (SourceFileArray.Count() > 0)
                {
                    MergeFiles(DestinationPath, SourceFileArray);

                    //string filenmae = Path.GetFileName(DestinationPath);
                    //FileInfo loFile = new FileInfo(DestinationPath);
                    //loFile.MoveTo(DestinationPath.Replace(".xls", ".pdf"));
                    //Response.Write("<script>");
                    //string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/"+ filenmae;
                    //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    //  string Path1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + ddlHH.SelectedItem.Text + dAsofDate.ToString("yyyy-MMdd") + "_TEST";


                    //CopytoSharepoint(Path1, DestinationPath);

                    if (rowCount == 0)
                    {
                        lblMessage.Text = "Billing Invoice Not Available";
                    }
                }
                else
                {
                    lblMessage.Text = "No Data Found";
                }
            }
            else
            {
                lblMessage.Text = "No Data Found";
            }

          //  lblMessage.Text = "Records Updated & File Saved Successfully";
        }
        catch (Exception ex)
        {
            lblMessage.Text = "error" + ex;
        }
    }

    #region old PDF
    public void GenratePortfolioOld1()
    {
        #region blue color

        int red = 183;
        int green = 221;
        int blue = 232;

        #endregion

        #region grey color

        int red1 = 233;
        int green1 = 237;
        int blue1 = 241;

        #endregion

        #region yellow color

        int red2 = 252;
        int green2 = 252;
        int blue2 = 181;

        #endregion

        #region Green
        int red3 = 198;
        int green3 = 239;
        int blue3 = 206;

        #endregion


        try
        {
            // int liPageSize = 26;  //--> Original Value
            int liPageSize = 28;

            DataSet newdataset;
            DB clsDB = new DB();
            newdataset = null;
            //int liPageSize = 18;
            //int liCurrentPage = 0;
            //lstAssetClass.s

            string billingVal = ddlBillFor.SelectedValue.ToString();
            string dDate = txtAUMDate.Text;
            string[] date = dDate.Split(new char[] { '/' });
            string day = null, Month = null;
            if (dDate != "")
            {
                if (date[0].Length == 1)
                    Month = "0" + date[0];
                else
                    Month = date[0];

                if (date[1].Length == 1)
                    day = "0" + date[1];
                else
                    day = date[1];

                dDate = Month + "/" + day + "/" + date[2];
                txtAUMDate.Text = dDate;
            }

            //  String lsSQL = "EXEC SP_R_BILLING @BillingForUUID = '7E1DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'";

            string lsSQL = "EXEC  SP_R_BILLING @ReportFlg = 1,@PdfFlg=1,  @BillingForUUID = '" + billingVal + "',@AsOfDate = '" + dDate + "'";

            newdataset = clsDB.getDataSet(lsSQL);

            int DSCount = newdataset.Tables[0].Rows.Count;
            DataTable table = newdataset.Tables[0].Copy();
            table.Columns["PosAssetClassName"].SetOrdinal(0);
            // table.Columns["SecurityName"].SetOrdinal(1);
            //table.Columns["AccountLegalEntityName"].SetOrdinal(2);
            table.Columns["PdfAccountName"].SetOrdinal(1);
            table.Columns["PortCode"].SetOrdinal(2);
            table.Columns["ssi_BillingMarketValue"].SetOrdinal(3);
            table.Columns["FinalBillingMarketValue"].SetOrdinal(4);
            table.Columns["FinalAUMMarketValue"].SetOrdinal(5);

            Random rand = new Random();
            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


            //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);
            //  iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
            String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingrpt.pdf";
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

            string HHName = ddlBillFor.SelectedItem.Text;
            //string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
            //if (strTitle != "")
            //    HHName = strTitle;

            string strheader = "BILLING WORKSHEET";
            string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";
            DateTime AUMdate = DateTime.ParseExact(txtAUMDate.Text, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            string vAUmdate = AUMdate.ToString(" MMMM dd, yyyy");
            //   DateTime asofDT = Convert.ToDateTime("12/31/2015");
            //   string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);
            //   dd MMMM ,yyyy

            iTextSharp.text.Table loTable = new iTextSharp.text.Table(6, table.Rows.Count);   // 2 rows, 2 columns           
            // lsTotalNumberofColumns = "9";
            iTextSharp.text.Cell loCell = new Cell();


            #region Table Style
            int[] headerwidths9 = { 17, 31, 10, 13, 13, 13 };
            loTable.SetWidths(headerwidths9);
            loTable.Width = 100;
            loTable.Border = 0;
            loTable.Cellspacing = 0;
            loTable.Cellpadding = 3;
            loTable.Locked = false;

            #endregion

            iTextSharp.text.Chunk lochunk = new Chunk();
            iTextSharp.text.Chunk lochunknew = new Chunk();
            Paragraph para = new Paragraph();
            int rowsize = table.Rows.Count;
            int colsize = 6;

            String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
            int liTotalPage = (table.Rows.Count / liPageSize);
            int liCurrentPage = 0;

            if (table.Rows.Count % liPageSize != 0)
            {
                liTotalPage = liTotalPage + 1;
            }
            else
            {
                // liPageSize = 26; //--> Original Value
                liPageSize = 28;
                liTotalPage = liTotalPage + 1;
            }

            for (int i = 0; i < rowsize; i++)
            {
                //string BillExcepTyp = Convert.ToString(table.Rows[i]["BillingExceptionType"]);
                //string AUMExcepTyp = Convert.ToString(table.Rows[i]["AUMExceptionType"]);

                string BillExcepTyp = Convert.ToString(table.Rows[i]["ColourBillingExceptionType"]);
                string AUMExcepTyp = Convert.ToString(table.Rows[i]["colourAUMExceptionType"]);
                string BillFeeTyp = Convert.ToString(table.Rows[i]["BillingFeeExceptionId"]);
                if (i % liPageSize == 0)
                {
                    document.Add(loTable);

                    if (i != 0)
                    {
                        liCurrentPage = liCurrentPage + 1;
                        //   document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                        document.NewPage();
                        //SetTotalPageCount("Short Term Performance");
                    }
                    loTable = new iTextSharp.text.Table(6, table.Rows.Count);
                    //  int[] headerwidths = { 27, 9, 9, 9, 12, 12, 12, 12, 10 };
                    loTable.SetWidths(headerwidths9);
                    loTable.Width = 100;

                    loTable.Border = 0;
                    loTable.Cellspacing = 0;
                    loTable.Cellpadding = 3;
                    loTable.Locked = false;

                    lochunk = new Chunk(HHName, setFontsAllFrutiger(14, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Colspan = 6;
                    loCell.HorizontalAlignment = 1;


                    lochunk = new Chunk("\n" + strheader, setFontsAllFrutiger(10, 0, 0));
                    loCell.Add(lochunk);

                    // lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                    // loCell.Add(lochunk);

                    lochunk = new Chunk("\n " + vAUmdate + " \n", setFontsAllFrutiger(10, 0, 1));
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loTable.AddCell(loCell);



                    // Add table columns
                    for (int k = 0; k < colsize; k++)
                    {
                        string ColHeader = Convert.ToString(table.Columns[k].ColumnName);
                        if (ColHeader == "PosAssetClassName")
                            ColHeader = "Asset";
                        //if (ColHeader.Contains("SecurityName"))
                        //    ColHeader = "Security Name";
                        if (ColHeader.Contains("PdfAccountName"))
                            ColHeader = "Legal Entity (Custodian Account)";
                        if (ColHeader.Contains("PortCode"))
                            ColHeader = "Port Code";
                        if (ColHeader.Contains("ssi_BillingMarketValue"))
                            ColHeader = "Total";
                        if (ColHeader.Contains("FinalBillingMarketValue"))
                            ColHeader = "Billing";
                        if (ColHeader.Contains("FinalAUMMarketValue"))
                            ColHeader = "AUM";

                        lochunk = new Chunk(ColHeader, setFontsAll(8, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        //loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        loCell.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        //if (k != 0)
                        //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        //else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }

                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                    //  png.SetAbsolutePosition(45, 557);//540
                    png.SetAbsolutePosition(43, 800); // potrait logo
                    png.ScalePercent(10);
                    document.Add(png);
                }

                for (int j = 0; j < colsize; j++)
                {
                    // string cellBackgroundColor = Convert.ToString(table.Rows[i]["ColourCode"]);
                    string ColValue = Convert.ToString(table.Rows[i][j]);
                    iTextSharp.text.Cell loCell1 = new iTextSharp.text.Cell();

                    #region not used
                    //loCell = new iTextSharp.text.Cell();
                    //if (ColValue == "Strategy Benchmark")
                    //    ColValue = "Weighted Benchmark";
                    //if (ColValue == "Gresham Advised Values")
                    //    ColValue = "Gresham Advised";

                    //if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True" || Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "" || i == 0 || i == 1)
                    //{
                    //    if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                    //    {
                    //        if (ColValue.Contains("-"))
                    //        {
                    //            ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                    //            ColValue = ColValue.Replace("(", "($");
                    //        }
                    //        else
                    //            ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                    //        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //    else if (j == 8 && ColValue != "")
                    //    {
                    //        ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //        lochunk = new Chunk(ColValue + "%", setFontsAll(7, 1, 0));
                    //    }
                    //    else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                    //    {
                    //        if (ColValue != "N/A")
                    //            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //    else
                    //    {
                    //        if ((i == 0 || i == 1) && j == 0)
                    //        {
                    //            string[] str = ColValue.Split('(');

                    //            lochunk = new Chunk(str[0], setFontsAll(7, 1, 0));
                    //            if (str.Length > 1)
                    //                lochunknew = new Chunk("\n(" + str[1], setFontsAll(6, 1, 0));
                    //        }
                    //        else if (j == 0 && cellBackgroundColor != "")
                    //        {
                    //            lochunk = new Chunk("   " + ColValue, setFontsAll(7, 1, 0));
                    //        }
                    //        else if (j == 0 && (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True"))
                    //        {
                    //            lochunk = new Chunk("       " + ColValue, setFontsAll(7, 1, 0));
                    //        }
                    //        else
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //}
                    //else
                    //{
                    //    if (j == 0) //component
                    #endregion


                    if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True") // if it is total asset
                    {
                        //if (j == 0 && ColValue != "")
                        //    ColValue = ColValue + " Total";

                        if ((j == 3 || j == 4 || j == 5) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = currencyFormat(ColValue);
                                //  ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                //ColValue = String.Format("{0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));
                                //ColValue = ColValue.Replace("(", "($");

                            }
                            //else if (Convert.ToDecimal(ColValue) == 0)
                            //{
                            //    ColValue = "$0.00";
                            //}
                            else
                            {
                                ColValue = currencyFormat(ColValue);
                                // ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                // ColValue = String.Format("${0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));

                            }
                            loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;


                        }

                        if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True" || (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))
                        {
                            //loCell1.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        }
                        //if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        //{
                        //    // loCell.BackgroundColor = new iTextSharp.text.Color(204, 206, 219);
                        //    loCell.BackgroundColor = new iTextSharp.text.Color(red1, green1, blue1);
                        //}

                        if (Convert.ToString(table.Rows[i]["DiffFlg"]) == "1")
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(red3, green3, blue3);
                        }



                        //lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        # region Color on PDF


                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            // lochunk.SetBackground(Color.YELLOW);
                            //lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }
                        //----------------------------------------------------------------------------------------------------------
                        else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            para = new Paragraph();
                            //   para.IndentationLeft = 20;
                            para.SpacingBefore = 0;
                            para.SpacingAfter = 0;
                            para.Leading = 6F;
                            para.Add(lochunk);
                        }
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True")
                        {

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        }
                        #endregion
                    }
                    else
                    {
                        if ((j == 3 || j == 4 || j == 5) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = currencyFormat(ColValue);
                                //  ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                //ColValue = String.Format("{0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));
                                //ColValue = ColValue.Replace("(", "($");

                            }
                            //else if (Convert.ToDecimal(ColValue) == 0)
                            //{
                            //    ColValue = "$0.00";
                            //}
                            else
                            {
                                ColValue = currencyFormat(ColValue);
                                // ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                // ColValue = String.Format("${0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));

                            }

                            loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        }

                        if (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")
                        {
                            if (j == 0)
                                ColValue = "TOTALS";

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));


                        }
                        //-------------------------------------------------------------------------------------------
                        else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            para = new Paragraph();
                            para.IndentationLeft = 20;
                            para.Leading = 6F;
                            // para.SpacingBefore = 0;
                            // para.SpacingAfter = 0;

                            para.Add(lochunk);
                        }
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True")
                        {

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        }

                        # region Color on PDF
                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != "") &&  (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 5 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != "") && (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }  
                        else if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            // lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);

                        }




                        //else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        //{
                        //    lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        //    para = new Paragraph();
                        //    para.IndentationLeft = 20;
                        //    para.SpacingBefore = 0;
                        //    para.SpacingAfter = 0;

                                                //    para.Add(lochunk);
                        //}
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));

                        }


                        #endregion

                    }
                    //    else
                    //    {
                    //        if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                    //        {
                    //            if (ColValue.Contains("-"))
                    //            {
                    //                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                    //                ColValue = ColValue.Replace("(", "($");
                    //            }
                    //            else
                    //                ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //        else if (j == 8 && ColValue != "")
                    //        {
                    //            lochunk = new Chunk(Convert.ToDecimal(ColValue) + "%", setFontsAll(7, 0, 0));
                    //        }
                    //        else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                    //        {
                    //            if (ColValue != "N/A")
                    //                ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //        else
                    //        {
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //    }
                    //}

                    // loCell1 = new iTextSharp.text.Cell();
                    // loCell1.BorderWidth = 1;
                    loCell1.BorderColor = Color.LIGHT_GRAY;
                    if (ColValue == "TOTALS")
                    {
                        loCell1.Add(lochunk);
                    }
                    else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                    {

                        loCell1.Add(para);
                    }

                    else
                    {

                        loCell1.Add(lochunk);
                    }
                    // loCell1.Add(lochunk);
                    // if ((i == 0 || i == 1) && j == 0)
                    //   loCell.Add(lochunknew);
                    // loCell1.Border = 0;

                    //if (i == 0 || i == 1 || i == 4)
                    //    loCell.Leading = 8F;
                    //else
                    //{
                    //    //  if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "True" && Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "")
                    //    //  loCell.Leading = 0F;
                    //    // else
                    //    // loCell.Leading = 4F;
                    //}

                    loCell1.Leading = 6F;
                    loCell1.VerticalAlignment = 5;

                    //if (j != 0)
                    //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    //if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Gresham Advised Values")
                    //    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                    //else if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")
                    //    loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;

                    //if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True" || (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))
                    //{
                    //    //loCell1.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                    //    loCell1.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                    //}
                    if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                    {
                        loCell1.BackgroundColor = new iTextSharp.text.Color(red1, green1, blue1);
                    }
                    if (Convert.ToString(table.Rows[i]["DiffFlg"]) == "1")
                    {
                        loCell1.BackgroundColor = new iTextSharp.text.Color(red3, green3, blue3);
                    }

                    //if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components")
                    //{
                    //    loCell.BorderColorBottom = iTextSharp.text.Color.WHITE;
                    //    loCell.BorderWidthBottom = 1.5f;
                    //}



                    //  if ((Convert.ToString(table.Rows[i]["ColourCode"]) != "" || (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")) && j != 0)
                    //  { 
                    //  }
                    //  else
                    // {
                    //  loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    loTable.AddCell(loCell1);

                    //  }
                }

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
                    liCurrentPage = liCurrentPage + 1;
                    //PdfPTable TabFooter = addFooter(lsDateTime, true, lsFooterText, false, 0, 0);
                    //TabFooter.WidthPercentage = 100f;
                    ////  TabFooter.TotalWidth = 100f;
                    //TabFooter.TotalWidth = 775;

                    //TabFooter.WriteSelectedRows(0, 4, 30, 47, writer.DirectContent);  --> Original Values
                    //TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
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
                try
                {

                    FileInfo loFile = new FileInfo(ls);
                    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                    Response.Write("<script>");
                    string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
                    Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    //  Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                    Response.Write("</script>");
                }
                catch (Exception ex)
                {
                    Response.Write(ex.ToString());
                }

            }
            //  return fsFinalLocation.Replace(".xls", ".pdf");

        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
        }


    }
    #endregion

    public string PDFInvoic()
    {
        string BillingUUID = "";
        #region blue color

        int red = 183;
        int green = 221;
        int blue = 232;

        #endregion

        #region grey color

        int red1 = 233;
        int green1 = 237;
        int blue1 = 241;

        #endregion

        #region yellow color

        int red2 = 252;
        int green2 = 252;
        int blue2 = 181;

        #endregion

        #region Green
        int red3 = 198;
        int green3 = 239;
        int blue3 = 206;

        #endregion

        String ls = null;
        try
        {
            //liPageSize = 26; --> Original Value
            int liPageSize = 30;

            DataSet newdataset;
            DB clsDB = new DB();
            newdataset = null;
            //int liPageSize = 18;
            //int liCurrentPage = 0;
            //lstAssetClass.s

            string strBillFor = ddlBillFor.SelectedValue;
            char[] delimiterChars = { '|' };
            string[] words = strBillFor.ToString().Split(delimiterChars);
            int len = words.Length;
            if (len > 0)
            {
                strBillFor = words[0];
                BillingUUID = words[0];
            }

            object HHValue = ddlHH.SelectedValue == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlHH.SelectedValue + "'";


            //   "SP_S_BILLINGINVOICE_CHECK @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");

            // String lsSQL = "SP_R_ANNUALFEECALCULATION";
            //   String lsSQL = "EXEC SP_R_BILLING @BillingForUUID='7E1DE384-D1A4-E511-9418-005056A0567E' , @AsOfDate='20160331'";

            String lsSQL = "EXEC SP_R_ANNUALFEECALCULATION @BillingForUUID='" + BillingUUID + "',@AumAsodfDate='" + txtAUMDate.Text + "',@HouseHoldUUID=" + HHValue;

            newdataset = clsDB.getDataSet(lsSQL);
            int DSCount = newdataset.Tables[0].Rows.Count;
            DataTable table = newdataset.Tables[0].Copy();
            DataTable table1 = newdataset.Tables[1].Copy();
            DataTable table2 = newdataset.Tables[2].Copy();
            table1.Columns.Remove("IdNmb");

            string ReportHeader = table.Rows[0]["Reportheader"].ToString();
            table.Columns.Remove("Reportheader");

            //if (table1.Rows.Count > 1)
            //{
            //    HHValue = strBillFor;
            //}


            Random rand = new Random();
            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 27, 31, 8);//10,10
            ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingFreeScheduleRpt.pdf";
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strGUID + ".xls";

            string HHName = ddlHH.SelectedItem.Text;// "TEST";
            //string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
            //if (strTitle != "")
            //    HHName = strTitle;

            string strheader = "Billing Report";

            //   DateTime asofDT = Convert.ToDateTime("12/31/2015");
            //   string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);


            PdfPTable loTable = new PdfPTable(3);
            PdfPTable loTableNote = new PdfPTable(1);
            //added 2 new column(Jeanne Request 5_29_2020)
            //PdfPTable loTable1 = new PdfPTable(6);
            PdfPTable loTable1 = new PdfPTable(8);
            PdfPTable loTableHeader = new PdfPTable(1);
            PdfPTable loTableNote1 = new PdfPTable(1);
            PdfPTable loTableNote2 = new PdfPTable(1);
            loTable.HorizontalAlignment = 0;

            PdfPCell loCell = new PdfPCell();
            PdfPCell loCell1 = new PdfPCell();


            #region Table Style
            int[] headerwidths9 = { 10, 10, 10 };
            loTable.SetWidths(headerwidths9);
            loTable.WidthPercentage = 85f;
            // loTable.SpacingBefore = 20f;
            //  loTable.SpacingAfter = 60f;
            loTable.HorizontalAlignment = 0;

            //loTable.Border = 0;
            //loTable.Cellspacing = 0;
            //loTable.Cellpadding = 3;
            //loTable.Locked = false;

            #endregion

            #region Table 1 Style
            //added 2 new column(Jeanne Request 5_29_2020)
            // int[] headerwidths = { 18, 10, 10, 13, 13, 5 };
            int[] headerwidths = { 25, 12, 12, 12, 12, 12, 12, 9 };
            //loTable1.WidthPercentage = 85f;
            loTable1.WidthPercentage = 100f;
            loTable1.SetWidths(headerwidths);
            loTable1.HorizontalAlignment = 0;

            #endregion

            #region Table notes style

            int[] headerwidths1 = { 100 };
            loTableNote.WidthPercentage = 100f;
            loTableNote.SetWidths(headerwidths1);
            loTableNote.HorizontalAlignment = 0;

            #endregion

            #region Table  Header

            int[] headerwidths2 = { 100 };
            loTableHeader.WidthPercentage = 100f;
            loTableHeader.SetWidths(headerwidths2);
            loTableHeader.HorizontalAlignment = 0;

            #endregion


            #region Table notes1 style

            int[] headerwidths3 = { 100 };
            loTableNote1.WidthPercentage = 85f;
            loTableNote1.SetWidths(headerwidths3);
            loTableNote1.HorizontalAlignment = 0;

            #endregion

            #region Table notes2 style

            int[] headerwidths4 = { 100 };
            loTableNote2.WidthPercentage = 85f;
            loTableNote2.SetWidths(headerwidths4);
            loTableNote2.HorizontalAlignment = 0;

            #endregion

            Paragraph lochunk = new Paragraph();
            Paragraph lochunknew = new Paragraph();

            int rowsize = table.Rows.Count;
            int colsize = 3;

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
                        //   document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                        document.NewPage();
                        //SetTotalPageCount("Short Term Performance");
                    }
                    loTable = new PdfPTable(3);
                    //  int[] headerwidths = { 27, 9, 9, 9, 12, 12, 12, 12, 10 };
                    loTable.SetWidths(headerwidths9);
                    //  loTable.SpacingBefore = 20f;
                    //  loTable.SpacingAfter = 60f;
                    loTable.WidthPercentage = 85f;
                    loTable.HorizontalAlignment = 0;
                    //loTable.Border = 0;
                    //loTable.Cellspacing = 0;
                    //loTable.Cellpadding = 3;
                    //loTable.Locked = false;

                    // lochunk = new Paragraph(HHName, setFontsAllFrutiger(14, 1, 0));
                    lochunk = new Paragraph(ReportHeader, setFontsAllFrutiger(14, 1, 0));
                    lochunk.SetAlignment("center");
                    loCell = new PdfPCell();
                    loCell.AddElement(lochunk);
                    //  loCell.Colspan = 3;
                    loCell.Border = 0;
                    //    loCell.HorizontalAlignment = 2;


                    lochunk = new Paragraph(strheader, setFontsAllFrutiger(10, 0, 0));
                    lochunk.SetAlignment("center");
                    loCell.AddElement(lochunk);
                    loCell.Border = 0;
                    // lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                    // loCell.Add(lochunk);

                    lochunk = new Paragraph(txtAUMDate.Text, setFontsAllFrutiger(10, 0, 1));
                    lochunk.SetAlignment("center");
                    loCell.AddElement(lochunk);
                    loCell.Border = 0;
                    loCell.PaddingBottom = 10f;
                    loTableHeader.AddCell(loCell);

                    document.Add(loTableHeader);

                    for (int k = 0; k < colsize; k++)
                    {
                        string ColHeader = Convert.ToString(table.Columns[k].ColumnName);


                        lochunk = new Paragraph(ColHeader, setFontsAll(8, 1, 0));
                        lochunk.SetAlignment("center");

                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        //loCell.Leading = 11F;
                        loCell.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        //  loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        //if (k != 0)
                        //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        //else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loCell.VerticalAlignment = Element.ALIGN_TOP;
                        loCell.PaddingBottom = 6f;
                        loTable.AddCell(loCell);
                    }

                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                    //  png.SetAbsolutePosition(45, 600);//540(43, 800)
                    png.SetAbsolutePosition(43, 800);
                    png.ScalePercent(10);
                    document.Add(png);
                }
                for (int j = 0; j < colsize; j++)
                {
                    // string cellBackgroundColor = Convert.ToString(table.Rows[i]["ColourCode"]);
                    string ColValue = Convert.ToString(table.Rows[i][j]);
                    string Desc = Convert.ToString(table.Rows[i]["Description"]);
                    string IndentFlg = Convert.ToString(table.Rows[i]["IndentFlg"]);
                    lochunk = new Paragraph(ColValue, setFontsAll(7, 0, 0));

                    loCell = new PdfPCell();
                    if ((j == 0) && ColValue != "")
                    {
                        if (IndentFlg == "False")
                        {
                            lochunk = new Paragraph(ColValue, setFontsAll(7, 1, 0));
                            //loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            lochunk.SetAlignment("left");

                        }
                        else
                        {

                            Chunk lchunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            lochunk = new Paragraph();
                            lochunk.IndentationLeft = 20;
                            //lochunk.Leading = 6F;
                            // para.SpacingBefore = 0;
                            // para.SpacingAfter = 0;

                            lochunk.Add(lchunk);
                        }
                    }
                    else if ((j == 1) && ColValue != "")
                    {
                        if (Desc.Contains("Fee Rate"))
                        {
                            ColValue = String.Format("{0:#,###0.00;(#,###0.00)}", Convert.ToDecimal(ColValue));
                            lochunk = new Paragraph(ColValue + "%", setFontsAll(7, 0, 0));
                        }
                        else
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                ColValue = ColValue.Replace("(", "($");
                            }
                            else
                                ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                            lochunk = new Paragraph(ColValue, setFontsAll(7, 0, 0));
                        }
                        lochunk.SetAlignment("right");

                    }
                    else
                    {
                        lochunk = new Paragraph(ColValue, setFontsAll(7, 0, 0));
                        //loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        lochunk.SetAlignment("left");
                    }

                    loCell.AddElement(lochunk);
                    // if ((i == 0 || i == 1) && j == 0)
                    //   loCell.Add(lochunknew);
                    // loCell.Border = 0;
                    loCell.BorderColor = Color.LIGHT_GRAY;

                    //if (i == 0 || i == 1 || i == 4)
                    //    loCell.Leading = 8F;
                    //else
                    //{
                    //    //  if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "True" && Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "")
                    //    //  loCell.Leading = 0F;
                    //    // else
                    //    // loCell.Leading = 4F;
                    //}
                    //  loCell.Leading = 4F;


                    // loCell.VerticalAlignment = Element.ALIGN_MIDDLE;

                    loTable.AddCell(loCell);

                }

                if (i == table.Rows.Count - 1)
                {
                    document.Add(loTable);
                    liCurrentPage = liCurrentPage + 1;
                }
            }

            if (table1.Rows.Count > 0)
            {
                Paragraph loParaBlank = new Paragraph();
                Paragraph loParaNote = new Paragraph();
                PdfPCell Pcell = new PdfPCell();
                Pcell.Border = 0;
                Pcell.PaddingTop = 23f; Pcell.PaddingBottom = 6f;

                loParaBlank = new Paragraph("*Total Annual Fee allocated as follows:", setFontsAllFrutiger(8, 0, 1));

                Pcell.AddElement(loParaBlank);
                loTableNote.AddCell(Pcell);
                document.Add(loTableNote);
            }

            #region table1 1


            Paragraph lochunk1 = new Paragraph();

            int rowsize1 = table1.Rows.Count;
            //added 2 new column(Jeanne Request 5_29_2020)
            // int colsize1 = 6;
            int colsize1 = 8;

            int liTotalPage1 = (table1.Rows.Count / liPageSize);
            int liCurrentPage1 = 0;

            if (table1.Rows.Count % liPageSize != 0)
            {
                liTotalPage1 = liTotalPage1 + 1;
            }
            else
            {
                //liPageSize = 26; --> Original Value
                liPageSize = 30;
                liTotalPage1 = liTotalPage1 + 1;
            }

            for (int i = 0; i < rowsize1; i++)
            {
                if (i % liPageSize == 0)
                {
                    document.Add(loTable1);

                    if (i != 0)
                    {
                        liCurrentPage1 = liCurrentPage1 + 1;
                        //   document.Add(addFooter("", liTotalPage1, liCurrentPage1, liPageSize, false, String.Empty));
                        document.NewPage();
                        //SetTotalPageCount("Short Term Performance");
                    }

                    for (int k = 0; k < colsize1; k++)
                    {
                        string ColHeader = Convert.ToString(table1.Columns[k].ColumnName);


                        if (ColHeader == "AUM Annual Fee")
                        {
                            ColHeader = "AUM\nAnnual Fee";
                        }
                        else   if (ColHeader == "Total Annual Fee")
                        {
                            ColHeader = "Total\nAnnual Fee";
                        }
                       // if (k == 0)
                       // { ColHeader = ""; }


                        lochunk1 = new Paragraph(ColHeader, setFontsAll(8, 1, 0));
                        lochunk1.SetAlignment("left");
                        loCell1 = new PdfPCell();
                        loCell1.AddElement(lochunk1);
                        loCell1.Border = 0;
                        // loCell1.Leading = 11F;
                        // loCell1.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        loCell1.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        //if (k != 0)
                        //    loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        //else
                        loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;

                        loCell1.PaddingBottom = 6f;
                        loTable1.AddCell(loCell1);
                    }

                    //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                    //png.SetAbsolutePosition(45, 557);//540
                    //png.ScalePercent(10);
                    //document.Add(png);
                }
                for (int j = 0; j < colsize1; j++)
                {
                    // string cellBackgroundColor = Convert.ToString(table1.Rows[i]["ColourCode"]);
                    string ColValue = Convert.ToString(table1.Rows[i][j]);

                    if ((j == 1 || j == 3 || j == 4 || j == 5 || j == 6) && ColValue != "")
                    {
                        if (ColValue.Contains("-"))
                        {
                            ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                            ColValue = ColValue.Replace("(", "($");
                        }
                        else
                            ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                        lochunk1 = new Paragraph(ColValue, setFontsAll(7, 0, 0));

                        lochunk1.SetAlignment("right");

                    }
                    else if ((j == 2 || j == 7) && ColValue != "")
                    {
                        ColValue = String.Format("{0:#,###0.00;(#,###0.00)}", Convert.ToDecimal(ColValue));
                        lochunk1 = new Paragraph(ColValue + "%", setFontsAll(7, 0, 0));
                        lochunk1.SetAlignment("right");
                    }
                    else
                        lochunk1 = new Paragraph(ColValue, setFontsAll(7, 1, 0));

                    loCell1 = new PdfPCell();
                    loCell1.AddElement(lochunk1);
                    // if ((i == 0 || i == 1) && j == 0)
                    //   loCell1.Add(lochunk1new);
                    // loCell1.Border = 0;

                    // loCell1.Leading = 4F;
                    loCell1.VerticalAlignment = 5;
                    loCell1.BorderColor = Color.LIGHT_GRAY;
                    loTable1.AddCell(loCell1);

                }

                if (i == table1.Rows.Count - 1)
                {
                    document.Add(loTable1);
                    liCurrentPage1 = liCurrentPage1 + 1;

                }
            }

            #endregion

            #region BlankSpace
            Paragraph ParaBlank = new Paragraph();

            PdfPCell BlankPcell1 = new PdfPCell();
            BlankPcell1.Border = 0;
            string blank = " ";
            ParaBlank = new Paragraph(blank, setFontsAllFrutiger(8, 0, 1));
            BlankPcell1.AddElement(ParaBlank);
            loTableNote1.AddCell(BlankPcell1);
            document.Add(loTableNote1);
            #endregion

            #region Table 2 Notes
            //table2.Rows[0].
            if (table2.Rows.Count > 0)
            {
                string val1 = table2.Rows[0]["AdvisoryNotes"].ToString();
                if (val1 != "")
                {
                    Paragraph loParaBlank1 = new Paragraph();
                    Paragraph loParaNote1 = new Paragraph();
                    PdfPCell Pcell1 = new PdfPCell();
                    Pcell1.Border = 0;
                    // Pcell1.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                    Pcell1.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                    string val = "Notes";
                    loParaBlank1 = new Paragraph(val, setFontsAllFrutiger(8, 0, 1));
                    Pcell1.AddElement(loParaBlank1);
                    loTableNote1.AddCell(Pcell1);
                    document.Add(loTableNote1);



                    Paragraph loParaBlank = new Paragraph();

                    PdfPCell Pcell = new PdfPCell();
                    Pcell.Border = 0;
                    //Pcell.BackgroundColor=
                    Pcell.PaddingTop = 10f; Pcell.PaddingBottom = 6f;
                    val1 = table2.Rows[0]["AdvisoryNotes"].ToString();
                    loParaBlank = new Paragraph(val1, setFontsAllFrutiger(8, 0, 1));

                    Pcell.AddElement(loParaBlank);
                    loTableNote2.AddCell(Pcell);
                    document.Add(loTableNote2);
                }
            }
            #endregion

            if (table.Rows.Count > 0)
            {
                document.Close();

                // FileInfo loFile = new FileInfo(ls);
                try
                {

                    //Response.ContentType = "Application/pdf";
                    //Response.AppendHeader("Content-Disposition", "attachment; abc.pdf");
                    //Response.TransmitFile(ls);

                    //FileInfo loFile = new FileInfo(ls);
                    //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                    //Response.Write("<script>");
                    //string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
                    //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    //Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                    //Response.Write("</script>");


                }
                catch (Exception ex)
                {
                    Response.Write(ex.ToString());
                }

            }
            //  return fsFinalLocation.Replace(".xls", ".pdf");

        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
        }
        return ls;

    }
    public string GeneratePdf(DataTable dataa, int num)
    {
        #region blue color

        int red = 183;
        int green = 221;
        int blue = 232;

        #endregion
        #region grey color

        int red1 = 233;
        int green1 = 237;
        int blue1 = 241;

        #endregion

        #region yellow color

        int red2 = 252;
        int green2 = 252;
        int blue2 = 181;

        #endregion

        #region Green
        int red3 = 193;
        int green3 = 239;
        int blue3 = 206;

        #endregion

        #region colorRed
        int Redred = 255;
        int Redgreen = 83;
        int Redblue = 83;
        #endregion

        String ls = null;
        try
        {
            // liPageSize = 24; --> Original Value
            //int liPageSize = 30;
            //    int liPageSize = 22;
            int liPageSize = 37;
            //DataSet newdataset;
            //DB clsDB = new DB();
            //newdataset = null;
            //int liPageSize = 18;
            //int liCurrentPage = 0;
            //lstAssetClass.s

            //String lsSQL = "EXEC SP_R_BILLING @BillingForUUID = '7E1DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'";
            //newdataset = clsDB.getDataSet(lsSQL);
            //int DSCount = newdataset.Tables[0].Rows.Count;
            // DataTable table = newdataset.Tables[0].Copy();
            DataTable table = dataa.Copy();
            table.Columns["PosAssetClassName"].SetOrdinal(0);
            //table.Columns["SecurityName"].SetOrdinal(1);
            // table.Columns["AccountLegalEntityName"].SetOrdinal(2);
            table.Columns["PdfAccountName"].SetOrdinal(1);
            table.Columns["PortCode"].SetOrdinal(2);
            table.Columns["ssi_BillingMarketValue"].SetOrdinal(3);
            table.Columns["FinalBillingMarketValue"].SetOrdinal(4);
            table.Columns["FinalAUMMarketValue"].SetOrdinal(5);
            DataRow dr = table.Rows[0];




            Random rand = new Random();
            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


            //  iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);//10,10
            //  iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
            ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingrpt.pdf";
            //String ls = path;
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

            string HHName = null;

            foreach (DataRow row in dataa.Rows)
            {
                HHName = row["BillingName"].ToString();
                if (HHName != "")
                    break;
            }


            //string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
            //if (strTitle != "")
            //    HHName = strTitle;

            string strheader = "BILLING WORKSHEET";
            string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";
            string dDate = txtAUMDate.Text;
            string[] date = dDate.Split(new char[] { '/' });
            string day = null, Month = null, year = null;
            if (dDate != "")
            {
                if (date[0].Length == 1)
                    Month = "0" + date[0];
                else
                    Month = date[0];

                if (date[1].Length == 1)
                    day = "0" + date[1];
                else
                    day = date[1];

                if (date[2].Length == 2)
                    year = "20" + date[2];
                else
                    year = date[2];


                dDate = Month + "/" + day + "/" + year;

            }


            DateTime AUMdate = DateTime.ParseExact(dDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            string vAUmdate = AUMdate.ToString("MMMM dd, yyyy");
            //   DateTime asofDT = Convert.ToDateTime("12/31/2015");
            //   string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);


            iTextSharp.text.Table loTable = new iTextSharp.text.Table(6, table.Rows.Count);   // 2 rows, 2 columns           
            // lsTotalNumberofColumns = "9";
            iTextSharp.text.Cell loCell = new Cell();


            #region Table Style
            //  int[] headerwidths9 = { 15, 23, 24, 10, 10, 10 };
            //  int[] headerwidths9 = { 15, 42, 10, 10, 10, 10 };
            //  int[] headerwidths9 = { 20, 34, 10, 11, 11, 11 };
            int[] headerwidths9 = { 17, 31, 10, 13, 13, 13 };
            loTable.SetWidths(headerwidths9);
            loTable.Width = 100;
            loTable.Border = 0;
            loTable.Cellspacing = 0;
            loTable.Cellpadding = 3;
            loTable.Locked = false;

            #endregion

            iTextSharp.text.Chunk lochunk = new Chunk();
            iTextSharp.text.Chunk lochunknew = new Chunk();
            Paragraph para = new Paragraph();
            int rowsize = table.Rows.Count;
            int colsize = 6;

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
                liPageSize = 37;
                liTotalPage = liTotalPage + 1;
            }

            for (int i = 0; i < rowsize; i++)
            {

                string BillExcepTyp = Convert.ToString(table.Rows[i]["ColourBillingExceptionType"]);
                string AUMExcepTyp = Convert.ToString(table.Rows[i]["colourAUMExceptionType"]);
                string BillFeeTyp = Convert.ToString(table.Rows[i]["BillingFeeExceptionId"]);

                if (i % liPageSize == 0)
                {
                    document.Add(loTable);

                    if (i != 0)
                    {
                        liCurrentPage = liCurrentPage + 1;
                        //   document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                        document.NewPage();
                        //SetTotalPageCount("Short Term Performance");
                    }
                    loTable = new iTextSharp.text.Table(6, table.Rows.Count);
                    //  int[] headerwidths = { 27, 9, 9, 9, 12, 12, 12, 12, 10 };
                    loTable.SetWidths(headerwidths9);
                    loTable.Width = 100;

                    loTable.Border = 0;
                    loTable.Cellspacing = 0;
                    loTable.Cellpadding = 3;
                    loTable.Locked = false;

                    lochunk = new Chunk(HHName, setFontsAllFrutiger(14, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Colspan = 6;
                    loCell.HorizontalAlignment = 1;


                    lochunk = new Chunk("\n" + strheader, setFontsAllFrutiger(10, 0, 0));
                    loCell.Add(lochunk);

                    // lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                    // loCell.Add(lochunk);

                    lochunk = new Chunk("\n " + vAUmdate + " \n", setFontsAllFrutiger(10, 0, 1));
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loTable.AddCell(loCell);



                    for (int k = 0; k < colsize; k++)
                    {
                        // string ColHeader = Convert.ToString(table.Columns[k].ColumnName);
                        // if (ColHeader == "PosAssetClassName")
                        //     ColHeader = "Asset Class (Position)";
                        // //if (ColHeader.Contains("SecurityName"))
                        // //    ColHeader = "Security Name";
                        //// if (ColHeader.Contains("AccountLegalEntityName"))   
                        // if (ColHeader.Contains("PdfAccountName")) 
                        //     ColHeader = "Legal Entity (Account)";
                        // if (ColHeader.Contains("PortCode"))
                        //     ColHeader = "Port Code";
                        // if (ColHeader.Contains("ssi_BillingMarketValue"))
                        //     ColHeader = "Total";
                        // if (ColHeader.Contains("FinalBillingMarketValue"))
                        //     ColHeader = "Billing";
                        // if (ColHeader.Contains("FinalAUMMarketValue"))
                        //     ColHeader = "AUM";
                        string ColHeader = Convert.ToString(table.Columns[k].ColumnName);
                        if (ColHeader == "PosAssetClassName")
                            ColHeader = "Asset";
                        //if (ColHeader.Contains("SecurityName"))
                        //    ColHeader = "Security Name";
                        if (ColHeader.Contains("PdfAccountName"))
                            ColHeader = "Legal Entity (Custodian Account)";
                        if (ColHeader.Contains("PortCode"))
                            ColHeader = "Port Code";
                        if (ColHeader.Contains("ssi_BillingMarketValue"))
                            ColHeader = "Total";
                        if (ColHeader.Contains("FinalBillingMarketValue"))
                            ColHeader = "Billing";
                        if (ColHeader.Contains("FinalAUMMarketValue"))
                            ColHeader = "AUM";

                        lochunk = new Chunk(ColHeader, setFontsAll(8, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        // loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        // loCell.BackgroundColor = new iTextSharp.text.Color(183, 221, 232);
                        loCell.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        //if (k != 0)
                        //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        //else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }

                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                    //  png.SetAbsolutePosition(45, 557);//540
                    png.SetAbsolutePosition(43, 800);// Portrait Logo
                    png.ScalePercent(10);
                    document.Add(png);
                }
                for (int j = 0; j < colsize; j++)
                {
                    // string cellBackgroundColor = Convert.ToString(table.Rows[i]["ColourCode"]);
                    string ColValue = Convert.ToString(table.Rows[i][j]);
                    loCell = new iTextSharp.text.Cell();


                    #region Not used
                    //if (ColValue == "Strategy Benchmark")
                    //    ColValue = "Weighted Benchmark";
                    //if (ColValue == "Gresham Advised Values")
                    //    ColValue = "Gresham Advised";

                    //if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True" || Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "" || i == 0 || i == 1)
                    //{
                    //    if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                    //    {
                    //        if (ColValue.Contains("-"))
                    //        {
                    //            ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                    //            ColValue = ColValue.Replace("(", "($");
                    //        }
                    //        else
                    //            ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                    //        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //    else if (j == 8 && ColValue != "")
                    //    {
                    //        ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //        lochunk = new Chunk(ColValue + "%", setFontsAll(7, 1, 0));
                    //    }
                    //    else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                    //    {
                    //        if (ColValue != "N/A")
                    //            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //    else
                    //    {
                    //        if ((i == 0 || i == 1) && j == 0)
                    //        {
                    //            string[] str = ColValue.Split('(');

                    //            lochunk = new Chunk(str[0], setFontsAll(7, 1, 0));
                    //            if (str.Length > 1)
                    //                lochunknew = new Chunk("\n(" + str[1], setFontsAll(6, 1, 0));
                    //        }
                    //        else if (j == 0 && cellBackgroundColor != "")
                    //        {
                    //            lochunk = new Chunk("   " + ColValue, setFontsAll(7, 1, 0));
                    //        }
                    //        else if (j == 0 && (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True"))
                    //        {
                    //            lochunk = new Chunk("       " + ColValue, setFontsAll(7, 1, 0));
                    //        }
                    //        else
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //}
                    //else
                    //{
                    //    if (j == 0) //component
                    #endregion
                    if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True")
                    {
                        //if (j == 0 && ColValue != "")
                        //    ColValue = ColValue + " Total";

                        if ((j == 3 || j == 4 || j == 5) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = currencyFormat(ColValue);
                                // ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                //  ColValue = ColValue.Replace("(", "($");
                            }
                            else
                                ColValue = currencyFormat(ColValue);
                            //ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;

                        }


                        if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True" || (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))
                        {
                            //loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                            //   loCell.BackgroundColor = new iTextSharp.text.Color(183, 221, 232); 
                            loCell.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        }
                        //if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        //{
                        //    // loCell.BackgroundColor = new iTextSharp.text.Color(204, 206, 219);
                        //    loCell.BackgroundColor = new iTextSharp.text.Color(red1, green1, blue1);
                        //}

                        if (Convert.ToString(table.Rows[i]["DiffFlg"]) == "1")
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(red3, green3, blue3);
                        }





                        // lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        #region Color on PDF
                        if (Convert.ToString(table.Rows[i]["DiffFlg"]) == "1")
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(red3, green3, blue3);
                        }

                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            // lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }
                        else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            para = new Paragraph();
                            //   para.IndentationLeft = 20;
                            para.SpacingBefore = 0;
                            para.SpacingAfter = 0;
                            para.Leading = 6F;
                            para.Add(lochunk);
                        }
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True")
                        {

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        #endregion

                    }
                    else
                    {
                        if ((j == 3 || j == 4 || j == 5) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = currencyFormat(ColValue);
                                //ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                //ColValue = ColValue.Replace("(", "($");
                            }
                            else
                                ColValue = currencyFormat(ColValue);
                            //ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));


                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        }

                        if (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")
                        {
                            if (j == 0)
                                ColValue = "TOTALS";

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));

                        }
                        else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            para = new Paragraph();
                            para.IndentationLeft = 20;
                            para.Leading = 6F;
                            // para.SpacingBefore = 0;
                            // para.SpacingAfter = 0;

                            para.Add(lochunk);
                        }
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True")
                        {

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));

                        # region Color on PDF

                        if (Convert.ToString(table.Rows[i]["DiffFlg"]) == "1")
                        {
                            loCell.BackgroundColor = new iTextSharp.text.Color(red3, green3, blue3);
                        }
                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != "") && Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }
                        else if (j == 5 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != "") && Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            //lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            // lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));

                        }
                        //else
                        //{
                        //    lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        //}
                        #endregion
                    }
                    //    else
                    //    {
                    //        if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                    //        {
                    //            if (ColValue.Contains("-"))
                    //            {
                    //                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                    //                ColValue = ColValue.Replace("(", "($");
                    //            }
                    //            else
                    //                ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //        else if (j == 8 && ColValue != "")
                    //        {
                    //            lochunk = new Chunk(Convert.ToDecimal(ColValue) + "%", setFontsAll(7, 0, 0));
                    //        }
                    //        else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                    //        {
                    //            if (ColValue != "N/A")
                    //                ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //        else
                    //        {
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //    }
                    //}


                    if (Convert.ToString(table.Rows[i]["ColourBillingExceptionType"]) == "5" && j == 4)
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(Redred, Redgreen, Redblue);
                    }
                    if (Convert.ToString(table.Rows[i]["ColourAUMExceptionType"]) == "5" && j == 5)
                    {
                        loCell.BackgroundColor = new iTextSharp.text.Color(Redred, Redgreen, Redblue);
                    }

                    loCell.BorderColor = Color.LIGHT_GRAY;
                    if (ColValue == "TOTALS")
                    {
                        loCell.Add(lochunk);
                    }
                    else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                    {

                        loCell.Add(para);
                    }

                    else
                    {

                        loCell.Add(lochunk);
                    }

                    //loCell.Add(lochunk);
                    //// if ((i == 0 || i == 1) && j == 0)
                    ////   loCell.Add(lochunknew);
                    //loCell.Border = 0;

                    //if (i == 0 || i == 1 || i == 4)
                    //    loCell.Leading = 8F;
                    //else
                    //{
                    //    //  if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "True" && Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "")
                    //    //  loCell.Leading = 0F;
                    //    // else
                    //    // loCell.Leading = 4F;
                    //}
                    loCell.Leading = 6F;
                    loCell.VerticalAlignment = 5;
                    //if (j != 0)
                    //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    //if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Gresham Advised Values")
                    //    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                    //else if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")
                    //    loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;

                    ////if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True" || (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))
                    ////{
                    ////    //loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                    //// //   loCell.BackgroundColor = new iTextSharp.text.Color(183, 221, 232); 
                    ////    loCell.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                    ////}
                    if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                    {
                        // loCell.BackgroundColor = new iTextSharp.text.Color(204, 206, 219);
                        loCell.BackgroundColor = new iTextSharp.text.Color(red1, green1, blue1);
                    }


                    //if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components")
                    //{
                    //    loCell.BorderColorBottom = iTextSharp.text.Color.WHITE;
                    //    loCell.BorderWidthBottom = 1.5f;
                    //}



                    //  if ((Convert.ToString(table.Rows[i]["ColourCode"]) != "" || (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")) && j != 0)
                    //  { 
                    //  }
                    //  else
                    // {
                    loTable.AddCell(loCell);
                    //  }
                }

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
                    liCurrentPage = liCurrentPage + 1;
                    //PdfPTable TabFooter = addFooter(lsDateTime, true, lsFooterText, false, 0, 0);
                    //TabFooter.WidthPercentage = 100f;
                    ////  TabFooter.TotalWidth = 100f;
                    //TabFooter.TotalWidth = 775;

                    //TabFooter.WriteSelectedRows(0, 4, 30, 47, writer.DirectContent);  --> Original Values
                    //TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
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
                try
                {

                    //FileInfo loFile = new FileInfo(ls);
                    //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                    //Response.Write("<script>");
                    string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";

                    //SourceFileArray[num] =  ls;

                    //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    ////  Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                    //Response.Write("</script>");
                }
                catch (Exception ex)
                {
                    Response.Write(ex.ToString());
                }

            }
            //  return fsFinalLocation.Replace(".xls", ".pdf");

        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
        }
        return ls;
    }
    public void GenratePortfolioOld()
    {
        #region blue color

        int red = 183;
        int green = 221;
        int blue = 232;

        #endregion

        #region grey color

        int red1 = 233;
        int green1 = 237;
        int blue1 = 241;

        #endregion

        #region yellow color

        int red2 = 252;
        int green2 = 252;
        int blue2 = 181;

        #endregion

        #region Green
        int red3 = 198;
        int green3 = 239;
        int blue3 = 206;

        #endregion

        #region colorRed
        int Redred = 255;
        int Redgreen = 83;
        int Redblue = 83;
        #endregion

        try
        {
            // int liPageSize = 26;  //--> Original Value
            int liPageSize = 37;

            DataSet newdataset;
            DB clsDB = new DB();
            newdataset = null;
            //int liPageSize = 18;
            //int liCurrentPage = 0;
            //lstAssetClass.s

            string billingVal = ddlBillFor.SelectedValue.ToString();
            string dDate = txtAUMDate.Text;
            string[] date = dDate.Split(new char[] { '/' });
            string day = null, Month = null;
            if (dDate != "")
            {
                if (date[0].Length == 1)
                    Month = "0" + date[0];
                else
                    Month = date[0];

                if (date[1].Length == 1)
                    day = "0" + date[1];
                else
                    day = date[1];

                dDate = Month + "/" + day + "/" + date[2];
                txtAUMDate.Text = dDate;
            }

            DateTime AUMdate = DateTime.ParseExact(txtAUMDate.Text, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            //  String lsSQL = "EXEC SP_R_BILLING @BillingForUUID = '7E1DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'";

            string lsSQL = "EXEC  SP_R_BILLING @ReportFlg = 1,@PdfFlg=1,  @BillingForUUID = '" + billingVal + "',@AsOfDate = '" + dDate + "'";

            newdataset = clsDB.getDataSet(lsSQL);

            int DSCount = newdataset.Tables[0].Rows.Count;
            DataTable table = newdataset.Tables[0].Copy();
            table.Columns["PosAssetClassName"].SetOrdinal(0);
            // table.Columns["SecurityName"].SetOrdinal(1);
            //table.Columns["AccountLegalEntityName"].SetOrdinal(2);
            table.Columns["PdfAccountName"].SetOrdinal(1);
            table.Columns["PortCode"].SetOrdinal(2);
            table.Columns["ssi_BillingMarketValue"].SetOrdinal(3);
            table.Columns["FinalBillingMarketValue"].SetOrdinal(4);
            table.Columns["FinalAUMMarketValue"].SetOrdinal(5);

            Random rand = new Random();
            //   string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();
            string strGUID = ddlBillFor.SelectedItem.Text + " " + AUMdate.ToString("yyyy-MMdd");
            //  DestinationPath = ddlBillFor.SelectedItem.Text + " " + AUMdate.ToString("yyyy-MMdd") + ".pdf";

            //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);
            //  iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
            String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingrpt.pdf";

            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".PDF";
            if (File.Exists(fsFinalLocation))
            {
                File.Delete(fsFinalLocation);

            }
            string HHName = ddlBillFor.SelectedItem.Text;
            //string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
            //if (strTitle != "")
            //    HHName = strTitle;

            string strheader = "BILLING WORKSHEET";
            string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";

            string vAUmdate = AUMdate.ToString(" MMMM dd, yyyy");
            //   DateTime asofDT = Convert.ToDateTime("12/31/2015");
            //   string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);
            //   dd MMMM ,yyyy

            iTextSharp.text.Table loTable = new iTextSharp.text.Table(6, table.Rows.Count);   // 2 rows, 2 columns           
            // lsTotalNumberofColumns = "9";
            iTextSharp.text.Cell loCell = new Cell();


            #region Table Style
            int[] headerwidths9 = { 17, 31, 10, 13, 13, 13 };
            loTable.SetWidths(headerwidths9);
            loTable.Width = 100;
            loTable.Border = 0;
            loTable.Cellspacing = 0;
            loTable.Cellpadding = 3;
            loTable.Locked = false;

            #endregion

            iTextSharp.text.Chunk lochunk = new Chunk();
            iTextSharp.text.Chunk lochunknew = new Chunk();
            Paragraph para = new Paragraph();
            int rowsize = table.Rows.Count;
            int colsize = 6;

            String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
            int liTotalPage = (table.Rows.Count / liPageSize);
            int liCurrentPage = 0;

            if (table.Rows.Count % liPageSize != 0)
            {
                liTotalPage = liTotalPage + 1;
            }
            else
            {
                // liPageSize = 26; //--> Original Value
                liPageSize = 37;
                liTotalPage = liTotalPage + 1;
            }

            for (int i = 0; i < rowsize; i++)
            {
                //string BillExcepTyp = Convert.ToString(table.Rows[i]["BillingExceptionType"]);
                //string AUMExcepTyp = Convert.ToString(table.Rows[i]["AUMExceptionType"]);

                string BillExcepTyp = Convert.ToString(table.Rows[i]["ColourBillingExceptionType"]);
                string AUMExcepTyp = Convert.ToString(table.Rows[i]["colourAUMExceptionType"]);
                string BillFeeTyp = Convert.ToString(table.Rows[i]["BillingFeeExceptionId"]);
                if (i % liPageSize == 0)
                {
                    document.Add(loTable);

                    if (i != 0)
                    {
                        liCurrentPage = liCurrentPage + 1;
                        //   document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                        document.NewPage();
                        //SetTotalPageCount("Short Term Performance");
                    }
                    loTable = new iTextSharp.text.Table(6, table.Rows.Count);
                    //  int[] headerwidths = { 27, 9, 9, 9, 12, 12, 12, 12, 10 };
                    loTable.SetWidths(headerwidths9);
                    loTable.Width = 100;

                    loTable.Border = 0;
                    loTable.Cellspacing = 0;
                    loTable.Cellpadding = 3;
                    loTable.Locked = false;

                    lochunk = new Chunk(HHName, setFontsAllFrutiger(14, 1, 0));
                    loCell = new Cell();
                    loCell.Add(lochunk);
                    loCell.Colspan = 6;
                    loCell.HorizontalAlignment = 1;


                    lochunk = new Chunk("\n" + strheader, setFontsAllFrutiger(10, 0, 0));
                    loCell.Add(lochunk);

                    // lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                    // loCell.Add(lochunk);

                    lochunk = new Chunk("\n " + vAUmdate + " \n", setFontsAllFrutiger(10, 0, 1));
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loTable.AddCell(loCell);



                    // Add table columns
                    for (int k = 0; k < colsize; k++)
                    {
                        string ColHeader = Convert.ToString(table.Columns[k].ColumnName);
                        if (ColHeader == "PosAssetClassName")
                            ColHeader = "Asset";
                        //if (ColHeader.Contains("SecurityName"))
                        //    ColHeader = "Security Name";
                        if (ColHeader.Contains("PdfAccountName"))
                            ColHeader = "Legal Entity (Custodian Account)";
                        if (ColHeader.Contains("PortCode"))
                            ColHeader = "Port Code";
                        if (ColHeader.Contains("ssi_BillingMarketValue"))
                            ColHeader = "Total";
                        if (ColHeader.Contains("FinalBillingMarketValue"))
                            ColHeader = "Billing";
                        if (ColHeader.Contains("FinalAUMMarketValue"))
                            ColHeader = "AUM";

                        lochunk = new Chunk(ColHeader, setFontsAll(8, 1, 0));
                        loCell = new iTextSharp.text.Cell();
                        loCell.Add(lochunk);
                        loCell.Border = 0;
                        loCell.Leading = 11F;
                        //loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        loCell.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        //if (k != 0)
                        //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        //else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }

                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                    //  png.SetAbsolutePosition(45, 557);//540
                    png.SetAbsolutePosition(43, 800); // potrait logo
                    png.ScalePercent(10);
                    document.Add(png);
                }

                for (int j = 0; j < colsize; j++)
                {
                    // string cellBackgroundColor = Convert.ToString(table.Rows[i]["ColourCode"]);
                    string ColValue = Convert.ToString(table.Rows[i][j]);
                    iTextSharp.text.Cell loCell1 = new iTextSharp.text.Cell();

                    #region not used
                    //loCell = new iTextSharp.text.Cell();
                    //if (ColValue == "Strategy Benchmark")
                    //    ColValue = "Weighted Benchmark";
                    //if (ColValue == "Gresham Advised Values")
                    //    ColValue = "Gresham Advised";

                    //if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True" || Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "" || i == 0 || i == 1)
                    //{
                    //    if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                    //    {
                    //        if (ColValue.Contains("-"))
                    //        {
                    //            ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                    //            ColValue = ColValue.Replace("(", "($");
                    //        }
                    //        else
                    //            ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                    //        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //    else if (j == 8 && ColValue != "")
                    //    {
                    //        ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //        lochunk = new Chunk(ColValue + "%", setFontsAll(7, 1, 0));
                    //    }
                    //    else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                    //    {
                    //        if (ColValue != "N/A")
                    //            ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //        lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //    else
                    //    {
                    //        if ((i == 0 || i == 1) && j == 0)
                    //        {
                    //            string[] str = ColValue.Split('(');

                    //            lochunk = new Chunk(str[0], setFontsAll(7, 1, 0));
                    //            if (str.Length > 1)
                    //                lochunknew = new Chunk("\n(" + str[1], setFontsAll(6, 1, 0));
                    //        }
                    //        else if (j == 0 && cellBackgroundColor != "")
                    //        {
                    //            lochunk = new Chunk("   " + ColValue, setFontsAll(7, 1, 0));
                    //        }
                    //        else if (j == 0 && (Convert.ToString(table.Rows[i]["AssetClassFlg"]) == "True"))
                    //        {
                    //            lochunk = new Chunk("       " + ColValue, setFontsAll(7, 1, 0));
                    //        }
                    //        else
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                    //    }
                    //}
                    //else
                    //{
                    //    if (j == 0) //component
                    #endregion


                    if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True") // if it is total asset
                    {
                        //if (j == 0 && ColValue != "")
                        //    ColValue = ColValue + " Total";

                        if ((j == 3 || j == 4 || j == 5) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = currencyFormat(ColValue);
                                //  ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                //ColValue = String.Format("{0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));
                                //ColValue = ColValue.Replace("(", "($");

                            }
                            //else if (Convert.ToDecimal(ColValue) == 0)
                            //{
                            //    ColValue = "$0.00";
                            //}
                            else
                            {
                                ColValue = currencyFormat(ColValue);
                                // ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                // ColValue = String.Format("${0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));

                            }
                            loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        }

                        if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True" || (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))
                        {
                            //loCell1.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                        }




                        //lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        # region Color on PDF


                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            // lochunk.SetBackground(Color.YELLOW);
                            //lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }
                        //----------------------------------------------------------------------------------------------------------
                        else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            para = new Paragraph();
                            //   para.IndentationLeft = 20;
                            para.SpacingBefore = 0;
                            para.SpacingAfter = 0;
                            para.Leading = 6F;
                            para.Add(lochunk);
                        }
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True")
                        {

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        }
                        #endregion
                    }
                    else
                    {
                        if (Convert.ToString(table.Rows[i]["DiffFlg"]) == "1")
                        {
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red3, green3, blue3);
                        }
                        if ((j == 3 || j == 4 || j == 5) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = currencyFormat(ColValue);
                                //  ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                //ColValue = String.Format("{0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));
                                //ColValue = ColValue.Replace("(", "($");

                            }
                            //else if (Convert.ToDecimal(ColValue) == 0)
                            //{
                            //    ColValue = "$0.00";
                            //}
                            else
                            {
                                ColValue = currencyFormat(ColValue);
                                // ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                // ColValue = String.Format("${0:#,###0.##;(#,###0.##)}", Convert.ToDecimal(ColValue));

                            }

                            loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        }

                        if (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")
                        {
                            if (j == 0)
                                ColValue = "TOTALS";

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));


                        }
                        //-------------------------------------------------------------------------------------------
                        else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            para = new Paragraph();
                            para.IndentationLeft = 20;
                            para.Leading = 6F;
                            // para.SpacingBefore = 0;
                            // para.SpacingAfter = 0;

                            para.Add(lochunk);
                        }

                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True")
                        {

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        }

                        # region Color on PDF
                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != "") && Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }
                        else if (j == 5 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != "") && Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            //  lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                            // lochunk.SetBackground(Color.YELLOW);
                            // lochunk.SetBackground(new iTextSharp.text.Color(red2, green2, blue2));
                            loCell1.BackgroundColor = new iTextSharp.text.Color(red2, green2, blue2);

                        }

                        if (Convert.ToString(table.Rows[i]["ColourBillingExceptionType"]) == "5" && j == 4)
                        {
                            loCell1.BackgroundColor = new iTextSharp.text.Color(Redred, Redgreen, Redblue);
                        }
                        if (Convert.ToString(table.Rows[i]["ColourAUMExceptionType"]) == "5" && j == 5)
                        {
                            loCell1.BackgroundColor = new iTextSharp.text.Color(Redred, Redgreen, Redblue);
                        }


                       //else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                        //{
                        //    lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                        //    para = new Paragraph();
                        //    para.IndentationLeft = 20;
                        //    para.SpacingBefore = 0;
                        //    para.SpacingAfter = 0;

                                               //    para.Add(lochunk);
                        //}
                        else if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else if (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));

                        }


                        #endregion

                    }
                    //    else
                    //    {
                    //        if ((j == 4 || j == 5 || j == 6 || j == 7) && ColValue != "")
                    //        {
                    //            if (ColValue.Contains("-"))
                    //            {
                    //                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                    //                ColValue = ColValue.Replace("(", "($");
                    //            }
                    //            else
                    //                ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));

                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //        else if (j == 8 && ColValue != "")
                    //        {
                    //            lochunk = new Chunk(Convert.ToDecimal(ColValue) + "%", setFontsAll(7, 0, 0));
                    //        }
                    //        else if ((j == 1 || j == 2 || j == 3) && ColValue != "")
                    //        {
                    //            if (ColValue != "N/A")
                    //                ColValue = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(ColValue));
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //        else
                    //        {
                    //            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));
                    //        }
                    //    }
                    //}

                    // loCell1 = new iTextSharp.text.Cell();
                    // loCell1.BorderWidth = 1;

                    if (Convert.ToString(table.Rows[i]["ColourBillingExceptionType"]) == "5" && j == 4)
                    {
                        loCell1.BackgroundColor = new iTextSharp.text.Color(Redred, Redgreen, Redblue);
                    }
                    if (Convert.ToString(table.Rows[i]["ColourAUMExceptionType"]) == "5" && j == 5)
                    {
                        loCell1.BackgroundColor = new iTextSharp.text.Color(Redred, Redgreen, Redblue);
                    }

                    loCell1.BorderColor = Color.LIGHT_GRAY;
                    if (ColValue == "TOTALS")
                    {
                        loCell1.Add(lochunk);
                    }
                    else if (j == 0 && Convert.ToString(table.Rows[i]["_AssetHeader"]) != "1")
                    {

                        loCell1.Add(para);
                    }

                    else
                    {

                        loCell1.Add(lochunk);
                    }
                    // loCell1.Add(lochunk);
                    // if ((i == 0 || i == 1) && j == 0)
                    //   loCell.Add(lochunknew);
                    // loCell1.Border = 0;

                    //if (i == 0 || i == 1 || i == 4)
                    //    loCell.Leading = 8F;
                    //else
                    //{
                    //    //  if (Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "True" && Convert.ToString(table.Rows[i]["AssetClassFlg"]) != "")
                    //    //  loCell.Leading = 0F;
                    //    // else
                    //    // loCell.Leading = 4F;
                    //}

                    loCell1.Leading = 6F;
                    loCell1.VerticalAlignment = 5;

                    //if (j != 0)
                    //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    //if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Gresham Advised Values")
                    //    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                    //else if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")
                    //    loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;

                    //if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True" || (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))
                    //{
                    //    //loCell1.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                    //    loCell1.BackgroundColor = new iTextSharp.text.Color(red, green, blue);
                    //}
                    if (Convert.ToString(table.Rows[i]["_AssetHeader"]) == "1")
                    {
                        loCell1.BackgroundColor = new iTextSharp.text.Color(red1, green1, blue1);
                    }



                    //if (Convert.ToString(table.Rows[i]["DiffFlg"]) == "1")
                    //{
                    //    loCell1.BackgroundColor = new iTextSharp.text.Color(red3, green3, blue3);
                    //}

                    //if (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components")
                    //{
                    //    loCell.BorderColorBottom = iTextSharp.text.Color.WHITE;
                    //    loCell.BorderWidthBottom = 1.5f;
                    //}



                    //  if ((Convert.ToString(table.Rows[i]["ColourCode"]) != "" || (Convert.ToString(table.Rows[i]["Overall Performance"]) == "Marketable Strategies Components" || Convert.ToString(table.Rows[i]["Overall Performance"]) == "Private Strategies Components")) && j != 0)
                    //  { 
                    //  }
                    //  else
                    // {
                    //  loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                    loTable.AddCell(loCell1);

                    //  }
                }

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
                    liCurrentPage = liCurrentPage + 1;
                    //PdfPTable TabFooter = addFooter(lsDateTime, true, lsFooterText, false, 0, 0);
                    //TabFooter.WidthPercentage = 100f;
                    ////  TabFooter.TotalWidth = 100f;
                    //TabFooter.TotalWidth = 775;

                    //TabFooter.WriteSelectedRows(0, 4, 30, 47, writer.DirectContent);  --> Original Values
                    //TabFooter.WriteSelectedRows(0, 4, 30, 30, writer.DirectContent);
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
                try
                {
                    fsFinalLocation.Replace(".xls", ".pdf");
                    if (File.Exists(fsFinalLocation))
                    {
                        File.Delete(fsFinalLocation);
                    }

                    FileInfo loFile = new FileInfo(ls);
                    loFile.MoveTo(fsFinalLocation);
                   strGUID= strGUID.Replace("'", "%27");
                    Response.Write("<script>");
                    string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
                    Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    //  Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                    Response.Write("</script>");
                }
                catch (Exception ex)
                {
                    Response.Write(ex.ToString());
                }

            }
            //  return fsFinalLocation.Replace(".xls", ".pdf");

        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
        }


    }
    public void MergeFiles(string destinationFile, string[] sourceFiles)
    {
        try
        {


            MergeNew(destinationFile, sourceFiles);
            return;

            int f = 0;
            // we create a reader for a certain document
            PdfReader reader = new PdfReader(sourceFiles[f]);
            // we retrieve the total number of pages
            int n = reader.NumberOfPages;
            //Console.WriteLine("There are " + n + " pages in the original file.");
            // step 1: creation of a document-object
            Document document = new Document(reader.GetPageSizeWithRotation(1));
            // step 2: we create a writer that listens to the document
            //FileInfo file = new FileInfo();
            //file.FullName = "e:\\repots\\1.txt";
            //file.Create();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(destinationFile, FileMode.Create));
            // step 3: we open the document
            document.Open();
            PdfContentByte cb = writer.DirectContent;
            PdfImportedPage page;
            int rotation;
            // step 4: we add content
            while (f < sourceFiles.Length)
            {
                int i = 0;
                while (i < n)
                {
                    i++;
                    document.SetPageSize(reader.GetPageSizeWithRotation(i));
                    document.NewPage();
                    page = writer.GetImportedPage(reader, i);
                    rotation = reader.GetPageRotation(i);
                    if (rotation == 90 || rotation == 270)
                    {
                        cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
                    }
                    else
                    {
                        cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                    }
                    //Console.WriteLine("Processed page " + i);
                }
                f++;
                if (f < sourceFiles.Length)
                {
                    if (sourceFiles[f] != null && Convert.ToString(sourceFiles[f]) != "")
                    {
                        reader = new PdfReader(sourceFiles[f]);
                        // we retrieve the total number of pages
                        n = reader.NumberOfPages;
                        //Console.WriteLine("There are " + n + " pages in the original file.");
                    }
                    else
                    {
                        //f++;
                        n = 0;
                    }
                }
            }
            // step 5: we close the document
            document.Close();
        }
        catch (Exception e)
        {
            string strOb = e.Message;
        }
    }
    public void MergeNew(string destinationFile, string[] lstFiles)
    {
        PdfReader reader = null;
        Document sourceDocument = null;
        PdfCopy pdfCopyProvider = null;
        PdfImportedPage importedPage;
        string outputPdfPath = destinationFile;


        sourceDocument = new Document();
        pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

        //Open the output file
        sourceDocument.Open();

        try
        {
            //Loop through the files list
            for (int f = 0; f <= lstFiles.Length - 1; f++)
            {
                if (lstFiles[f] != null)
                {
                    int pages = get_pageCcount(lstFiles[f]);

                    reader = new PdfReader(lstFiles[f]);
                    //Add pages of current file
                    for (int i = 1; i <= pages; i++)
                    {
                        importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                        pdfCopyProvider.AddPage(importedPage);
                    }

                    reader.Close();
                }
            }
            //At the end save the output file
            sourceDocument.Close();


            string HH = ddlHH.SelectedItem.Text;
            string Filename = HH + ".pdf";

            //Response.ContentType = "Application/pdf";
            //Response.AppendHeader("Content-Disposition", "attachment; filename=" + Filename);
            //Response.TransmitFile(destinationFile);
            //  Response.End();

            string filename = Path.GetFileNameWithoutExtension(destinationFile);

            FileInfo loFile = new FileInfo(destinationFile);
            filename = filename.Replace("'", "%27");
            loFile.MoveTo(destinationFile.Replace(".xls", ".pdf"));
            Response.Write("<script>");
            string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + filename + ".pdf";
            //  Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
            Response.Write("window.open('" + lsFileNamforFinalXls + "', '_newtab')");
            Response.Write("</script>");

            //FileInfo loFile = new FileInfo(destinationFile);
            //loFile.MoveTo(filename.Replace(".xls", ".pdf"));
            //Response.Write("<script>");
            //string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + filename;
            //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");

            //lblMessage.Text = "Records Updated Successfully";
        }
        catch (Exception ex)
        {
            throw ex;
        }


    }

    private int get_pageCcount(string file)
    {
        using (StreamReader sr = new StreamReader(System.IO.File.OpenRead(file)))
        {
            Regex regex = new Regex(@"/Type\s*/Page[^s]");
            MatchCollection matches = regex.Matches(sr.ReadToEnd());

            return matches.Count;
        }
    }

    public string currencyFormat(string Value)
    {
        string value = Value.Replace(",", "").Replace("$", "").Replace("%", "").Replace("(", "-").Replace(")", "");

        decimal ul = 0;
        if (value == "")
            ul = 0;//text.Text = "";
        else
            ul = Convert.ToDecimal(value);

        value = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);

        return value;
    }

    protected void BindAdvisors()//Fucntion To Bind Advisor DropDown
    {
        try
        {
            /*populate Advisor DropDown*/
            string sqlstr = "[SP_S_BILLING_ADVISORS] ";//fetch data from storedprocedure
            BindDropdown(ddlAdvisor, sqlstr, "SecondaryOwnerName", "SecondaryOwnerUUID");//fucntion bind all data
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
        }
    }

    protected void BindHouseHold()//Function To Bind HouseHold DropDown
    {
        try
        {
            /*check for value or null*/
            object AdvisorValue = Convert.ToString(ddlAdvisor.SelectedValue) == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";

            /*pass the advisior value to fetch Household*/
            string sqlstr = "[SP_S_BILLING_HOUSEHOLD] @SecondaryOwnerUUID = " + AdvisorValue + " ";
            BindDropdown(ddlHH, sqlstr, "HouseHoldName", "HouseHoldUUID");
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
        }
    }

    protected void BindBillingName()//Function TO Bind Biiling_Name DropDown
    {
        try
        {
            //check selected value equal to key or check null 
            object HouseHoldValue = Convert.ToString(ddlHH.SelectedValue) == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlHH.SelectedValue + "'";

            /* pass the HouseHold value to fetch Billing Name*/
            string sqlstr = "[SP_S_BILLING_NAME] @HouseHoldUUID = " + HouseHoldValue + " ";
            BindDropdown(ddlBillFor, sqlstr, "BillingName", "BillingUUIDAndFeeType");

        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
        }
    }

    private void BindDropdown(DropDownList ddl, string sqlstr, string TextField, string ValueField)
    {
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        //establish connection
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);//Fill Dataset 

        ddl.DataTextField = TextField;
        ddl.DataValueField = ValueField;

        ddl.DataSource = ds;
        ddl.DataBind();//bind texfield and valuefield

        // ddl.Items.Insert(0, "--SELECT--");
        // ddl.Items[0].Value = "0";
    }

    #region Common Methods
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
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanai.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanaz.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
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

    public iTextSharp.text.Font setFontsAllTimesNewRoman(int size, int bold, int italic)
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

        #region WITH NEW FONTS FROM Times NEW ROMAN
        string fontpath = HttpContext.Current.Server.MapPath(".");

        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\times.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\timesbd.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\timesi.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\timesbi.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
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



    #endregion
    protected void ddlAdvisor_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindHouseHold();
        BindBillingName();
        lbGroup.Visible = false;
    }
    protected void ddlHH_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindBillingName();
        lbGroup.Visible = false;
    }
    protected void ddlBillFor_SelectedIndexChanged1(object sender, EventArgs e)
    {
        string BillingUUID = "";
        if (ddlBillFor.SelectedValue != "" && ddlBillFor.SelectedValue != "ALL")
        {
            string strBillFor = ddlBillFor.SelectedValue;
            char[] delimiterChars = { '|' };
            string[] words = strBillFor.ToString().Split(delimiterChars);
            int len = words.Length;
            if (len > 0)
            {
                strBillFor = words[0];
                BillingUUID = words[0];
            }
        }


        object AdvisorValue = Convert.ToString(ddlAdvisor.SelectedValue) == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object ddlBillForValue = Convert.ToString(BillingUUID) == "" ? "null" : "'" + ddlBillFor.SelectedValue + "'";
        string sqlstr = "[SP_S_BILLING_HOUSEHOLD] @SecondaryOwnerUUID = " + AdvisorValue + ",@BillingUUID= " + ddlBillForValue + " ";
        DB clsDB = new DB();
        DataSet dsBilling = clsDB.getDataSet(sqlstr);
        string HHID = "";
        if (dsBilling.Tables[0].Rows.Count > 2)
        {
            ddlHH.SelectedIndex = 0;
            HHID = Convert.ToString(dsBilling.Tables[0].Rows[0]["HouseHoldUUID"]);
        }
        else
        {
            ddlHH.SelectedValue = Convert.ToString(dsBilling.Tables[0].Rows[1]["HouseHoldUUID"]);
            HHID = Convert.ToString(dsBilling.Tables[0].Rows[1]["HouseHoldUUID"]);
        }
        lblMessage.Text = "";

        string SQLQuery = "SP_S_HOUSEHOLD_BILLING @HouseholdID='" + HHID + "'";
        DataSet dsBillings = clsDB.getDataSet(SQLQuery);

        if (dsBillings.Tables[0].Rows.Count > 1)
        {
            lbGroup.Visible = true;
            BindListBox(dsBillings.Tables[0]);
        }
        else
        {
            lbGroup.Visible = false;
        }

    }
    protected void txtAUMDate_TextChanged(object sender, EventArgs e)
    {

    }
    protected void BindListBox(DataTable dtData)
    {
        lbGroup.ClearSelection();
        lbGroup.DataSource = dtData;
        lbGroup.DataTextField = "BillingName";
        //lbGroup.DataValueField = "ProductID";
        lbGroup.DataBind();

    }
}