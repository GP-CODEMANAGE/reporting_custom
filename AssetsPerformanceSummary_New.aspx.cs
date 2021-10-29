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
using iTextSharp.text.pdf;
using System.IO;


public partial class AssetsPerformanceSummary_New : System.Web.UI.Page
{
    public string lsTotalNumberofColumns;
    GeneralMethods clsGM = new GeneralMethods();
    //public int liPageSize = 26;   --> Original Vlaue
    public int liPageSize = 30;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            BindHouseHolds();
            Bind_AssetClass();
        }
    }

    //private void BindHouseHolds()
    //{
    //    string sqlstr = string.Empty;

    //    sqlstr = "[sp_s_Get_HouseHoldName] ";
    //    BindDropdown(ddlHousehold, sqlstr, "name", "accountid");
    //}

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

    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        //string sqlstr = string.Empty;

        //sqlstr = "[sp_s_Get_GroupName_Only] @HouseHoldNameTxt='" + ddlHousehold .SelectedItem.Text+ "'";
        //BindDropdown(ddlGroup, sqlstr, "groupid", "groupname");

        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("[sp_s_Get_GroupName_Only] @HouseHoldNameTxt='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "',@MarkovType=2");
        ddlGroup.Items.Clear();
        ddlGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        GeneratePDF();
    }

    private void GeneratePDF()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        //int liPageSize = 18;
        //int liCurrentPage = 0;
        //lstAssetClass.s

        string StartDate = txtStartDate.Text == "" ? "null" : "'" + txtStartDate.Text + "'";
        //string spName = "SP_R_PERF_CALCS_NEW_UNIT_STUDY_NEW_GA";
        string spName = "SP_R_PERF_CALCS_NEW_UNIT_STUDY_NEW_GA_BASEDATA";
        if (Request.QueryString.Count > 0)
        {
            if (Request.QueryString["type"] == "new")
            {
                //spName = "SP_R_PERF_CALCS_NEW_UNIT_STUDY_NEW";
                //spName = "SP_R_PERF_CALCS_NEW_UNIT_STUDY_NEW_GA";
                spName = "SP_R_PERF_CALCS_NEW_UNIT_STUDY_NEW_GA_BASEDATA";
                //Response.Write("New");
            }
        }

        string strAssetClass = lstAssetClass.SelectedValue == "0" ? "'" + GetAllItemsTextFromListBox(lstAssetClass, false) + "'" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";
        String lsSQL = "" + spName + " @GroupName = '" + ddlGroup.SelectedItem.Text.Replace("'", "''") + "'" +
                                                        ", @AsofDate = '" + txtAsofdate.Text + "',@AssetNameTxt = " + strAssetClass + ",@PdfFlg=1" +
                                                        ",@StarDT=" + StartDate + "";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        Random rand = new Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString(); 


        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "AssetPerformanceReport.pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        string FooterText = "";//footer text is in below method

        AddFooter(document, FooterText);
        document.Open();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

        string HHName = Convert.ToString(ddlHousehold.SelectedItem.Text);
        string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
        if (strTitle != "")
            HHName = strTitle;

        string strheader = "GRESHAM ADVISED ASSETS";
        string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";

        DateTime asofDT = Convert.ToDateTime(txtAsofdate.Text);

        string AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);


        iTextSharp.text.Table loTable = new iTextSharp.text.Table(9, table.Rows.Count);   // 2 rows, 2 columns           
        lsTotalNumberofColumns = "9";
        iTextSharp.text.Cell loCell = new Cell();
        setTableProperty(loTable);
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
                    document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                    document.NewPage();
                }
                loTable = new iTextSharp.text.Table(9, table.Rows.Count);
                setTableProperty(loTable);

                lochunk = new Chunk(HHName, setFontsAllFrutiger(14, 1, 0));
                loCell = new Cell();
                loCell.Add(lochunk);
                loCell.Colspan = 9;
                loCell.HorizontalAlignment = 1;


                lochunk = new Chunk("\n" + strheader, setFontsAllFrutiger(10, 0, 0));
                loCell.Add(lochunk);

                lochunk = new Chunk("\n" + Title, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
                loCell.Add(lochunk);

                lochunk = new Chunk("\n" + AsOfDate + "\n", setFontsAllFrutiger(10, 0, 1));
                loCell.Add(lochunk);
                loCell.Border = 0;
                loTable.AddCell(loCell);

                decimal FeePercent = Convert.ToDecimal(newdataset.Tables[1].Rows[0][0]);
                FeePercent = Math.Round(FeePercent, 0);
                lochunk = new Chunk("Returns are shown net of all manager fees, but gross of Gresham’s fee, which is currently " + FeePercent + " basis points.  Gresham’s fee covers a range of interrelated investment, planning and advisory services.", setFontsAllFrutiger(7, 1, 0));
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

                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
                png.SetAbsolutePosition(45, 557);//540
                png.ScalePercent(10);
                document.Add(png);
            }
            for (int j = 0; j < colsize; j++)
            {
                string ColValue = Convert.ToString(table.Rows[i][j]);
string cellBackgroundColor = Convert.ToString(table.Rows[i]["ColourCode"]);
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
                liCurrentPage = liCurrentPage + 1;

                //Chunk lochunk4 = new Chunk("\n Qualitative Summary", setFontsAllFrutiger(9, 1, 0));
                //Chunk lochunk5 = new Chunk("\n Gresham Advised Equity strategies and Gresham Advised Marketable Equity strategies will not include fixed income at this time.", setFontsAllFrutiger(7, 0, 0));
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

    public iTextSharp.text.Font Font7Normal()
    {
        return setFontsAll(7, 0, 0);
    }

    public void setTableProperty(iTextSharp.text.Table fotable)
    {
        //int[] headerwidths = { 28, 9, 9, 9, 9, 9, 9, 9, 7 };

        setWidthsoftable(fotable);

        fotable.Width = 100;
        //if (Type == ReportType.FundMemorandum)
        //{
        //    fotable.Alignment = 0;
        //}
        //else
        //{
        //    fotable.Alignment = 1;
        //}
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
                int[] headerwidths3 = { 30, 9, 9 };
                fotable.SetWidths(headerwidths3);
                fotable.Width = 49;
                break;
            case "4":
                int[] headerwidths4 = { 45, 5, 15, 15 };
                fotable.SetWidths(headerwidths4);
                fotable.Width = 80;
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
                int[] headerwidths9 = { 27, 9, 9, 9, 14, 12, 12, 14, 10 };
                fotable.SetWidths(headerwidths9);
                fotable.Width = 100;
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
        Phrase footPhraseImg = new Phrase("Gresham Advisors, LLC | 333 W. Wacker Dr. Suite 700 | Chicago, IL 60606 | P 312.960.0200 | F 312.960.0204 | www.greshampartners.com", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
        footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
        footer.Alignment = Element.ALIGN_CENTER;
        document.Footer = footer;
    }
    public void AddFooter(iTextSharp.text.Document document, string FooterText)
    {
        Phrase footPhraseImg = new Phrase();
        footPhraseImg.Add(new Chunk("Gresham Advised Assets (GAA): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150))));
        footPhraseImg.Add(new Chunk("All Gresham advised investments except cash (includes private strategies - private real assets and private equity).", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150))));
        footPhraseImg.Add(new Chunk("\nGresham Advised Marketable Assets (Marketable GAA): ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150))));
        footPhraseImg.Add(new Chunk("All Gresham advised investments except cash and private strategies.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150))));
        footPhraseImg.Add(new Chunk("\nWeighted Benchmark: ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150))));
        footPhraseImg.Add(new Chunk("The average of the benchmark return for each asset class, weighted by that asset class' percentage of total marketable GAA each month.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150))));
        //footPhraseImg.Add(new Chunk("\nRisk Adjusted Performance: ", setFontsAll(7, 1, 0, new iTextSharp.text.Color(150, 150, 150))));
        //footPhraseImg.Add(new Chunk("The annualized return in excess of the return you would expect from your portfolio given its level of market exposure and the market's overall return.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150))));
        footPhraseImg.Add(new Chunk("\nSee notes for this illustration located in the Appendix under Index Definitions for important information.", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150))));
        footPhraseImg.Leading = 8f;
        //        string FooterText = "Gresham Advised Assets (GAA): All Gresham advised investments except cash (includes non-marketable strategies - illiquid real assets and private equity)." +
        //"\nGresham Advised Marketable Assets (Marketable GAA): All Gresham advised investments except cash and non-marketable strategies." +
        //"\nWeighted Benchmark: The average of the benchmark return for each asset class, weighted by that asset class' percentage of total marketable GAA each month." +
        //"\nRisk Adjusted Performance: The annualized return in excess of the return you would expect from your portfolio given its level of market exposure and the market's overall return." +
        //"\nSee notes for this illustration located in the Appendix under Index Definitions for important information.";

        //Phrase footPhraseImg = new Phrase(FooterText, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
        footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
        footer.Alignment = Element.ALIGN_LEFT;
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

    public void AddStyle(string selector, string styles)
    {

        //string STYLE_DEFAULT_TYPE = "style";
        //this._Styles.LoadTagStyle(selector, STYLE_DEFAULT_TYPE, styles);
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

    public string RoundUp(string lsFormatedString)
    {
        lsFormatedString = String.Format("{0:#,###0.00;(#,###0.00)}", Convert.ToDecimal(lsFormatedString));
        return lsFormatedString;
    }
}
