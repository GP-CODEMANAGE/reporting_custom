using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Reporting.WebForms;

public partial class ACR : System.Web.UI.Page
{
    int liPageSize = 19;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    int tempLineCount;
    public string lsTotalNumberofColumns;
    object year = null;
    object RespParty = null;
    object HH = null;
    object TaskType = null;
    object Status = null;

    protected void Page_Load(object sender, EventArgs e)
    {






        if (!IsPostBack)
        {
 string vContain1 = "-------Application connecting to ----" + System.Net.ServicePointManager.SecurityProtocol + "-------";
            Response.Write(vContain1);
            bindDropDowns();
        }
    }
    #region Populating the Dropdownlists
    protected void bindDropDowns()
    {
        try
        {
            /* Populating Year DropDown */
            DB clsDB = new DB();
            DataSet ds = clsDB.getDataSet("SP_S_ANNUALREVIEW_YEAR");
            ddlYear.Items.Clear();

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                ddlYear.Items.Add(new System.Web.UI.WebControls.ListItem(ds.Tables[0].Rows[i][0].ToString(), ds.Tables[0].Rows[i][0].ToString()));
            }
            if (ddlYear.Items.Count > 1)
                ddlYear.SelectedIndex = 0;
            ddlYear.Items.Insert(0, (new System.Web.UI.WebControls.ListItem("ALL", "ALL")));

            /* Populating Responsible DropDown */
            ds = clsDB.getDataSet("SP_S_ANNUALREVIEW_RESPONSIBLEPARTY");
            ddlRespParty.Items.Clear();
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                ddlRespParty.Items.Add(new System.Web.UI.WebControls.ListItem(ds.Tables[0].Rows[i][0].ToString(), ds.Tables[0].Rows[i][1].ToString()));
            }

            /* Populating Household DropDown */
            ds = clsDB.getDataSet("SP_S_ANNUALREVIEW_HOUSEHOLD");
            ddlHH.Items.Clear();
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                ddlHH.Items.Add(new System.Web.UI.WebControls.ListItem(ds.Tables[0].Rows[i][0].ToString(), ds.Tables[0].Rows[i][1].ToString()));
            }

            /* Populating Task Type DropDown */
            ds = clsDB.getDataSet("SP_S_ANNUALREVIEW_TASKTYPE");
            ddlTaskType.Items.Clear();
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                ddlTaskType.Items.Add(new System.Web.UI.WebControls.ListItem(ds.Tables[0].Rows[i][0].ToString(), ds.Tables[0].Rows[i][1].ToString()));
            }

            /* Populating Status ListBox */
            ListBoxStatuses.Items.Clear();
            GeneralMethods clsGM = new GeneralMethods();
            clsGM.getListForBindListBox(ListBoxStatuses, "SP_S_ANNUALREVIEW_STATUS", "StatusName", "StatusID");
            ListBoxStatuses.Items[0].Value = "0";
            ListBoxStatuses.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
        }
    }
    #endregion

    #region CODE NOT IN USE - Generating the PDF File using iText
    protected void btnGeneratePDF_Click(object sender, EventArgs e)
    {
        /* Generating the PDF File using the selected values */
        try
        {
            clearMesssage();
            liPageSize = 19;
            DataSet newdataset = null;
            DataTable loInsertblankRow = null;
            DB clsDB = new DB();

            String lsFooterTxt = "";// "See notes for this illustration located in the Appendix under Commitment Schedule for important information.";
            //String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);

            string lsSQL = getFinalSp();
            // Response.Write(lsSQL);
            newdataset = clsDB.getDataSet(lsSQL);
            if (newdataset.Tables.Count > 0)
            {
                if (newdataset.Tables[0].Rows.Count < 1)
                {
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "No Record found";
                    return;
                }
            }
            DataTable table = newdataset.Tables[0].Clone();
            table.AcceptChanges();
            loInsertblankRow = newdataset.Tables[0].Copy();
            //loInsertblankRow.Tables.Add(table);
            //lodataset.Tables[0].Clear();
            table.Clear();
            table = null;
            table = loInsertblankRow.Clone();

            string strGUID = DateTime.Now.ToString("MMddyyhhmmss");
            String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";

            for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Rows.Count; liBlankRow++)
            {
                if (loInsertblankRow.Rows.Count > 1)
                {
                    table.ImportRow(loInsertblankRow.Rows[liBlankRow]);
                }
            }
            table.AcceptChanges();
            DataSet lodataset = new DataSet();
            lodataset.Tables.Add(table);
            DataSet loInsertdataset = lodataset.Copy();

            /* Removing the additional columns from the dataset, which are not required in the PDF */
            //loInsertdataset.Tables[0].Columns.RemoveAt(14);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(13);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(12);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(11);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(10);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(9);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(8);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(7);
            //loInsertdataset.Tables[0].AcceptChanges();
            //loInsertdataset.Tables[0].Columns.RemoveAt(6);
            //loInsertdataset.Tables[0].AcceptChanges();

            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
            String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
            iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
            document.Open();

            lsTotalNumberofColumns = "5";//loInsertdataset.Tables[0].Columns.Count + "";
            PdfPTable PdfPtable = new PdfPTable(int.Parse(lsTotalNumberofColumns));

            int[] widthHeader = { 50, 29, 10, 5, 6 };
            PdfPtable.SetWidths(widthHeader);
            PdfPtable.HorizontalAlignment = 1;

            PdfPtable.TotalWidth = 100f;
            PdfPtable.WidthPercentage = 100f;

            PdfPCell PdfCell = new PdfPCell();

            String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
            int liTotalPage = (loInsertdataset.Tables[0].Rows.Count / (liPageSize - 3));
            int liCurrentPage = 0;
            if (loInsertdataset.Tables[0].Rows.Count % liPageSize != 0)
            {
                liTotalPage = liTotalPage + 1;
            }
            else
            {
                liPageSize = 17;
                liTotalPage = liTotalPage + 1;
            }

            //check the length of the column name to set the pagesize.
            for (int j = 0; j < loInsertdataset.Tables[0].Columns.Count; j++)
            {
                if (loInsertdataset.Tables[0].Columns[j].ColumnName.Length > 19)
                {
                    liPageSize = 17;
                }
            }

        PageSplitLogic:
            for (int liRowCount = 0; liRowCount < loInsertdataset.Tables[0].Rows.Count; liRowCount++)
            {
                if (liRowCount % liPageSize == 0)
                {
                    document.Add(PdfPtable);

                    if (liRowCount != 0)
                    {
                        liCurrentPage++;
                        document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, string.Empty));//document.add(addfooter(lsdatetime, litotalpage, licurrentpage, lipagesize, false, string.empty));                        
                        document.NewPage();
                    }

                    loInsertdataset.AcceptChanges();
                    setHeader(document, loInsertdataset, newdataset);

                    PdfPtable = new PdfPTable(int.Parse(lsTotalNumberofColumns));
                    PdfPtable.SetWidths(widthHeader);
                }
                int colsize = loInsertdataset.Tables[0].Columns.Count;

                /*
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////             
                /////////////////////////////////////////   Populating the cells of the Grid   ////////////////////////////////////////
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                */
                for (int liColumnCount = 0; liColumnCount < colsize; liColumnCount++)
                {
                    if (liColumnCount <= 4)
                    {
                        string strItalic = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["_ItalicsFlg"]);
                        string strBold = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["_BoldFlg"]);
                        string strMerge = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["_MergeFlg"]);
                        string strColor = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["_ColourCode"]);
                        string strRecordType = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount]["_RecordType"]);

                        iTextSharp.text.Paragraph lochunk = new Paragraph();
                        String lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);

                        if (lsFormatedString.Contains("Transactions to Consider for Bunching"))
                        {
                            //nothing
                        }
                        if (lsFormatedString.Length > 300)
                        {
                            int noOfLines = lsFormatedString.Length / 100;
                            if (noOfLines > 3)
                            {
                                document.Add(PdfPtable);
                                liCurrentPage++;
                                document.Add(addFooter("", liTotalPage, liCurrentPage, liPageSize, false, string.Empty));//document.add(addfooter(lsdatetime, litotalpage, licurrentpage, lipagesize, false, string.empty));                        
                                document.NewPage();
                                loInsertdataset.AcceptChanges();
                                setHeader(document, loInsertdataset, newdataset);
                                PdfPtable = new PdfPTable(int.Parse(lsTotalNumberofColumns));
                                PdfPtable.SetWidths(widthHeader);
                            }
                        }


                        //if (strMerge == null && liColumnCount == 1)
                        //{
                        //    lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount-1]);
                        //}
                        if (strBold == "1" && strColor.Contains("#") && strMerge == "1") //&& lsFormatedString != "")
                        {
                            lochunk = new Paragraph(lsFormatedString, setFontsAll(7, 1, 0));                            
                            PdfCell = new PdfPCell();
                            PdfCell.BackgroundColor = new Color(System.Drawing.ColorTranslator.FromHtml(strColor));
                        }
                        else if (strBold == "1" && strItalic == "1" && strMerge == "1") //&& lsFormatedString != "")
                        {
                            lochunk = new Paragraph(lsFormatedString, setFontsAll(7, 1, 1));
                            PdfCell = new PdfPCell();
                            lochunk.IndentationLeft = 7f;
                        }
                        else if (strMerge != "1" && liColumnCount == 0)
                        {
                            lochunk = new Paragraph(lsFormatedString, setFontsAll(7, 0, 0));
                            PdfCell = new PdfPCell();
                            lochunk.IndentationLeft = 15f;
                        }
                        //else if (strColor != null)
                        //{
                        //    lochunk = new Chunk(lsFormatedString, setFontsAll(7, 0, 0));
                        //    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml(strColor));
                        //}
                        else
                        {
                            lochunk = new Paragraph("" + lsFormatedString, setFontsAll(7, 0, 0));
                            PdfCell = new PdfPCell();
                        }
                        PdfCell.BorderWidth = 0.1F;
                        PdfCell.BorderColor = Color.LIGHT_GRAY;
                        PdfCell.VerticalAlignment = 1;
                        PdfCell.UseBorderPadding = true;
                        if (liColumnCount != 0)
                        {                            
                            PdfCell.HorizontalAlignment = 2;
                        }
                        //  loCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));

                        if (liColumnCount == 3 || liColumnCount == 4 || liColumnCount == 2)
                        {
                            PdfCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;//loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        }
                        if (liColumnCount == 5 || liColumnCount == 1)
                        {
                            PdfCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;//loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            if (liColumnCount == 6)
                            {
                                //loCell.NoWrap = false;
                                //loCell.MaxLines = 50;
                                //loCell.Leading = 8f;
                            }
                        }
                        if (liColumnCount == 2)
                        {
                            PdfCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        }

                        ///*=========START WITH BOLD AND SUPERBOLD FLAG========*/
                        //if (liColumnCount == 0)
                        //{

                        //    String abc = "" + lodataset.Tables[0].Rows[liRowCount][0].ToString();
                        //    lochunk = new Paragraph(abc, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                        //    //loCell.BackgroundColor = iTextSharp.text.Color.WHITE;
                        //}


                        //if (strMerge == "1")
                        //{
                        //    loCell.Add(lochunk);
                        //}
                        //else
                        //{
                        PdfCell.AddElement(lochunk);

                        //}
                        if (ddlHH.SelectedItem.Text == "ALL" && strRecordType == "Category")
                        { }
                        else
                        {                            
                            PdfPtable.AddCell(PdfCell);
                        }
                    }
                    else
                    {

                    }
                }
                try
                {
                    if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
                    {
                        PdfPtable.SplitRows = false;
                        PdfPtable.KeepTogether = true;
                        document.Add(PdfPtable);                        
                        liCurrentPage = liCurrentPage + 1;
                        document.Add(addFooter("", liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));//document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                    }
                }
                catch (Exception Ex)
                {

                }
            }
            /* Throws a Popup window for the PDF file */
            if (loInsertdataset.Tables[0].Rows.Count > 0)
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                Response.Write("<script>");
                string strFileNameOfPDF = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
                Response.Write("window.open('" + strFileNameOfPDF + "', 'mywindow')");
                Response.Write("</script>");
                lblMessage.ForeColor = System.Drawing.Color.Green;
                lblMessage.Text = "Annual Client Report was generated succesfully!";
            }
            else
            {
                lblMessage.ForeColor = System.Drawing.Color.Red;
                lblMessage.Text = "No data was found for generating the report.";
            }
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while generating the report. Details: " + ex.Message;
        }
    }
    public string getFinalSp()
    {
        year = ddlYear.SelectedValue;
        String lsSQL = "";

        if (ddlYear.SelectedItem.Text == "ALL")
            year = "null";
        if (ddlRespParty.SelectedIndex > 0)
            RespParty = "'" + ddlRespParty.SelectedValue + "'";
        else
            RespParty = "null";
        if (ddlHH.SelectedIndex > 0)
            HH = "'" + ddlHH.SelectedValue + "'";
        else
            HH = "null";
        if (ddlTaskType.SelectedIndex > 0)
            TaskType = "'" + ddlTaskType.SelectedValue + "'";
        else
            TaskType = "null";
        if (ListBoxStatuses.SelectedIndex == 0)
            Status = "null";
        else
        {
            GeneralMethods clsGM = new GeneralMethods();
            Status = "'" + clsGM.GetMultipleSelectedItemsFromListBox(ListBoxStatuses) + "'";
        }

        lsSQL = "SP_R_ANNUALREVIEW @Year=" + year +
                                                ",@HouseHoldUUID=" + HH +
                                                ",@TaskTypID=" + TaskType +
                                                ",@StatusID=" + Status +
                                                ",@ResponsiblePartyUUID=" + RespParty + "";

        return lsSQL;
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
            case "1":
                int[] headerwidths1 = { 100 };
                fotable.SetWidths(headerwidths1);
                fotable.Width = 100;
                break;
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
                int[] headerwidths5 = { 50, 29, 10, 5, 6 };//int[] headerwidths5 = { 50, 25, 9, 7, 8 };//int[] headerwidths5 = { 30, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths5);
                fotable.Width = 100;
                break;
            case "6":
                int[] headerwidths6 = { 10, 29, 9, 9, 9, 9 };//int[] headerwidths6 = { 30, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths6);
                fotable.Width = 76;
                break;
            case "7":
                int[] headerwidths7 = { 12, 20, 20, 9, 12, 10, 20 };
                fotable.SetWidths(headerwidths7);
                fotable.Width = 94;
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
    public iTextSharp.text.Table addFooter(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt)
    {

        iTextSharp.text.Table fotable = new iTextSharp.text.Table(2, 1);
        fotable.Width = 97;
        fotable.Border = 0;
        int[] headerwidths = { 54, 43 };
        fotable.SetWidths(headerwidths);
        fotable.Cellpadding = 0;
        Cell loCell = new Cell();
        Chunk loChunk = new Chunk();

        for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData; liCounter++)
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
            loCell.BorderWidth = 0;
            loCell.Add(loChunk);
            fotable.AddCell(loCell);

            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.BorderWidth = 0;
            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
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


        /* Please uncomment this to show page numbers */

        //loCell = new Cell();
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
        //loCell.Leading = 15f;//25f
        //loCell.HorizontalAlignment = 2;
        //loCell.BorderWidth = 0;
        //loCell.Add(loChunk);
        //fotable.AddCell(loCell);

        loCell = new Cell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font7GreyItalic());
        loCell.Add(loChunk);
        loCell.Leading = 15f;//25f
        loCell.BorderWidth = 0;
        loCell.HorizontalAlignment = 2;
        fotable.AddCell(loCell);
        //fotable.TableFitsPage = true;

        return fotable;
    }
    public void setHeader(Document foDocument, DataSet loInsertdataset, DataSet loDatatset)
    {
        //iTextSharp.text.Table loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, 4);   // 2 rows, 2 columns        
        // DataSet OldDataset = loInsertdataset.Copy();

        DataTable table = loDatatset.Tables[0].Clone();

        DataSet OldDataset = new DataSet();
        OldDataset.Tables.Add(table);

        //for (int liNewdataset = loInsertdataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        //{
        //    if (loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || loInsertdataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
        //    {
        //        loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
        //    }
        //}

        DataSet DSTest = loDatatset.Copy();

        //DSTest.Tables[0].Columns.RemoveAt(14);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(13);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(12);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(11);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(10);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(9);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(8);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(7);
        //DSTest.Tables[0].AcceptChanges();
        //DSTest.Tables[0].Columns.RemoveAt(6);
        //DSTest.Tables[0].AcceptChanges();

        DataSet AddDataset = DSTest.Copy();
        AddDataset.AcceptChanges();
        
        PdfPTable pdfPTable = new PdfPTable(int.Parse(lsTotalNumberofColumns));
        lsTotalNumberofColumns = "1";

        PdfPTable pdfPHeaderTable = new PdfPTable(int.Parse(lsTotalNumberofColumns));
        lsTotalNumberofColumns = "5";

        int[] widthHeaderTable = { 50, 29, 10, 5, 6 };
        pdfPTable.SetWidths(widthHeaderTable);

        pdfPTable.TotalWidth = 100f;
        //pdfPTable.WidthPercentage = 100f;

        int[] widthDataTable = { 100 };
        pdfPHeaderTable.SetWidths(widthDataTable);

        pdfPHeaderTable.TotalWidth = 100f;
        //pdfPHeaderTable.WidthPercentage = 100f;

        Chunk loParagraph = new Chunk();


        //////// Set Header Title details for PDF  ////////        
        string strYear = string.Empty;
        string strHH = string.Empty;
        string strTaskType = string.Empty;
        string strRespParty = string.Empty;

        if (ddlYear.SelectedValue != null)
            strYear = ddlYear.SelectedItem.Text;
        if (ddlHH.SelectedValue != null)
            strHH = ddlHH.SelectedItem.Text;
        if (ddlRespParty.SelectedValue != null)
            strRespParty = ddlRespParty.SelectedItem.Text;
        if (ddlTaskType.SelectedValue != null)
            strTaskType = ddlTaskType.SelectedItem.Text;

        GeneralMethods clsGM = new GeneralMethods();
        string strStatus = clsGM.GetMultipleSelectedItemsFromListBox(ListBoxStatuses);
        //Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));

        /*
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////             
        /////////////////////////////////////////   Populating the cells of the Header   //////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        */

        Chunk lochunk = new Chunk("", setFontsAll(12, 1, 0));
        PdfPCell loCell = new PdfPCell();

        loCell.AddElement(lochunk);

        if (strYear == "ALL")
        {
            lochunk = new Chunk("All Years - Annual Client Review\n", setFontsAll(12, 0, 0));
            loCell.AddElement(lochunk);

        }
        else
        {
            lochunk = new Chunk(strYear + " Annual Client Review\n", setFontsAll(12, 1, 0));
            loCell.AddElement(lochunk);

        }

        if (strHH == "ALL")
        {
            //lochunk = new Chunk("All Households", setFontsAll(12, 0, 0));
            //loCell.Add(lochunk);
            if (strRespParty == "ALL")
            {
                lochunk = new Chunk("All Responsible Parties", setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
            else if (strRespParty != "ALL")
            {
                lochunk = new Chunk(strRespParty, setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
            if (strTaskType == "ALL")
            {
                lochunk = new Chunk(" - All Tasks", setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
            else
            {
                lochunk = new Chunk(" - " + strTaskType, setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
            if (strStatus == "0")
            {
                //lochunk = new Chunk(" - " + clsGM.GetALLItemsTEXTFromListBox(ListBoxStatuses) + " Status", setFontsAll(12, 0, 0));
                lochunk = new Chunk(" - All Statuses", setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
            else
            {
                clsGM = new GeneralMethods();
                string strSelectedStatus = clsGM.GetMultipleSelectedItemsTEXTFromListBox(ListBoxStatuses);
                lochunk = new Chunk(" - " + strSelectedStatus + " Status", setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
        }
        else
        {
            lochunk = new Chunk(strHH, setFontsAll(12, 0, 0));
            loCell.AddElement(lochunk);

            if (strRespParty != "ALL")
            {
                lochunk = new Chunk(" - " + strRespParty, setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
            if (strTaskType != "ALL")
            {
                lochunk = new Chunk(" - " + strTaskType, setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
            if (strStatus != "0")
            {
                clsGM = new GeneralMethods();
                string strSelectedStatus = clsGM.GetMultipleSelectedItemsTEXTFromListBox(ListBoxStatuses);
                lochunk = new Chunk(" - " + strSelectedStatus + " Status", setFontsAll(12, 0, 0));
                loCell.AddElement(lochunk);

            }
        }


        //loParagraph.Chunks.Add(lochunk);

        loCell.Colspan = 1;//loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        //loCell.HorizontalAlignment = 1;

        //lochunk = new Chunk("\n" + DateTime.Now.ToString("MMMM dd, yyyy"), setFontsAll(8, 0, 1)); //To Show date in header uncomment this
        //loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
        pdfPHeaderTable.AddCell(loCell);

        /////////////////////////////* Populating the cells of the Column Header *////////////////////////////////
        for (int liColumnCount = 0; liColumnCount < AddDataset.Tables[0].Columns.Count; liColumnCount++)
        {
            if (liColumnCount <= 4)
            {
                if (Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName) != "")
                {
                    if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals(" "))
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace(" ", "Task"), setFontsAll(7, 1, 0));
                    }
                    else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("TASK"))
                    {
                        lochunk = new Chunk("  " + Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("TASK", ""), setFontsAll(7, 1, 0));
                    }
                    else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("NOTES"))
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("NOTES", "Notes"), setFontsAll(7, 1, 0));
                    }
                    else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Responsible Party"))
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Responsible Party", "Responsible Party"), setFontsAll(7, 1, 0));
                    }
                    else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Status"))
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Status", "Status"), setFontsAll(7, 1, 0));
                    }
                    else if (AddDataset.Tables[0].Columns[liColumnCount].ColumnName.Equals("Date Complete"))
                    {
                        lochunk = new Chunk(Convert.ToString(AddDataset.Tables[0].Columns[liColumnCount].ColumnName).Replace("Date Complete", "Date Complete"), setFontsAll(7, 1, 0));
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
                loCell = new PdfPCell();

                loCell.AddElement(lochunk);
                loCell.Border = 0;

                //loCell.NoWrap = false;//true;

                loCell.SetLeading(-2F, 0F);

                if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
                {
                    loCell.SetLeading(10F, 0F);
                    //loCell.Leading = 9f;
                }
                loCell.SetLeading(10F, 0f);
                
                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));//iTextSharp.text.Color.LIGHT_GRAY;
                loCell.VerticalAlignment = Element.ALIGN_MIDDLE;//1;//5 ,6 bottom : WASTE VALUES - 3,4

                pdfPTable.AddCell(loCell);
            }
        }
        
        foDocument.Add(pdfPHeaderTable);
        foDocument.Add(pdfPTable);
        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        loCell = new PdfPCell(png);
        loCell.HorizontalAlignment = Element.ALIGN_CENTER;
        loCell.VerticalAlignment = Element.ALIGN_MIDDLE;        
        //png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        pdfPHeaderTable.AddCell(loCell);
        foDocument.Add(png);
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
    public void setBottomWidthWhite(Cell foCell)
    {
        foCell.BorderWidthBottom = 0;
        foCell.BorderColorBottom = new iTextSharp.text.Color(255, 255, 255);
    }
    #endregion

    #region Clear_message_on_dropdown_change
    protected void ddlYear_SelectedIndexChanged(object sender, EventArgs e)
    {
        clearMesssage();
    }
    protected void ddlRespParty_SelectedIndexChanged(object sender, EventArgs e)
    {
        clearMesssage();
    }
    protected void ddlHH_SelectedIndexChanged(object sender, EventArgs e)
    {
        clearMesssage();
    }
    protected void ddlTaskType_SelectedIndexChanged(object sender, EventArgs e)
    {
        clearMesssage();
    }
    protected void ListBoxStatuses_SelectedIndexChanged(object sender, EventArgs e)
    {
        clearMesssage();
    }
    private void clearMesssage()
    {
        if (lblMessage.Text != string.Empty)
            lblMessage.Text = string.Empty;
    }
    #endregion

    #region Generating the PDF File using RDL
    public void getValues()
    {
        year = ddlYear.SelectedValue;

        if (ddlYear.SelectedItem.Text == "ALL")
            year = "null";
        if (ddlRespParty.SelectedIndex > 0)
            RespParty = ddlRespParty.SelectedValue;
        else
            RespParty = "null";
        if (ddlHH.SelectedIndex > 0)
            HH = ddlHH.SelectedValue;
        else
            HH = "null";
        if (ddlTaskType.SelectedIndex > 0)
            TaskType = ddlTaskType.SelectedValue;
        else
            TaskType = "null";
        if (ListBoxStatuses.SelectedIndex == 0)
            Status = "null";
        else
        {
            GeneralMethods clsGM = new GeneralMethods();
            Status = clsGM.GetMultipleSelectedItemsFromListBox(ListBoxStatuses);
        }
    }
    protected void btnGeneratePDFReport_Click(object sender, EventArgs e)
    {
        try
        {
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string extension;
            DB clsDB = new DB();

            //////// Set Header Title details for PDF  ////////        
            string strYear = string.Empty;
            string strHH = string.Empty;
            string strTaskType = string.Empty;
            string strRespParty = string.Empty;
            string strHeaderLine1 = string.Empty;
            string strHeaderLine2 = string.Empty;

            if (ddlYear.SelectedValue != null)
                strYear = ddlYear.SelectedItem.Text;
            if (ddlHH.SelectedValue != null)
                strHH = ddlHH.SelectedItem.Text;
            if (ddlRespParty.SelectedValue != null)
                strRespParty = ddlRespParty.SelectedItem.Text;
            if (ddlTaskType.SelectedValue != null)
                strTaskType = ddlTaskType.SelectedItem.Text;
            GeneralMethods clsGM = new GeneralMethods();
            string strStatus = clsGM.GetMultipleSelectedItemsFromListBox(ListBoxStatuses);

            if (strYear == "ALL")
                strHeaderLine1 = "All Years - Annual Client Review";
            else
                strHeaderLine1 = strYear + " Annual Client Review";

            if (strHH == "ALL")
            {
                //lochunk = new Chunk("All Households", setFontsAll(12, 0, 0));
                //loCell.Add(lochunk);
                if (strRespParty == "ALL")
                    strHeaderLine2 = "All Responsible Parties";
                else if (strRespParty != "ALL")
                    strHeaderLine2 = strRespParty;
                if (strTaskType == "ALL")
                    strHeaderLine2 = strHeaderLine2 + " - All Tasks";
                else
                    strHeaderLine2 = strHeaderLine2 + " - " + strTaskType;
                if (strStatus == "0")
                    strHeaderLine2 = strHeaderLine2 + " - All Statuses";
                else
                {
                    clsGM = new GeneralMethods();
                    string strSelectedStatus = clsGM.GetMultipleSelectedItemsTEXTFromListBox(ListBoxStatuses);
                    strHeaderLine2 = strHeaderLine2 + " - " + strSelectedStatus + " Status";
                }
            }
            else
            {
                strHeaderLine2 = strHH;
                if (strRespParty != "ALL")
                    strHeaderLine2 = strHeaderLine2 + " - " + strRespParty;
                if (strTaskType != "ALL")
                    strHeaderLine2 = strHeaderLine2 + " - " + strTaskType;
                if (strStatus != "0")
                {
                    clsGM = new GeneralMethods();
                    string strSelectedStatus = clsGM.GetMultipleSelectedItemsTEXTFromListBox(ListBoxStatuses);
                    strHeaderLine2 = strHeaderLine2 + " - " + strSelectedStatus + " Status";
                }
            }

            string lsSQL = getFinalSp();
            DataSet newdataset = clsDB.getDataSet(lsSQL);

            if (newdataset.Tables.Count > 0)
            {
                if (newdataset.Tables[0].Rows.Count < 1)
                {
                    lblMessage.ForeColor = System.Drawing.Color.Red;
                    lblMessage.Text = "No Records were found";
                    return;
                }
            }
            //Response.Redirect("http://webdevserver/ReportServer_NEWSERVERDB/Pages/ReportViewer.aspx?%2fBasicSSRSReport%2fAnnualClient&rs:Command=Render&rs:Format=PDF", true);
            //Response.Redirect("http://sql-test/ReportServer/Pages/ReportViewer.aspx?%2fAnnualClientReview%2fAllHH&rs:Command=Render&rs:Format=PDF", true);

            //Fetch Values from UI controls
            getValues();

            string strGUID = DateTime.Now.ToString("MMddyyhhmmss");
            ReportParameter[] param = new ReportParameter[7];
            param[0] = new ReportParameter("Year", Convert.ToString(year));
            param[1] = new ReportParameter("HouseHoldUUID", Convert.ToString(HH));
            param[2] = new ReportParameter("TaskTypID", Convert.ToString(TaskType));
            param[3] = new ReportParameter("StatusID", Convert.ToString(Status));
            param[4] = new ReportParameter("ResponsiblePartyUUID", Convert.ToString(RespParty));
            param[5] = new ReportParameter("HeaderLine1", strHeaderLine1);
            param[6] = new ReportParameter("HeaderLine2", strHeaderLine2);

            //ReportDataSource rds = new ReportDataSource("SQLDataSource", newdataset.Tables[0]);
            ReportViewer viewer = new ReportViewer();

            viewer.ProcessingMode = ProcessingMode.Remote;
            viewer.ServerReport.ReportServerCredentials = new ReportServerNetworkCredentials();            
            
            //viewer.ServerReport.ReportServerUrl = new Uri("http://sql-test/ReportServer/"); //report server for TEST SERVER
           // viewer.ServerReport.ReportServerUrl = new Uri("http://grpao1-vwsql02/ReportServer"); //report server for PROD SERVER
 //viewer.ServerReport.ReportServerUrl = new Uri("http://gp-db1/ReportServer"); //report server for GP_CRM1 SERVER
            viewer.ServerReport.ReportServerUrl = new Uri("http://gp-PRODDB/ReportServer"); //report server for GP_PRODCRM SERVER // added 4_17_2019

            if (ddlHH.SelectedIndex > 0)
                viewer.ServerReport.ReportPath = "/AnnualClientReview/SingleHH"; // rdl name
            else
                viewer.ServerReport.ReportPath = "/AnnualClientReview/AllHH"; // rdl name

            viewer.ServerReport.SetParameters(param);
            viewer.ServerReport.Refresh();
            //viewer.LocalReport.DataSources.Add(rds);
            byte[] bytes = viewer.ServerReport.Render("PDF", null, out mimeType, out encoding, out extension, out streamids, out warnings);

            Response.Buffer = true;
            Response.Clear();
            Response.ContentType = mimeType;
            Response.AddHeader("content-disposition", "attachment; filename= " + strGUID + "." + extension);
            Response.OutputStream.Write(bytes, 0, bytes.Length); // create the file  
            Response.Flush(); // send it to the client to download  
            Response.End();            
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while generating the report. Details: " + ex.Message;
        }
    }
    #endregion
}