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
using System.Net;
using System.Collections;
using Spire.Xls;
using System.Drawing;
using System.Data.Common;
using System.Xml;
using System.Data.SqlClient;

public partial class BDPipeline : System.Web.UI.Page 
{
    String sqlstr = string.Empty;
    Boolean fbCheckExcel = false;
    GeneralMethods clsGM = new GeneralMethods();
    DataSet ds;
    SqlDataAdapter dagersham;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            mvShowReport.ActiveViewIndex = 0;

            BindState(ddlState);
            BindCity(ddlCity);

            BindContactOwner(ddlContactOwner);
            BindContactType(lstContactType);
            BindFirms(ddlFirms);
            lstType.SelectedIndex = 0;
        }
    }

    private void BindContactType(ListBox ddl)
    {
        ddl.Items.Clear();
        sqlstr = "EXEC SP_S_Contact_Type_LKUP";
        clsGM.getListForBindListBox(ddl, sqlstr, "Type", "Type");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }


    private void BindFirms(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "EXEC SP_S_FIRM_LKUP";
        clsGM.getListForBindDDL(ddl, sqlstr, "Firm", "Firm");



        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindContactOwner(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "EXEC SP_S_BUSINESS_DEVELOPMENT_CONTACT_OWNER";
        clsGM.getListForBindDDL(ddl, sqlstr, "Contact Owner", "OwnerId");



        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }


    private void BindState(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "EXEC SP_S_STATE_LKUP";
        clsGM.getListForBindDDL(ddl, sqlstr, "State", "State");



        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    private void BindCity(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "EXEC SP_S_CITY_LKUP";
        clsGM.getListForBindDDL(ddl, sqlstr, "City", "City");



        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    protected void Button1_Click(object sender, EventArgs e)
    {

        lblError.Text = "";
        lblmessage.Text = "";
        if (RadioButton1.Checked)
        {
            try
            {
                fbCheckExcel = false;

            }
            catch (Exception ex)
            {
                //Response.Write(ex.ToString());
                //Response.Write(ex.StackTrace);
            }

            Generatereport();
        }
        if (RadioButton2.Checked)
        {
            try
            {
                fbCheckExcel = true;

            }
            catch (Exception ex)
            {
                //Response.Write(ex.ToString());
                //Response.Write(ex.StackTrace);
            }

            generatesExcelsheets();
        }
    }

    public void Generatereport()
    {
        //  String lsSQL = "SP_R_INVESTMENT_TEAM_REPORT_DETAIL @ClosedDT = " + ClosedDT + ",@FundIdListTxt = " + FundIdListTxt;// "select * from investment_report";// getFinalSp();
        object ContactOwnerId = ddlContactOwner.SelectedValue == "0" ? "null" : "'" + ddlContactOwner.SelectedValue + "'";
        lblContactName.Text = "by Rank";//ddlContactOwner.SelectedValue == "0" ? "" : ddlContactOwner.SelectedItem.Text;

        string startDate = txtStartDate.Text.Trim() == "" ? "null" : "'" + txtStartDate.Text.Trim() + "'";
        string endDate = txtEndDate.Text.Trim() == "" ? "null" : "'" + txtEndDate.Text.Trim() + "'";

        object PipelineId = ddlPipeline.SelectedValue == "0" ? "null" : "'" + ddlPipeline.SelectedValue + "'";
        Label1.Text = ddlPipeline.SelectedValue == "0" ? "" : ddlPipeline.SelectedItem.Text + "  " + "Pipeline";
        object StateId = ddlState.SelectedValue == "0" ? "null" : "'" + ddlState.SelectedValue + "'";
        object CityId = ddlCity.SelectedValue == "0" ? "null" : "'" + ddlCity.SelectedValue + "'";
        object FirmId = ddlFirms.SelectedValue == "0" ? "null" : "'" + ddlFirms.SelectedValue + "'";

        object ContactType = lstContactType.SelectedValue == "0" || lstType.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstContactType) + "'"; //lstContactType.SelectedValue == "0" ? "null" : "'" + lstContactType.SelectedValue + "'";
        object Type = lstType.SelectedValue == "0" || lstType.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'";

        //String lsSQL = "exec SP_R_BUSINESS_DEVELOPMENT_TOUCHPOINTS_REPORT  @ContactOwnerid = " + ContactOwnerId
        //                                                                  + ",@StartDT=" + startDate
        //                                                                  + ",@EndDT=" + endDate;

        if (txtStartDate.Text == "" && txtEndDate.Text == "")
        {
            lblDate.Text = "all dates";
        }

        String lsSQL = "exec SP_R_TOUCHPOINTS_BD_PIPELINE_REPORT  @ContactOwnerid = " + ContactOwnerId + ", @State= " + StateId + ", @City = " + CityId + ", @Firm = " + FirmId + " , @Pipeline = " + PipelineId + " , @StartDt = " + startDate + ", @EndDt = " + endDate + ", @Contacttype = " + ContactType + ", @Type = " + Type;
        
        DB clsDB = new DB();
        DataSet lodataset;
        lodataset = null;
        //  Response.Write(lsSQL);
        lodataset = clsDB.getDataSet(lsSQL);
        gvReport.AutoGenerateColumns = false;
        //gvReport.Controls.Clear();
        //gvReport.Columns.Clear();
        gvReport.DataSource = null;
        gvReport.DataBind();

        //lodataset.Tables[0].Rows.RemoveAt(0);
        //lodataset.AcceptChanges();
        //lodataset.Tables[0].Rows.RemoveAt(0);
        //lodataset.AcceptChanges();

        if (lodataset.Tables.Count >0)
        {
            if (lodataset.Tables[0].Rows.Count < 2)
            {
                lblError.Text = "No Record found";
                return;
            }
            else
            {
                mvShowReport.ActiveViewIndex = 1;
            }

        }
        else if (lodataset.Tables.Count == 0)
        {
            lblError.Text = "No Record found";
            return;
        }
        


        lblDate.Text = startDate.Replace("null", "").Replace("'", "") + " - " + endDate.Replace("null", "").Replace("'", "");

        if (txtStartDate.Text == "" && txtEndDate.Text == "")
        {
            lblDate.Text = "all dates";
        }


        if (lblDate.Text.StartsWith("- ") || lblDate.Text.EndsWith(" -"))
            lblDate.Text = lblDate.Text.Replace("-", "");

        if (lodataset.Tables.Count > 0)
        {
            foreach (DataColumn dc in lodataset.Tables[0].Columns)
            {
                BoundField newboundfiled = new BoundField();
                newboundfiled.DataField = dc.ColumnName;

                newboundfiled.HeaderText = dc.ColumnName;

                if (dc.ColumnName.Substring(0, 1) != "_")
                {
                    // gvReport.Columns.Add(newboundfiled);
                }

                if (dc.DataType.ToString() == "System.Decimal" || dc.DataType.ToString() == "System.Double")
                {
                    newboundfiled.HtmlEncode = false;
                    newboundfiled.DataFormatString = "{0:$#,###0;($#,###0)}";
                    newboundfiled.HeaderStyle.HorizontalAlign = HorizontalAlign.Right;
                    newboundfiled.ItemStyle.HorizontalAlign = HorizontalAlign.Right;

                }
                else
                {
                    newboundfiled.HeaderStyle.HorizontalAlign = HorizontalAlign.Left;
                    newboundfiled.ItemStyle.HorizontalAlign = HorizontalAlign.Left;
                }

                newboundfiled.HeaderStyle.VerticalAlign = VerticalAlign.Middle;
                //newboundfiled.HeaderStyle.HorizontalAlign = HorizontalAlign.Left;

            }
        }

       

        //gvReport.Columns[5].Visible = true;
        //gvReport.Columns[6].Visible = true;
        //gvReport.Columns[7].Visible = true;
        if (lodataset.Tables.Count > 0)
        {
            gvReport.DataSource = lodataset;
            gvReport.DataBind();
        }
        //gvReport.Columns[5].Visible = false;
        //gvReport.Columns[6].Visible = false;
        //gvReport.Columns[7].Visible = false;

        if (gvReport.Rows.Count > 0)
        {
            lblmessage.Text = "";
            lblmessage.Visible = false;
            // gvReport.HeaderStyle.Font.Bold = true;
            gvReport.HeaderStyle.Font.Size = FontUnit.Point(12);
            gvReport.Columns[0].ItemStyle.Width = 100;
            gvReport.Columns[1].ItemStyle.Width = 100;
            //gvReport.Columns[2].ItemStyle.Width = 100;
            //gvReport.Columns[3].ItemStyle.Width = 10;
            //gvReport.Columns[4].ItemStyle.Width = 100;
            gvReport.Columns[5].ItemStyle.Width = 100;
            gvReport.Columns[6].ItemStyle.Width = 100;
            gvReport.Columns[7].ItemStyle.Width = 100;
            //gvReport.Columns[8].ItemStyle.Width = 100;
            //gvReport.Columns[0].HeaderStyle.Width = 200; // Fund

            gvReport.HeaderRow.Cells[0].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[1].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[2].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[3].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[4].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[5].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[6].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[7].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.Cells[8].Style.Add("padding-left", "10px");
            gvReport.HeaderRow.BackColor = Color.Gray;


            for (int i = 0; i < lodataset.Tables[0].Rows.Count; i++)
            {
                gvReport.Rows[i].Cells[0].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[1].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[2].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[3].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[4].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[5].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[6].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[7].Style.Add("padding-left", "10px");
                gvReport.Rows[i].Cells[8].Style.Add("padding-left", "10px");
            }
        }
        else
        {
            lblmessage.Text = "No Records found.";
            lblmessage.Visible = true;
        }
    }
    protected void btnBack_Click(object sender, EventArgs e)
    {
        lblError.Text = "";
        mvShowReport.ActiveViewIndex = 0;
    }
    protected void BtnExport_Click(object sender, EventArgs e)
    {
        fbCheckExcel = true;
        generatesExcelsheets();
    }
    public void generatesExcelsheets()
    {
        object ContactOwnerId = ddlContactOwner.SelectedValue == "0" ? "null" : "'" + ddlContactOwner.SelectedValue + "'";

        //lblDate.Text = DateTime.Now.ToString("MMM dd, yyyy");
        string startDate = txtStartDate.Text.Trim() == "" ? "null" : "'" + txtStartDate.Text.Trim() + "'";
        string endDate = txtEndDate.Text.Trim() == "" ? "null" : "'" + txtEndDate.Text.Trim() + "'";

        object PipelineId = ddlPipeline.SelectedValue == "0" ? "null" : "'" + ddlPipeline.SelectedValue + "'";
        object StateId = ddlState.SelectedValue == "0" ? "null" : "'" + ddlState.SelectedValue + "'";
        object CityId = ddlCity.SelectedValue == "0" ? "null" : "'" + ddlCity.SelectedValue + "'";
        object FirmId = ddlFirms.SelectedValue == "0" ? "null" : "'" + ddlFirms.SelectedValue + "'";
        //lblDate.Text = DateTime.Now.ToString("MMM dd, yyyy");
        object ContactType = lstContactType.SelectedValue == "0" || lstType.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstContactType) + "'"; //lstContactType.SelectedValue == "0" ? "null" : "'" + lstContactType.SelectedValue + "'";
        object Type = lstType.SelectedValue == "0" || lstType.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'";


        //String lsSQL = "exec SP_R_BUSINESS_DEVELOPMENT_TOUCHPOINTS_REPORT  @ContactOwnerid = " + ContactOwnerId
        //                                                                  + ",@StartDT=" + startDate
        //                                                                  + ",@EndDT=" + endDate;

        String lsSQL = "exec SP_R_TOUCHPOINTS_BD_PIPELINE_REPORT  @ContactOwnerid = " + ContactOwnerId + ", @State= " + StateId + ", @City = " + CityId + ", @Firm = " + FirmId + " , @Pipeline = " + PipelineId + " , @StartDt = " + startDate + ", @EndDt = " + endDate + ", @Contacttype = " + ContactType + ", @Type = " + Type;
        
        
        DB clsDB = new DB();
        DataSet lodataset;
        lodataset = null;
        lodataset = clsDB.getDataSet(lsSQL);

        string TitleContact = ddlContactOwner.SelectedValue == "0" ? "" : ddlContactOwner.SelectedItem.Text;

        if (lodataset.Tables.Count > 0)
        {
            if (lodataset.Tables[0].Rows.Count < 1)
            {
                lblError.Text = "No Record found";
                return;
            }
        }

        //lodataset.Tables[0].Rows.RemoveAt(0);
        //lodataset.AcceptChanges();
        ////lodataset.Tables[0].Rows.RemoveAt(0);
        //lodataset.AcceptChanges();
        //DataSet lodataset = GetReportData();
        DataSet loInsertdataset = lodataset.Copy();

        if (loInsertdataset.Tables.Count > 0)
        {
            if (loInsertdataset.Tables[0].Rows.Count < 1)
            {
                lblError.Text = "No Records found.";
                return;
            }

        }

        if (loInsertdataset.Tables.Count > 0)
        {
            loInsertdataset.Tables[0].Columns.Remove("_ContactId");
            loInsertdataset.Tables[0].Columns.Remove("_OrderNmb");
            loInsertdataset.Tables[0].Columns.Remove("_typeflg");
            loInsertdataset.Tables[0].Columns.Remove("_BDRankProspectStatus");
            loInsertdataset.Tables[0].Columns.Remove("_contactname");
            loInsertdataset.Tables[0].Columns.Remove("_contactType");
            loInsertdataset.Tables[0].Columns.Remove("_OwnerId");
        }

        int liTtrow = 0;

        loInsertdataset.AcceptChanges();

        String lsFileNamforFinalXls = "BDPipeline_" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xls";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\RecommendationRelated.xls");
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls);
        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls.Replace("xls", "xml"));
        // Response.Write(strDirectory);
        string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strDirectory + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
        DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");

        FileInfo loFile = new FileInfo(strDirectory1);
        loFile.CopyTo(strDirectory, true);


        using (DbConnection connection = factory.CreateConnection())
        {
            connection.ConnectionString = connectionString;
            connection.Open();
            string sheetNumber = string.Empty;
            for (int j = 0; j < loInsertdataset.Tables.Count; j++)
            {
                ////////////////////////// Header insert Starts //////////////////////////////
                sheetNumber = Convert.ToString(j + 1);
                String lsFirstColumn = "Insert into [Sheet" + sheetNumber + "$] (";
                String lsFiled = "";
                String lsFieldvalue = "";
                for (int liColumns = 0; liColumns < loInsertdataset.Tables[j].Columns.Count; liColumns++)
                {
                    lsFieldvalue += "'" + loInsertdataset.Tables[j].Columns[liColumns].ColumnName.Replace("'", "''") + "'";
                    lsFiled += "id" + (liColumns + 1);
                    if (liColumns < loInsertdataset.Tables[j].Columns.Count - 1)
                    {
                        lsFieldvalue = lsFieldvalue + ",";
                        lsFiled = lsFiled + ",";
                    }
                }
                lsFirstColumn = lsFirstColumn + lsFiled + ")" + " Values (" + lsFieldvalue + ")";

                //using (DbConnection connection = factory.CreateConnection())
                //{
                //    connection.ConnectionString = connectionString;

                using (DbCommand command = connection.CreateCommand())
                {
                    try
                    {
                        command.CommandText = lsFirstColumn;
                        //connection.Open();
                        command.ExecuteNonQuery();
                        //connection.Close();
                    }
                    catch (Exception exc)
                    {
                        //Response.Write(exc.Message);
                    }
                }
                //}
                ////////////////////////// Header insert ends //////////////////////////////
                int insertCount = 0;
                string fieldData = string.Empty;
                for (int liCounter = 0; liCounter < loInsertdataset.Tables[j].Rows.Count; liCounter++)
                {
                    lsFirstColumn = "Insert into [Sheet" + sheetNumber + "$] (";

                    lsFieldvalue = "";
                    for (int liColumns = 0; liColumns < loInsertdataset.Tables[j].Columns.Count; liColumns++)
                    {
                        //if (liColumns != 0 && !loInsertdataset.Tables[0].Columns[liColumns].ColumnName.Contains("_"))
                        //{
                        fieldData = loInsertdataset.Tables[j].Rows[liCounter][liColumns].ToString().Replace("'", "''").Length > 250 ? loInsertdataset.Tables[j].Rows[liCounter][liColumns].ToString().Replace("'", "''").Substring(0, 250) : loInsertdataset.Tables[j].Rows[liCounter][liColumns].ToString().Replace("'", "''");
                        lsFieldvalue += "'" + fieldData + "'";
                        //lsFieldvalue = lsFieldvalue.Length > 250 ? lsFieldvalue.Substring(0, 250) + "'" : lsFieldvalue;
                        if (liColumns < loInsertdataset.Tables[j].Columns.Count - 1)
                        {
                            lsFieldvalue = lsFieldvalue + ",";
                            //lsFieldvalue = lsFieldvalue.Length > 250 ? lsFieldvalue.Substring(0, 250) + "'" + "," : lsFieldvalue + ",";
                        }
                        //}
                    }
                    lsFirstColumn = lsFirstColumn + lsFiled + ")" + " Values (" + lsFieldvalue + ")";

                    using (DbCommand command = connection.CreateCommand())
                    {

                        try
                        {
                            command.CommandText = lsFirstColumn;
                            command.ExecuteNonQuery();
                            insertCount++;
                        }
                        catch (Exception exc)
                        {
                            //Response.Write(lsFirstColumn);
                            Response.Write("<br/>" + exc.Message);

                        }
                    }

                }

                //Response.Write(insertCount);
            }
            connection.Close();
        }

        if (1 == 1)
        {
            int DatarowCount = 2000;
            string strCombName = "by Rank";

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(strDirectory);

            for (int sheetNo = 0; sheetNo < 1; sheetNo++)
            {
                Worksheet sheet = workbook.Worksheets[sheetNo];

                DatarowCount = sheet.Rows.Length;
                #region StyleUsing Spire.xls
                //Gets worksheet

                //Worksheet sheetCover = workbook.Worksheets[0];
                sheet.PageSetup.TopMargin = 0.25;

                //sheet.GridLinesVisible = false;

                //remove header
                for (int liRemoveheader = 1; liRemoveheader < 23; liRemoveheader++)
                {
                    sheet.Range[1, liRemoveheader].Text = "";
                }

                bool blstbxSelected = false;
                int count = 0;


                //sheet.Range[5, 1].Text = strCombName;


                switch (sheetNo)
                {
                    case 0:
                        sheet.Range[2, 1].Text = ddlPipeline.SelectedItem.Text + "  " + "Pipeline";
                        sheet.Range[3, 1].Text = strCombName;
                        break;
                    //case 1:
                    //    sheet.Range[5, 1].Text = "Partial or Full Withdrawls - Cross";
                    //    break;
                    //case 2:
                    //    sheet.Range[5, 1].Text = "Partial Withdrawls - No Cross";
                    //    break;
                    //case 3:
                    //    sheet.Range[5, 1].Text = "Full Withdrawls - NO Cross";
                    //    break;
                }

                //string startDate = txtStartDate.Text.Trim() == "" ? "null" : "'" + txtStartDate.Text.Trim() + "'";
               // string endDate = txtEndDate.Text.Trim() == "" ? "null" : "'" + txtEndDate.Text.Trim() + "'";

                string lblDate  = startDate.Replace("null", "").Replace("'", "") + " - " + endDate.Replace("null", "").Replace("'", "");
                if (txtStartDate.Text == "" && txtEndDate.Text == "")
                {
                    sheet.Range[4, 1].Text = "all dates";
                }
                else
                {
                    sheet.Range[4, 1].Text = lblDate;
                }
                if (lblDate.StartsWith("- ") || lblDate.EndsWith(" -"))
                    sheet.Range[4, 1].Text = lblDate.Replace("-", "");

                //ssheet.Range[4, 1].Text = DateTime.Now.ToString("MMM dd, yyyy");
                sheet.Range[4, 1].Style.Font.IsItalic = true;

                sheet.Range[4, 1].Style.Font.IsBold = false;
                sheet.Range[4, 1].Style.HorizontalAlignment = HorizontalAlignType.Center;
                //sheet.Range[7, 1, DatarowCount, 5].RowHeight = 25;
                //sheet.Range[7, 1, DatarowCount, 5].Style.Font.Size = 10;
                //////////////// new

                //  sheet.Range[6, 1, 500, 1].ColumnWidth = 35;

                if (lodataset.Tables.Count > 0)
                {
                    for (int liCounter = 0; liCounter < lodataset.Tables[sheetNo].Rows.Count; liCounter++)
                    {
                        int lisrc = liCounter + 7;
                        int liColumnHigeshWidth = 0;
                        //for (int liColumns = 2; liColumns < loInsertdataset.Tables[sheetNo].Columns.Count; liColumns++)
                        //{
                        //    try
                        //    {
                        //        if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && !sheet.Range[lisrc, liColumns].Text.Contains("%"))
                        //        {
                        //            if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                        //                sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                        //            sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                        //            sheet.Range[lisrc, liColumns].NumberFormat = "#,##0_);[Black]\\(#,##0\\)";
                        //        }
                        //        /* if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && sheet.Range[lisrc, liColumns].Text.Contains("%"))
                        //         {
                        //             sheet.Range[lisrc, liColumns].Text = sheet.Range[lisrc, liColumns].Text.Replace("%", "");
                        //             if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                        //                 sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                        //             sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text) / 100);
                        //             sheet.Range[lisrc, liColumns].NumberFormat = "#,##0_);[Black]\\(#,##0\\)";// "$#,##0.0%_);\\($#,##0.0%\\)";
                        //         }*/




                        //    }
                        //    catch
                        //    {
                        //        //Response.Write("<br>Error: " + lisrc + "  " + liColumns + " " + sheet.Range[lisrc, liColumns].Text);
                        //    }



                        //}

                        ///// wrap text

                        if (liCounter > 6)
                        {
                            ///////////

                            if (sheet.Range[liCounter, 9].Value.Length <= 65)
                            {

                                sheet.Range[liCounter, 9].RowHeight = 16.5;
                            }
                            else if (sheet.Range[liCounter, 9].Value.Length > 65 && sheet.Range[liCounter, 9].Value.Length <= 130)
                            {
                                sheet.Range[liCounter, 9].RowHeight = 32.10;
                            }
                            else if (sheet.Range[liCounter, 9].Value.Length > 130 && sheet.Range[liCounter, 9].Value.Length <= 195)
                            {
                                sheet.Range[liCounter, 9].RowHeight = 48.15;
                            }
                            else if (sheet.Range[liCounter, 9].Value.Length > 195 && sheet.Range[liCounter, 9].Value.Length <= 260)
                            {
                                sheet.Range[liCounter, 9].RowHeight = 64.20;
                            }

                            if (sheet.Range[liCounter, 5].Value.Length <= 65 && sheet.Range[liCounter, 5].RowHeight < 16.5)
                            {
                                sheet.Range[liCounter, 5].RowHeight = 16.5;
                            }
                            else if (sheet.Range[liCounter, 5].Value.Length > 65 && sheet.Range[liCounter, 5].Value.Length <= 130 && sheet.Range[liCounter, 5].RowHeight < 32.10)
                            {
                                sheet.Range[liCounter, 5].RowHeight = 32.10;
                            }
                            else if (sheet.Range[liCounter, 5].Value.Length > 130 && sheet.Range[liCounter, 5].Value.Length <= 195 && sheet.Range[liCounter, 5].RowHeight < 48.15)
                            {
                                sheet.Range[liCounter, 5].RowHeight = 48.15;
                            }
                            else if (sheet.Range[liCounter, 5].Value.Length > 195 && sheet.Range[liCounter, 5].Value.Length <= 260 && sheet.Range[liCounter, 5].RowHeight < 64.20)
                            {
                                sheet.Range[liCounter, 5].RowHeight = 64.20;
                            }

                            if (sheet.Range[liCounter, 9].Text != null)
                            {
                                string[] CountNewLine = sheet.Range[liCounter, 9].Text.Split('\r');
                                sheet.Range[liCounter, 9].RowHeight = sheet.Range[liCounter, 9].RowHeight + (18.5 * (CountNewLine.Length - 1));
                            }

                        }

                    }
                

             


                /////////////// new1 


                //sheet.Range[5, 1, 5, 1].Style.Color = System.Drawing.Color.FromArgb(112, 128, 144);

                /* ---------------NEW LOGIC TEST-------------*/
                //int lisrc = 0; 
               // sheet.Range[6, 1, 6, 5].Style.Font.Color = System.Drawing.Color.White;

                for (int liCounter = 0; liCounter < lodataset.Tables[sheetNo].Rows.Count; liCounter++)
                {
                    int lisrc = liCounter + 7;

                    if (liCounter == 0)
                    {
                        for (int liColumns = 1; liColumns <= loInsertdataset.Tables[sheetNo].Columns.Count; liColumns++)
                        {
                            //Header Setting           

                            sheet.Range[6, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                            sheet.Range[6, liColumns].Style.Font.Size = 11;
                            sheet.Range[6, liColumns].RowHeight = 45;
                            sheet.Range[6, liColumns].VerticalAlignment = VerticalAlignType.Center;
                            sheet.Range[6, liColumns].Style.Font.IsBold = true;
                            sheet.Range[6, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Center;
                            sheet.Range[6, liColumns].IsWrapText = true;
                            sheet.Range[6, liColumns].Style.Color = System.Drawing.Color.Gray; //System.Drawing.Color.FromArgb(216, 216, 216);

                        }
                    }

                    //if (Convert.ToString(lodataset.Tables[0].Rows[liCounter]["_ordernmb"]).ToLower() == "4") //Rank
                    //{
                    //    sheet.Range[lisrc, 1, lisrc, 9].Style.Color = System.Drawing.Color.White; //System.Drawing.Color.FromArgb(216, 216, 216);
                    //}

                    ///// Set Header Formating _typeflg

                    //Contact Data

                    
                    if (Convert.ToString(lodataset.Tables[0].Rows[liCounter]["_ordernmb"]).ToLower() == "0") //Rank
                    {
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Color = System.Drawing.Color.Brown; //System.Drawing.Color.FromArgb(216, 216, 216);
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.Color = System.Drawing.Color.Black;
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.Size = 12;
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.IsBold = true;
                        sheet.Range[lisrc, 1, lisrc, 9].RowHeight = 15;

                    } // Contact
                    else if (Convert.ToString(lodataset.Tables[0].Rows[liCounter]["_ordernmb"]).ToLower() == "1")
                    {
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Color = System.Drawing.Color.FromArgb(216, 216, 216);  //System.Drawing.Color.DarkGray;
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.Color = System.Drawing.Color.Black; //System.Drawing.Color.FromName("#B0B0B0");
                        
                        sheet.Range[lisrc, 1, lisrc, 9].RowHeight = 15;
                        sheet.Range[lisrc, 4].Style.Font.Size = 10;
                        sheet.Range[lisrc, 2].Style.Font.IsBold = true;
                        
                    }//Rank Total 
                    else if (Convert.ToString(lodataset.Tables[0].Rows[liCounter]["_ordernmb"]).ToLower() == "3" )
                    {
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Color = System.Drawing.Color.Brown;//System.Drawing.Color.FromArgb(216, 216, 216);
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.Color = System.Drawing.Color.Black;
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.IsBold = true;
                        if (txtStartDate.Text != "" || txtEndDate.Text != "")
                        {
                            sheet.Range[lisrc, 3].Text = sheet.Range[lisrc, 3].Text.Replace("all dates", txtStartDate.Text + "-" + txtEndDate.Text);
                        }
                    } //Final Total
                    else if(Convert.ToString(lodataset.Tables[0].Rows[liCounter]["_ordernmb"]).ToLower() == "5")
                    {
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Color = System.Drawing.Color.FromArgb(216, 216, 216);  //System.Drawing.Color.FromArgb(216, 216, 216);
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.Color = System.Drawing.Color.Black;
                        sheet.Range[lisrc, 1, lisrc, 9].Style.Font.IsBold = true;
                        if (txtStartDate.Text != "" || txtEndDate.Text != "")
                        {
                            sheet.Range[lisrc, 3].Text = sheet.Range[lisrc, 3].Text.Replace("all dates", txtStartDate.Text + "-" + txtEndDate.Text);
                        }
                    }
                    else if (Convert.ToString(lodataset.Tables[0].Rows[liCounter]["_ordernmb"]).ToLower() == "4") //Rank
                    {
                        sheet.Range[lisrc, 3].Text ="";
                    }
                    else if (Convert.ToString(lodataset.Tables[0].Rows[liCounter]["_ordernmb"]).ToLower() == "2") //Rank
                    {
                        sheet.Range[lisrc, 8].IsWrapText = true; //To Wrap the text in subject column

                        //To move text to next line if '\r' after word
                        try
                        {
                             // Data Formatting 
                            if (lisrc > 6)
                            {
                                sheet.Range[lisrc, 6, DatarowCount, 6].Style.HorizontalAlignment = HorizontalAlignType.Center;//Date
                                sheet.Range[lisrc, 6, DatarowCount, 6].Style.VerticalAlignment = VerticalAlignType.Top;//Date

                                sheet.Range[lisrc, 7, DatarowCount, 7].Style.HorizontalAlignment = HorizontalAlignType.Center; //Touchpoint Type - BD Type
                                sheet.Range[lisrc, 7, DatarowCount, 7].Style.VerticalAlignment = VerticalAlignType.Top; //Touchpoint Type - BD Type

                                sheet.Range[lisrc, 8, DatarowCount, 8].Style.VerticalAlignment = VerticalAlignType.Top; //Subject

                                sheet.Range[lisrc, 9, DatarowCount, 9].Style.VerticalAlignment = VerticalAlignType.Top;//Owner Details
                            }

                            if (sheet.Range[lisrc, 9].Value != "") // move to next line in excel cell
                            {
                                if (sheet.Range[lisrc, 9].Value.Contains("\r"))
                                {
                                    
                                    string GetText = string.Empty;
                                    string[] NewLine = sheet.Range[lisrc, 9].Text.Split('\r');

                                    for (int i = 0; i < NewLine.Length; i++)
                                    {
                                        if (i == 0)
                                        {
                                            GetText = "=\"" + NewLine[0] + "\"";
                                        }
                                        else
                                        {
                                            GetText = GetText + "&CHAR(10)&" + "\"" + NewLine[i].Replace("\"", "\"\"") + "\"";
                                        }
                                    }
                                    //sheet.Range[11, 9].Formula = "=\"Xy'z:,\"&CHAR(10)&\",ABC\"";
                                    sheet.Range[lisrc, 9].Formula = GetText;

                                   
                                }
                            }
                        }
                        catch (Exception Ex)
                        {

                        }
                    }
                    ///
                }
            }

                // sheet.Range[6, 1, DatarowCount, 5].HorizontalAlignment = HorizontalAlignType.Left; // header row
                //sheet.Range[7, 1, DatarowCount, 5].HorizontalAlignment = HorizontalAlignType.Left; // all data alignment

                //sheet.Range[6, 11, 6, 11].RowHeight = 65; // header row height
                //sheet.Range[7, 11, 5000, 11].RowHeight = 16.5; // all data row height

            if (lodataset.Tables.Count > 0)
            {
                sheet.Range[6, 1, DatarowCount, 1].ColumnWidth = 19.57; //Contact Type Rank
                sheet.Range[6, 2, DatarowCount, 2].ColumnWidth = 19.43; //Contact Name
                sheet.Range[6, 3, DatarowCount, 3].ColumnWidth = 30; //Firm
                sheet.Range[6, 4, DatarowCount, 4].ColumnWidth = 19.43; //Location
                sheet.Range[6, 5, DatarowCount, 5].ColumnWidth = 7; //# 
                sheet.Range[6, 6, DatarowCount, 6].ColumnWidth = 10.71; //Date
                sheet.Range[6, 7, DatarowCount, 7].ColumnWidth = 13; //Touchpoint Type - BD Type
                sheet.Range[6, 8, DatarowCount, 8].ColumnWidth = 39.86; //Subject
                sheet.Range[6, 9, DatarowCount, 9].ColumnWidth = 66.57; //Owner Details
                
                sheet.Range[6, 9, DatarowCount, 9].IsWrapText = true; // Owner Details
            }
                //sheet.Range[6, 5, DatarowCount, 5].IsWrapText = true; // Details

               // sheet.Range[7, 1, DatarowCount, 5].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
               // sheet.Range[7, 1, DatarowCount, 5].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(145, 145, 145);


                //sheet.Range[7, 1, 500, 1].HorizontalAlignment = HorizontalAlignType.Center;

                #endregion
                /**/
            }
            //Save workbook to disk
            // workbook.Save();
            #region Save workbook
            workbook.SaveAsXml(strDirectory2);
            workbook = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(strDirectory2);
            XmlElement businessEntities = xmlDoc.DocumentElement;
            XmlNode loNode = businessEntities.LastChild;
            XmlNode loNode1 = businessEntities.FirstChild;
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
                                            lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&\"Frutiger 55 Roman,Italic\"&10Page &P of &N&R&\"Frutiger 55 Roman,Italic\"&10&KD8D8D8&D,&T";
                                        else
                                            lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&R&\"Frutiger 55 Roman,Italic\"&10&KD8D8D8&D,&T";



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
                                    //if (lxNodess.Attributes[0].InnerText == "#33CCCC")
                                    //    lxNodess.Attributes[0].InnerText = "#B7DDE8";

                                    //if (lxNodess.Attributes[0].InnerText == "#969696")//#33CCCC
                                    //{
                                    //    lxNodess.Attributes[0].InnerText = "#B7DDE8";
                                    //}

                                    if (lxNodess.Attributes[0].InnerText == "#969696")//#C0C0C0
                                    {
                                        lxNodess.Attributes[0].InnerText = "#D8D8D8";
                                    }

                                    if (lxNodess.Attributes[0].InnerText == "#C0C0C0")
                                        lxNodess.Attributes[0].InnerText = "#B7DDE8";

                                    

                                    if (lxNodess.Attributes[0].InnerText == "#993300")
                                        lxNodess.Attributes[0].InnerText = "#D8D8D8";


                                    if (lxNodess.Attributes[0].InnerText == "#808080")
                                        lxNodess.Attributes[0].InnerText = "#808080";

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
                                            lxNodessss.Attributes["ss:Color"].InnerText = "#B7DDE8";
                                        }

                                        //if (lxNodessss.Attributes["ss:Color"].InnerText == "#708090")
                                        //{
                                        //    lxNodessss.Attributes["ss:Color"].InnerText = "#d3d3d3";
                                        //}

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
            loFile.CopyTo(strDirectory, true);
            loFile = null;
            loFile = new FileInfo(strDirectory2);
            loFile.Delete();
            #endregion

        }
        
        Response.Write("<script>");
        lsFileNamforFinalXls = "./ExcelTemplate/" + lsFileNamforFinalXls;
        Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        Response.Write("</script>");
        //Response.Redirect("report.xls");
    }
    protected void gvReport_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            string OrderNmb = DataBinder.Eval(e.Row.DataItem, "_OrderNmb").ToString();

            if (OrderNmb == "0")
            {
                e.Row.BackColor = System.Drawing.Color.LightGray;
            }
            else if (OrderNmb == "1")
            {
                e.Row.BackColor = System.Drawing.Color.LightBlue;
               
            }
            else if (OrderNmb == "2")
            {
                e.Row.BackColor = System.Drawing.Color.White;
                e.Row.Cells[5].VerticalAlign = VerticalAlign.Top;
                e.Row.Cells[6].VerticalAlign = VerticalAlign.Top;
                e.Row.Cells[7].VerticalAlign = VerticalAlign.Top;
                e.Row.Cells[8].VerticalAlign = VerticalAlign.Top;
                e.Row.Cells[7].Wrap = true;

                try
                {

                    if (e.Row.Cells[8].Text != "") // move to next line in excel cell
                    {
                        if (e.Row.Cells[8].Text.Contains("\r"))
                        {

                            string GetText = string.Empty;
                            string[] NewLine = e.Row.Cells[8].Text.Split('\r');

                            for (int i = 0; i < NewLine.Length; i++)
                            {
                                if (i == 0)
                                {
                                    GetText = "" + NewLine[0] + "";
                                }
                                else
                                {
                                    GetText = GetText + "<br/>" + "" + NewLine[i].Replace("\"", "\"\"") + "";
                                }
                            }
                            //sheet.Range[11, 9].Formula = "=\"Xy'z:,\"&CHAR(10)&\",ABC\"";
                            e.Row.Cells[8].Text = GetText;
                        }
                    }
                }
                catch (Exception Ex)
                {

                }
            }
            else if (OrderNmb == "3")
            {
                e.Row.BackColor = System.Drawing.Color.LightGray;
                if (txtStartDate.Text != "" || txtEndDate.Text != "")
                {
                    e.Row.Cells[2].Text = e.Row.Cells[2].Text.Replace("all dates", txtStartDate.Text + "-" + txtEndDate.Text);
                }
            }
            else if (OrderNmb == "4")
            {
                e.Row.Cells[2].Text = "";
                
            }
            else if (OrderNmb == "5")
            {
                e.Row.BackColor = System.Drawing.Color.LightBlue;
                if (txtStartDate.Text != "" || txtEndDate.Text != "")
                {
                    e.Row.Cells[2].Text = e.Row.Cells[2].Text.Replace("all dates", txtStartDate.Text + "-" + txtEndDate.Text);
                }
            }
        }
    }
}
