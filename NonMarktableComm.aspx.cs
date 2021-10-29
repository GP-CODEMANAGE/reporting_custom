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

public partial class NonMarktableComm : System.Web.UI.Page
{
    String sqlstr = string.Empty;
    Boolean fbCheckExcel = false;
    GeneralMethods clsGM = new GeneralMethods();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            mvShowReport.ActiveViewIndex = 0;
            FillDropDownList();
            BindCloseDate();
            BindType();
            BindPartnership();

            //fillHousehold();
        }

    }

    public void BindCloseDate()
    {
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("SP_S_GRESHAM_NON_MARKETABLE_CLOSED_DATE");
        ddlclosedate.Items.Clear();
        ddlclosedate.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlclosedate.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }

    public void BindType()
    {
        sqlstr = "SP_S_GRESHAM_FUND_TYPE  @TypeIdNmb=2";
        clsGM.getListForBindListBox(lstType, sqlstr, "TypeNametxt", "TypeIdNmb");

        lstType.Items.Insert(0, "All");
        lstType.Items[0].Value = "3,9";
        lstType.SelectedIndex = 0;
    }

    public void BindPartnership()
    {
        lstpartnership.Items.Clear();

        string strType = lstType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'";

        sqlstr = "SP_S_GRESHAM_NON_MARKETABLE_FUND @TypeListTxt=" + strType;
        clsGM.getListForBindListBox(lstpartnership, sqlstr, "FundName", "FundId");

        lstpartnership.Items.Insert(0, "All");
        lstpartnership.Items[0].Value = "0";
        lstpartnership.SelectedIndex = 0;
    }

    public string getFinalSp()
    {
        String lsSQL = "";
        if (ddlAllocationGroup.SelectedValue != "")
        {
            lsSQL = "SP_R_Advent_Report_Allocation @AllocationGroupNameTxt='" + Convert.ToString(ddlAllocationGroup.SelectedValue).Replace("'", "''") + "', ";
        }
        else
        {
            lsSQL = "SP_R_Advent_Report_Other";
        }
        lsSQL = lsSQL + " @UUID = '" + System.Guid.NewGuid().ToString() + "'," + "@HouseholdName = '" + ddlHousehold.SelectedItem + "'," + "@EndAsofDate = '" + txtAsofdate.Text + "',";

        if (!String.IsNullOrEmpty(txtpriorperiod.Text))
        {
            lsSQL += "@StartAsofDate = '" + txtpriorperiod.Text + "',";
        }
        else
        {
            lsSQL += "@StartAsofDate = " + "null" + ",";
        }


        lsSQL += "@LookThruDetailTxt = '" + ddlLookthrough.SelectedItem + "'," +
        "@ContactFullNameTxt = '" + ddlContact.SelectedItem + "'," +
        "@VersionTxt = '" + ddlVersion.SelectedItem + "'," +
         "@summaryflgtxt = '" + drpSummary.SelectedItem + "'," +
           "@ReportType = '" + ddlAlignment.SelectedItem + "'," +
        "@ReportGroupFlg = " + ddlReportGroupflag.SelectedValue +
        ",@Report2GroupFlg = " + ddlReportgroupflag2.SelectedValue;
        // Response.Write(lsSQL);
        return lsSQL;
    }

    public void generatesExcelsheets()
    {

        #region Spire License Code
        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion

        object ClosedDT = ddlclosedate.SelectedValue == "0" ? "null" : "'" + ddlclosedate.SelectedValue + "'";
        object AdvisorId = ddlAdvisor.SelectedValue == "0" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object AssociateId = ddlAssociate.SelectedValue == "0" ? "null" : "'" + ddlAssociate.SelectedValue + "'";

        string fundidlst = null;

        for (int i = 0; i < lstpartnership.Items.Count; i++)
        {
            if (lstpartnership.Items[i].Selected)
            {
                if (string.IsNullOrEmpty(fundidlst))
                {
                    fundidlst = lstpartnership.Items[i].Value;
                }
                else
                {
                    fundidlst = fundidlst + "," + lstpartnership.Items[i].Value;
                }
            }
        }

        string TypeList = null;

        for (int i = 0; i < lstType.Items.Count; i++)
        {
            if (lstType.Items[i].Selected)
            {
                if (string.IsNullOrEmpty(TypeList))
                {
                    TypeList = lstType.Items[i].Value;
                }
                else
                {
                    TypeList = TypeList + "," + lstType.Items[i].Value;
                }
            }
        }


        object FundIdListTxt = fundidlst == null || fundidlst == "0" ? "null" : "'" + fundidlst.ToString().Replace("'", "''").Trim() + "'";
        object TypeListTxt = TypeList == null || TypeList == "0" ? "null" : "'" + TypeList.ToString().Replace("'", "''").Trim() + "'";
        String lsSQL = "SP_S_NON_MARKETABLE_COMMITMENT_REPORT @ClosedDT = " + ClosedDT + ",@FundIdListTxt = " + FundIdListTxt + ",@FundTypeIdListTxt = " + TypeListTxt + " , @AdvisorId = " + AdvisorId + ", @AssociateId = " + AssociateId;// "select * from investment_report";// getFinalSp();

        //String lsSQL = "SP_S_NON_MARKETABLE_COMMITMENT_REPORT @ClosedDT = " + ClosedDT + ",@FundIdListTxt = " + FundIdListTxt;// "select * from investment_report";// getFinalSp();

        //String lsSQL = "investment_rpt"; //getFinalSp();

        DB clsDB = new DB();
        DataSet lodataset;
        lodataset = null;
        lodataset = clsDB.getDataSet(lsSQL);

        lodataset = AddRejectedRecomSection(lodataset);

        lodataset = AddTotals(lodataset);

        if (lodataset.Tables[0].Rows.Count < 2)
        {
            lblError.Text = "No Records found.";
            return;
        }


        lodataset.Tables[0].Columns["TOTAL Commitment"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);
        lodataset.Tables[0].Columns["Notes"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);

        DataSet loInsertblankRow = lodataset.Copy();
        lodataset.Tables[0].Clear();
        lodataset.Clear();
        lodataset = null;
        lodataset = loInsertblankRow.Clone();
        int liBlankCounter = 1;

        //loInsertblankRow -- all data with blank rows and _columns
        for (int i = 0; i < loInsertblankRow.Tables[0].Rows.Count; i++)
        {

            if (loInsertblankRow.Tables[0].Rows[i]["_UnderlineFlg"].ToString().ToUpper() == "FALSE" && loInsertblankRow.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "FALSE")
            {
                for (int j = 2; j < loInsertblankRow.Tables[0].Columns.Count; j++)
                {
                    loInsertblankRow.Tables[0].Rows[i][j] = " ";
                }
            }
        }

        //loInsertdataset -- data with columns to be displayed.
        DataSet loInsertdataset = loInsertblankRow.Copy();
        loInsertdataset.Tables[0].Columns.Remove("_BoldFlg");
        loInsertdataset.Tables[0].Columns.Remove("_UnderLineFlg");
        loInsertdataset.Tables[0].Columns.Remove("_Anziano ID");
        loInsertdataset.Tables[0].Columns.Remove("_FundName");

        int liTtrow = 0;

        loInsertdataset.AcceptChanges();

        String lsFileNamforFinalXls = "NonMarkatableCommitements_" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\NonMarkatableCommitements.xlsx");
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);
        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls.Replace("xlsx", "xml"));
        // Response.Write(strDirectory);
        //  string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strDirectory + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
        string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + strDirectory + "';Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";    // change by abhi 11/07/2017
        DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");

        FileInfo loFile = new FileInfo(strDirectory1);
        loFile.CopyTo(strDirectory, true);

        String lsFirstColumn = "Insert into [Sheet1$] (";
        String lsFiled = "";
        String lsFieldvalue = "";
        for (int liColumns = 0; liColumns < loInsertdataset.Tables[0].Columns.Count; liColumns++)
        {
            lsFieldvalue += "'" + loInsertdataset.Tables[0].Columns[liColumns].ColumnName.Replace("'", "''") + "'";
            lsFiled += "id" + (liColumns + 1);
            if (liColumns < loInsertdataset.Tables[0].Columns.Count - 1)
            {
                lsFieldvalue = lsFieldvalue + ",";
                lsFiled = lsFiled + ",";
            }
        }
        lsFirstColumn = lsFirstColumn + lsFiled + ")" + " Values (" + lsFieldvalue + ")";
        #region not used
        //using (DbConnection connection = factory.CreateConnection())
        //{
        //    connection.ConnectionString = connectionString;

        //    using (DbCommand command = connection.CreateCommand())
        //    {
        //        try
        //        {
        //            command.CommandText = lsFirstColumn;
        //            connection.Open();
        //            command.ExecuteNonQuery();
        //            connection.Close();
        //        }
        //        catch
        //        {
        //            //Response.Write(lsFirstColumn);
        //        }
        //    }
        //}
        ////loInsertdataset = loInsertblankRow.Copy();
        //for (int liCounter = 0; liCounter < loInsertdataset.Tables[0].Rows.Count; liCounter++)
        //{

        //    lsFirstColumn = "Insert into [Sheet1$] (";

        //    lsFieldvalue = "";
        //    for (int liColumns = 0; liColumns < loInsertdataset.Tables[0].Columns.Count; liColumns++)
        //    {
        //        //if (liColumns != 0 && !loInsertdataset.Tables[0].Columns[liColumns].ColumnName.Contains("_"))
        //        //{
        //        lsFieldvalue += "'" + loInsertdataset.Tables[0].Rows[liCounter][liColumns].ToString().Replace("'", "''") + "'";
        //        if (liColumns < loInsertdataset.Tables[0].Columns.Count - 1)
        //        {
        //            lsFieldvalue = lsFieldvalue + ",";
        //        }
        //        //}
        //    }
        //    lsFirstColumn = lsFirstColumn + lsFiled + ")" + " Values (" + lsFieldvalue + ")";
        //    using (DbConnection connection = factory.CreateConnection())
        //    {
        //        connection.ConnectionString = connectionString;

        //        using (DbCommand command = connection.CreateCommand())
        //        {
        //            try
        //            {
        //                command.CommandText = lsFirstColumn;
        //                //  Response.Write(lsFirstColumn);
        //                connection.Open();
        //                command.ExecuteNonQuery();
        //                connection.Close();
        //            }
        //            catch
        //            {
        //                //Response.Write(lsFirstColumn);
        //                //Response.End();
        //            }
        //        }
        //    }
        //}
        #endregion

        Workbook workbooknew = new Workbook();
        workbooknew.LoadFromFile(strDirectory);

        Worksheet sheetnew = workbooknew.Worksheets[0];
        sheetnew.InsertDataTable(loInsertdataset.Tables[0], true, 6, 1);

        workbooknew.SaveToFile(strDirectory);

        if (1 == 1)
        {
            #region StyleUsing Spire.xls
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(strDirectory);

            //Gets first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            //Worksheet sheetCover = workbook.Worksheets[0];
            sheet.PageSetup.TopMargin = 0.25;
            sheet.Range["A2"].Text = "Non-Marketable Commitments";
            sheet.Range["A2"].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A2"].Style.Font.Size = 20;
            sheet.Range["A2"].RowHeight = 30;

            //sheet.Range["A3"].Text = "Investment Team Report";
            //sheet.Range["A3"].Style.Font.IsBold = true;
            if (!Convert.ToString(FundIdListTxt).Contains(",") && Convert.ToString(FundIdListTxt) != "null")
                sheet.Range["A3"].Text = lstpartnership.SelectedItem.Text;

            sheet.Range["A3"].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A3"].Style.Font.FontName = "Frutiger 55 Roman";
            sheet.Range["A3"].Style.Font.IsBold = true;
            sheet.Range["A3"].Style.Font.Size = 20;
            sheet.Range["A3"].RowHeight = 30;


            sheet.GridLinesVisible = false;
            //remove header
            for (int liRemoveheader = 1; liRemoveheader < 23; liRemoveheader++)
            {
                sheet.Range[1, liRemoveheader].Text = "";
            }

            sheet.Range[5, 1, 5, loInsertdataset.Tables[0].Columns.Count].Style.Interior.Color = System.Drawing.Color.FromArgb(49, 132, 155);
            sheet.Range[5, 1, 5, loInsertdataset.Tables[0].Columns.Count].RowHeight = 44.10;
            lodataset = loInsertblankRow.Copy();// all data with blank rows and _columns

            /*------added by ME-------*/
            //lodataset = loInsertblankRow.Copy();
            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 7;

                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    if (liColumns != 1 && liColumns != loInsertdataset.Tables[0].Columns.Count && !String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text))
                    {
                        try
                        {
                            if (!sheet.Range[lisrc, liColumns].Text.Contains("E"))
                            {
                                sheet.Range[lisrc, liColumns].Text = Convert.ToString(Math.Round(Convert.ToDecimal(sheet.Range[lisrc, liColumns].Text), 2));
                            }
                            else
                            {
                                sheet.Range[lisrc, liColumns].Text = Convert.ToString(Math.Round(Convert.ToDecimal(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text))));
                            }

                        }
                        catch
                        {
                            //Response.Write(sheet.Range[lisrc, liColumns].Text);
                        }
                    }
                    //Header Setting           
                    if (liCounter == 0)
                    {
                        sheet.Range[6, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                        sheet.Range[6, liColumns].Style.Font.Size = 9;
                        sheet.Range[6, liColumns].RowHeight = 35;
                        sheet.Range[6, liColumns].VerticalAlignment = VerticalAlignType.Center;
                        sheet.Range[6, liColumns].Style.Font.IsBold = true;
                        sheet.Range[6, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                    }

                    sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                    sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                    sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                    sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                    sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = Color.FromArgb(216, 216, 216);

                    //if (liColumns != 1 || liColumns != 2 || liColumns != 5)
                    //    sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Right;
                    sheet.Range[lisrc, liColumns].VerticalAlignment = VerticalAlignType.Center;
                }

            }

            //      workbook.SaveToFile(strDirectory, ExcelVersion.Version2016);

            //sheet.Range[6, 1, 500, 1].ColumnWidth = 35;
            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 7;
                int liColumnHigeshWidth = 0;
                for (int liColumns = 2; liColumns < loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    try
                    {
                        if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && !sheet.Range[lisrc, liColumns].Text.Contains("%"))
                        {
                            if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                                sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                            sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                            sheet.Range[lisrc, liColumns].NumberFormat = "$#,##0_);[Black]\\($#,##0\\)";
                        }
                        if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && sheet.Range[lisrc, liColumns].Text.Contains("%"))
                        {
                            sheet.Range[lisrc, liColumns].Text = sheet.Range[lisrc, liColumns].Text.Replace("%", "");
                            if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                                sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                            sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text) / 100);
                            sheet.Range[lisrc, liColumns].NumberFormat = "$#,##0_);[Black]\\($#,##0\\)";// "$#,##0.0%_);\\($#,##0.0%\\)";
                        }
                        if (!String.IsNullOrEmpty(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text))
                        {
                            if (sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Contains("("))
                                sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Replace("(", "").Replace(")", ""))).ToString();
                            sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberValue = Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text);
                            //sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "#,##0.0_);\\(#,##0.0\\)";
                            if (ddlAlignment.SelectedItem.ToString() == "Horizontal")
                                sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "$#,##0.0_);[Black]\\($#,##0.0\\)";

                            else

                                // Response.Write("ll");
                                sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "$#,##0_);[Black]\\($#,##0\\)";
                            //"&quot;$&quot;#,##0_);[Red]\(&quot;$&quot;#,##0\)" ;
                            //"$#,##0_);\\($#,##0\\)";


                        }
                    }
                    catch
                    {
                        //Response.Write("<br>Error: " + lisrc + "  " + liColumns + " " + sheet.Range[lisrc, liColumns].Text);
                    }
                }
            }

            sheet.Range[7, 1, sheet.Rows.Length, sheet.Columns.Length].Style.Font.Size = 12;
            sheet.Range[6, 1, sheet.Rows.Length, sheet.Columns.Length].Style.Font.FontName = "Frutiger 55 Roman";
            sheet.Range[7, 1, sheet.Rows.Length, sheet.Columns.Length].RowHeight = 19.50;

            workbook.SaveToFile(strDirectory, ExcelVersion.Version2016);


            /* ---------------NEW LOGIC TEST-------------*/
            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 7;
                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    //Header Setting           
                    if (liCounter == 0)
                    {
                        sheet.Range[6, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                        sheet.Range[6, liColumns].Style.Font.Size = 14;
                        sheet.Range[6, liColumns].RowHeight = 42;
                        sheet.Range[6, liColumns].VerticalAlignment = VerticalAlignType.Center;
                        sheet.Range[6, liColumns].Style.Font.IsBold = true;
                        sheet.Range[6, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                        sheet.Range[6, liColumns].IsWrapText = true;
                        sheet.Range[6, liColumns].Style.Color = System.Drawing.Color.FromArgb(192, 192, 192);//(216, 216, 216);


                        if (sheet.Range[6, liColumns].Value.Contains("Proposed Amount"))
                        {
                            string newLine = ((char)13).ToString() + ((char)10).ToString();
                            //sheet.Range[5, liColumns].IsWrapText = true;
                            //sheet.Range[5, liColumns, 5, liColumns + 1].Merge(false);
                            // sheet.Merge(sheet.Range[5, liColumns], sheet.Range[5, liColumns + 1]);
                            sheet.Range[5, liColumns].Text = "     Close Date " + sheet.Range[6, liColumns].Value.Substring(0, 10);
                            sheet.Range[5, liColumns].Style.Font.Color = System.Drawing.Color.White;
                            sheet.Range[5, liColumns].Style.Font.Size = 16;
                            sheet.Range[5, liColumns].Style.Font.IsBold = true;
                            //sheet.Merge(sheet.Range[5, liColumns], sheet.Range[5, liColumns + 1]);
                            //sheet.Range[5, liColumns, 5, liColumns + 1].Merge(false);
                        }

                        if (sheet.Range[6, liColumns].Value.Contains("Proposed Amount") || sheet.Range[6, liColumns].Value.Contains("Confirmed Amount"))
                        {
                            sheet.Range[6, liColumns].Text = sheet.Range[6, liColumns].Value.Substring(11, sheet.Range[6, liColumns].Value.Length - 11);
                        }
                    }

                    //main name of fund
                    if (lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString().ToUpper() == "TRUE" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "TRUE")
                    {
                        sheet.Range[lisrc, liColumns].Text = sheet.Range[lisrc, liColumns].Value.ToString() == "NULL" ? " " : sheet.Range[lisrc, liColumns].Value.ToString();
                        sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                        //sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(216, 216, 216);
                        //sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 14;
                        sheet.Range[lisrc, liColumns].Style.Font.IsBold = true;
                        //sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thick;
                        sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                        sheet.Range[lisrc, liColumns].Style.Color = System.Drawing.Color.FromArgb(165, 165, 165);
                        sheet.Range[lisrc, liColumns].RowHeight = 40.25;
                    }

                    //total line
                    // if (lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString() == "" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "TRUE")
                    if (lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "2" || lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "4")
                    {
                        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                        sheet.Range[lisrc, liColumns].Style.Color = System.Drawing.Color.FromArgb(192, 192, 192);//(216, 216, 216);
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 14.0;
                        sheet.Range[lisrc, liColumns].RowHeight = 40.25;
                        sheet.Range[lisrc, liColumns].Style.Font.IsBold = true;
                        sheet.Range[lisrc, 6, lisrc, 6].Text = "";

                        if (lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "2")
                        {
                            sheet.Range[lisrc, 1].Text = "Total";
                            sheet.Range[lisrc, 1].RowHeight = 20;
                            sheet.Range[lisrc, 1, lisrc, 12].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thick;
                            sheet.Range[lisrc, 1, lisrc, 12].Style.Borders[BordersLineType.EdgeTop].Color = Color.Black;
                        }
                    }

                    if (lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString().ToUpper() == "0" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "0" && lodataset.Tables[0].Rows[liCounter]["Notes"].ToString().ToUpper() != "REJECTED RECOMMENDATIONS HEADERROW")
                    {
                        sheet.Range[lisrc, 2, lisrc, sheet.Columns.Length].Text = "";
                        sheet.Range[lisrc, 1].Style.Font.Size = 14;
                        sheet.Range[lisrc, 1].Style.Font.IsBold = true;
                        sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                        sheet.Range[lisrc, liColumns].Style.Color = System.Drawing.Color.FromArgb(181, 221, 232);//(165, 165, 165);
                        sheet.Range[lisrc, liColumns].RowHeight = 20.25;

                    }

                    if (lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString().ToUpper() == "3" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "3")
                    {
                        sheet.Range[lisrc, 2, lisrc, sheet.Columns.Length].Text = "";
                    }

                    if (lodataset.Tables[0].Rows[liCounter]["Notes"].ToString().ToUpper() == "REJECTED RECOMMENDATIONS HEADERROW" && lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString().ToUpper() == "0" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "0")
                    {
                        sheet.Range[lisrc, 2, lisrc, sheet.Columns.Length].Text = "";
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 16;
                        sheet.Range[lisrc, liColumns].Style.Font.IsBold = true;
                        sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(49, 132, 155);
                        sheet.Range[lisrc, liColumns].Style.Color = System.Drawing.Color.FromArgb(49, 132, 155);
                        sheet.Range[lisrc, liColumns].Style.Font.Color = System.Drawing.Color.White;

                        sheet.Range[lisrc, 4, lisrc, 4].Text = "Close Date " + ddlclosedate.SelectedItem.Text;
                        sheet.Range[lisrc, 4, lisrc, 4].Style.Font.Color = System.Drawing.Color.White;
                    }

                    if (sheet.Range[6, liColumns].Value.Contains("Confirmed Amount") || sheet.Range[6, liColumns].Value.Contains("Investing Entity") || sheet.Range[6, liColumns].Value.Contains("TOTAL Commitmen"))
                    {
                        sheet.Range[lisrc, liColumns].Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Medium;
                        sheet.Range[lisrc, liColumns].Borders[BordersLineType.EdgeRight].Color = System.Drawing.Color.Gray;//FromArgb(216, 216, 216);

                    }

                    if (lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString().ToUpper() == "9" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "9")
                    {
                        sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 14;
                        sheet.Range[lisrc, liColumns].RowHeight = 42;
                        sheet.Range[lisrc, liColumns].VerticalAlignment = VerticalAlignType.Center;
                        sheet.Range[lisrc, liColumns].Style.Font.IsBold = true;
                        sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                        sheet.Range[lisrc, liColumns].IsWrapText = true;
                        sheet.Range[lisrc, liColumns].Style.Color = System.Drawing.Color.FromArgb(192, 192, 192);//(216, 216, 216);

                        if (liColumns == 3)
                            sheet.Range[lisrc, liColumns].Text = "Proposed Amount";
                        if (liColumns == 4)
                            sheet.Range[lisrc, liColumns].Text = "Confirmed Amount";
                        if (liColumns == 5)
                            sheet.Range[lisrc, liColumns].Text = "TOTAL Commitment";

                        if (sheet.Range[lisrc, liColumns].Value == "Proposed Amount" || sheet.Range[lisrc, liColumns].Value == "Confirmed Amount")
                        {
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].ColumnWidth = 20.86;
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].HorizontalAlignment = HorizontalAlignType.Right;
                        }

                        if (sheet.Range[lisrc, liColumns].Value == "TOTAL Commitment")
                        {
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].ColumnWidth = 20.86;
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].HorizontalAlignment = HorizontalAlignType.Center;
                        }

                        if (sheet.Range[lisrc, liColumns].Value == "Investing Entity")
                        {
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].ColumnWidth = 60;
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].HorizontalAlignment = HorizontalAlignType.Left;
                        }

                        if (sheet.Range[lisrc, liColumns].Value == "Investment/Household")
                        {
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].ColumnWidth = 48;
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].HorizontalAlignment = HorizontalAlignType.Left;
                        }

                        if (sheet.Range[lisrc, liColumns].Value == "Notes")
                        {
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].ColumnWidth = 61;
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].HorizontalAlignment = HorizontalAlignType.Left;
                            sheet.Range[lisrc, liColumns, lisrc, liColumns].IsWrapText = true;
                        }

                    }

                    if (lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString().ToUpper() == "6" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "6")
                    {
                        sheet.Range[lisrc, 6].Text = "";
                    }

                    if (lodataset.Tables[0].Rows[liCounter]["_UnderlineFlg"].ToString().ToUpper() == "4" && lodataset.Tables[0].Rows[liCounter]["_BoldFlg"].ToString().ToUpper() == "4")
                    {
                        sheet.Range[lisrc, 1].RowHeight = 20;
                        //sheet.Range[lisrc,1,lisrc, 12].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                        //sheet.Range[lisrc, 1, lisrc, 12].Style.Borders[BordersLineType.EdgeTop].Color = Color.FromArgb(0, 0, 0);

                        sheet.Range[lisrc, 1, lisrc, 12].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thick;
                        sheet.Range[lisrc, 1, lisrc, 12].Style.Borders[BordersLineType.EdgeTop].Color = Color.Black;
                    }

                    if (lodataset.Tables[0].Rows[liCounter]["Notes"].ToString().ToUpper() == "BLANKROW")
                    {
                        sheet.Range[lisrc, 12].Text = "";
                    }

                }
                /*DYNAMIC HEIGHT FOR NOTES AS PER CHARACTER COUNT */
                /* if (sheet.Range[lisrc, 5].Text != null)
                 {
                     int charcount = sheet.Range[lisrc, 5].Text.Length;
                     double rowheight = charcount / 50;//FOR COLUMN WIDTH 54.14 , 50 CHAR FIT PROPERLY
                     int finalrowheight = (Convert.ToInt32(rowheight) + 1) * 15;//15 IS HEIGHT FOR 1 ROW
                     sheet.Range[lisrc, 5].RowHeight = finalrowheight;
                 }
                 */
            }

            //  workbook.SaveToFile(strDirectory, ExcelVersion.Version2016);

            //sheet.Range["A6"].HorizontalAlignment = HorizontalAlignType.Left;
            //sheet.Range["B6"].HorizontalAlignment = HorizontalAlignType.Left;
            //sheet.Range["E6"].HorizontalAlignment = HorizontalAlignType.Left;

            for (int k = 1; k <= sheet.Columns.Length; k++)
            {
                if (sheet.Range[6, k].Value == "Proposed Amount" || sheet.Range[6, k].Value == "Confirmed Amount")
                {
                    sheet.Range[6, k, sheet.Rows.Length, k].ColumnWidth = 20.86;
                    sheet.Range[6, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Right;
                }

                if (sheet.Range[6, k].Value == "TOTAL Commitment")
                {
                    sheet.Range[6, k, sheet.Rows.Length, k].ColumnWidth = 20.86;
                    sheet.Range[6, k, 6, k].HorizontalAlignment = HorizontalAlignType.Center;
                }

                if (sheet.Range[6, k].Value == "Investing Entity")
                {
                    sheet.Range[6, k, sheet.Rows.Length, k].ColumnWidth = 60;
                    sheet.Range[6, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Left;
                }

                if (sheet.Range[6, k].Value == "Investment/Household")
                {
                    sheet.Range[6, k, sheet.Rows.Length, k].ColumnWidth = 48;
                    sheet.Range[6, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Left;
                }

                if (sheet.Range[6, k].Value == "Notes")
                {
                    sheet.Range[6, k, sheet.Rows.Length, k].ColumnWidth = 61;
                    sheet.Range[6, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Left;
                    sheet.Range[6, k, sheet.Rows.Length, k].IsWrapText = true;
                }
            }

            //sheet.Range[39,1,39, 12].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            //sheet.Range[39,1,39, 12].Style.Borders[BordersLineType.EdgeTop].Color = Color.FromArgb(0, 0, 0);

            sheet.Range[39, 1, 39, 12].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range[39, 1, 39, 12].Style.Borders[BordersLineType.EdgeTop].Color = Color.Black;

            //sheet.Range[6, 1, 500, 1].ColumnWidth = 41.57;
            //sheet.Range[6, 2, 500, 2].ColumnWidth = 52;
            //sheet.Range[6, 3, 500, 3].ColumnWidth = 20;
            //sheet.Range[6, 4, 500, 4].ColumnWidth = 20;
            //sheet.Range[6, 5, 500, 5].ColumnWidth = 35;
            //sheet.Range[6, 5, 500, 5].IsWrapText = true;



            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 7;
                string val1 = string.Empty;
                try
                {
                     val1 = loInsertdataset.Tables[0].Rows[liCounter][2].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 3].Text = "";
                        sheet.Range[lisrc, 3].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                        sheet.Range[lisrc, 3].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][3].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 4].Text = "";
                        sheet.Range[lisrc, 4].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                        sheet.Range[lisrc, 4].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][4].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 5].Text = "";
                        sheet.Range[lisrc, 5].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                        sheet.Range[lisrc, 5].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][5].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 6].Text = "";
                        sheet.Range[lisrc, 6].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                        sheet.Range[lisrc, 6].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][6].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 7].Text = "";
                        sheet.Range[lisrc, 7].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                        sheet.Range[lisrc, 7].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][7].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 8].Text = "";
                        sheet.Range[lisrc, 8].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                        sheet.Range[lisrc, 8].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][8].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 9].Text = "";
                        sheet.Range[lisrc, 9].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));

                        sheet.Range[lisrc, 9].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][9].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 10].Text = "";
                        sheet.Range[lisrc, 10].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));

                        sheet.Range[lisrc, 10].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][10].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 11].Text = "";
                        sheet.Range[lisrc, 11].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));

                        sheet.Range[lisrc, 11].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][11].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 12].Text = "";
                        sheet.Range[lisrc, 12].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));

                        sheet.Range[lisrc, 12].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }
                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][12].ToString();
                    if (val1 != "")
                    {
                        sheet.Range[lisrc, 13].Text = "";
                        sheet.Range[lisrc, 13].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));

                        sheet.Range[lisrc, 13].NumberFormat = "$#,##0_);[Black]($#,##0)";
                    }
                }
                catch { }



                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][13].ToString();
                    if (val1 != "")
                    {
                        if (val1 == "blankRow")
                        {
                            sheet.Range[lisrc, 14].Text = "";
                        }
                        else if (val1 != "")
                        {

                            sheet.Range[lisrc, 14].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet.Range[lisrc, 14].Text = "";
                            sheet.Range[lisrc, 14].NumberFormat = "$#,##0_);[Black]($#,##0)";
                        }
                    }
                }
                catch { }

                int colcount = loInsertdataset.Tables[0].Columns.Count;

                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][colcount - 2].ToString();
                    if (val1 != "")
                    {
                        if (val1 == "blankRow")
                        {
                            sheet.Range[lisrc, colcount - 1].Text = "";
                        }
                        else if (val1 != "")
                        {

                            sheet.Range[lisrc, colcount - 1].Text = "";
                            sheet.Range[lisrc, colcount - 1].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet.Range[lisrc, colcount - 1].NumberFormat = "$#,##0_);[Black]($#,##0)";
                        }
                    }
                }
                catch { }

                try
                {
                    val1 = loInsertdataset.Tables[0].Rows[liCounter][colcount - 1].ToString();
                    if (val1 != "")
                    {
                        if (val1 == "blankRow")
                        {
                            sheet.Range[lisrc, colcount].Text = "";
                        }
                        else if (val1 != "")
                        {

                            sheet.Range[lisrc, colcount].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet.Range[lisrc, colcount].Text = "";
                            sheet.Range[lisrc, colcount].NumberFormat = "$#,##0_);[Black]($#,##0)";
                        }
                    }
                }
                catch { }



                //try
                //{
                //    val1 = loInsertdataset.Tables[0].Rows[liCounter][14].ToString();
                //    if (val1 != "")
                //    {
                //        if (val1 == "blankRow")
                //        {
                //            sheet.Range[lisrc, 14].Text = "";
                //        }
                //        else
                //        {
                //            sheet.Range[lisrc, 15].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                //            sheet.Range[lisrc, 15].Text = "";
                //            sheet.Range[lisrc, 15].NumberFormat = "$#,##0_);[Black]($#,##0)";
                //        }
                //    }
                //}
                //catch { }

                //try
                //{
                //    val1 = loInsertdataset.Tables[0].Rows[liCounter]["TOTAL Commitment"].ToString();
                //    if (val1 != "")
                //    {

                //            sheet.Range[lisrc, 14].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                //            sheet.Range[lisrc, 14].Text = "";
                //            sheet.Range[lisrc, 14].NumberFormat = "$#,##0_);[Black]($#,##0)";

                //    }
                //}
                //catch { }


            }

            int colcount1 = loInsertdataset.Tables[0].Columns.Count;

            sheet.Range[7, colcount1 - 3, 4000, colcount1 - 1].HorizontalAlignment = HorizontalAlignType.Right;

            sheet.Range[7, 13, 4000, 13].HorizontalAlignment = HorizontalAlignType.Right;

            workbook.SaveToFile(strDirectory, ExcelVersion.Version2016);



            #region not used
            /**/

            ////Save workbook to disk
            //// workbook.Save();
            //workbook.SaveAsXml(strDirectory2);
            //workbook = null;
            //XmlDocument xmlDoc = new XmlDocument();
            //xmlDoc.Load(strDirectory2);
            //XmlElement businessEntities = xmlDoc.DocumentElement;
            //XmlNode loNode = businessEntities.LastChild;
            //XmlNode loNode1 = businessEntities.FirstChild;
            /////    businessEntities.RemoveChild(loNode);      comment becaue of for spire error 

            //foreach (XmlNode lxNode in businessEntities)
            //{
            //    if (lxNode.Name == "ss:Worksheet")
            //    {
            //        foreach (XmlNode lxPagingNode in lxNode.ChildNodes)
            //        {
            //            if (lxPagingNode.Name == "x:WorksheetOptions")
            //            {
            //                foreach (XmlNode lxPagingSetup in lxPagingNode.ChildNodes)
            //                {
            //                    if (lxPagingSetup.Name == "x:PageSetup")
            //                    {
            //                        //  lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&0022Frutiger 55 Roman,Regular0022&8 Page &P of &N &R&0022Frutiger 55 Roman,italic0022&8  &KD8D8D8&D, &T";
            //                        try
            //                        {
            //                            if (!lxNode.Attributes[0].InnerText.ToLower().Contains("cover"))
            //                                lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&\"Frutiger 55 Roman,Italic\"&10Page &P of &N&R&\"Frutiger 55 Roman,Italic\"&10&KD8D8D8&D,&T";
            //                            else
            //                                lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&R&\"Frutiger 55 Roman,Italic\"&10&KD8D8D8&D,&T";



            //                        }
            //                        catch { }
            //                    }
            //                }
            //            }

            //        }
            //    }

            //    if (lxNode.Name == "ss:Styles")
            //    {
            //        foreach (XmlNode lxNodes in lxNode.ChildNodes)
            //        {
            //            try
            //            {

            //                foreach (XmlNode lxNodess in lxNodes.ChildNodes)
            //                {
            //                    if (lxNodess.Name == "ss:Interior")
            //                    {
            //                        if (lxNodess.Attributes[0].InnerText == "#969696")//#33CCCC
            //                        {
            //                            lxNodess.Attributes[0].InnerText = "#B7DDE8";
            //                        }

            //                        if (lxNodess.Attributes[0].InnerText == "#969696")//#C0C0C0
            //                        {
            //                            //lxNodess.Attributes[0].InnerText = "#D8D8D8";
            //                        }
            //                        if (lxNodess.Attributes[0].InnerText == "#008080")//#C0C0C0
            //                        {
            //                            lxNodess.Attributes[0].InnerText = "#31849B";
            //                        }
            //                    }
            //                }

            //                foreach (XmlNode lxNodess in lxNodes.ChildNodes)
            //                {
            //                    if (lxNodess.Name == "ss:Borders")
            //                    {
            //                        foreach (XmlNode lxNodessss in lxNodess.ChildNodes)
            //                        {
            //                            if (lxNodessss.Attributes["ss:Color"].InnerText == "#C0C0C0")
            //                            {
            //                                lxNodessss.Attributes["ss:Color"].InnerText = "#F2F2F2";
            //                            }
            //                        }

            //                    }
            //                }

            //            }
            //            catch
            //            {
            //            }
            //        }
            //    }
            //}

            //xmlDoc.Save(strDirectory2);
            //xmlDoc = null;
            //loFile = null;
            //loFile = new FileInfo(strDirectory);
            //loFile.Delete();
            //loFile = new FileInfo(strDirectory2);
            //loFile.CopyTo(strDirectory, true);
            //loFile = null;
            ////loFile = new FileInfo(strDirectory2);
            ////loFile.Delete();
            #endregion

            #endregion

            #region delete spire.xls Region
            //connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strDirectory + ";Extended Properties=\"Excel 8.0;HDR=No;\"";
            //for (int liExtracounter = 1; liExtracounter < 13; liExtracounter++)
            //{
            //    using (DbConnection connection = factory.CreateConnection())
            //    {
            //        connection.ConnectionString = connectionString;
            //        using (DbCommand command = connection.CreateCommand())
            //        {
            //            command.CommandText = "Update [Evaluation Warning$B" + liExtracounter + ":B" + liExtracounter + "] Set F1=''";
            //            connection.Open();
            //            command.ExecuteNonQuery();



            //            connection.Close();
            //        }
            //    }
            //}
            //connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strDirectory + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            //using (DbConnection connection = factory.CreateConnection())
            //{
            //    connection.ConnectionString = connectionString;
            //    using (DbCommand command = connection.CreateCommand())
            //    {
            //        command.CommandText = "DROP table  [Evaluation Warning$]";
            //        connection.Open();
            //        command.ExecuteNonQuery();
            //        connection.Close();
            //    }
            //}


            #endregion
        }

        #region New xls to xlsx code
        //Workbook workbook1 = new Workbook();
        //workbook1.LoadFromXml(strDirectory2);
        //workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);

        //workbook1 = new Workbook();
        //workbook1.LoadFromFile(strDirectory);
        //Worksheet sheet1 = workbook1.Worksheets[0];
        //workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);


        loFile = new FileInfo(strDirectory2);
        loFile.Delete();
        loFile = null;
        lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + lsFileNamforFinalXls;
        #endregion

        Response.Write("<script>");
        //   lsFileNamforFinalXls = "./ExcelTemplate/" + lsFileNamforFinalXls;
        Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        Response.Write("</script>");
        //Response.Redirect("report.xls");
    }

    private DataSet AddRejectedRecomSection(DataSet lodataset)
    {
        DataSet TestDataset = new DataSet();
        DataTable table = new DataTable();
        if (lodataset.Tables.Count > 1)//&& lodataset.Tables[1].Rows.Count > 1
        {

            for (int i = 0; i < lodataset.Tables.Count; i++)
            {
                if (i % 2 == 0)
                {
                    if (lodataset.Tables[i].Rows.Count > 1)
                    {
                        DataRow dr2 = lodataset.Tables[i].NewRow();
                        dr2[0] = "zzzzzzzzzzzzzzzzz";//                                                                                Close Date " + ddlclosedate.SelectedItem.Text;

                        dr2["_BoldFlg"] = 3;
                        dr2["_UnderLineFlg"] = 3;
                        dr2["Notes"] = "blankRow";
                        //dr2["Notes"] = "REJECTED RECOMMENDATIONS HEADERROW";
                        lodataset.Tables[i].Rows.Add(dr2);
                        lodataset.AcceptChanges();
                        ///////////////////////////////
                        DataRow dr = lodataset.Tables[i].NewRow();
                        dr[1] = "REJECTED RECOMMENDATIONS";//                                                                                Close Date " + ddlclosedate.SelectedItem.Text;

                        dr["_BoldFlg"] = 0;
                        dr["_UnderLineFlg"] = 0;
                        dr["Notes"] = "REJECTED RECOMMENDATIONS HEADERROW";
                        // dr[3] = ;

                        lodataset.Tables[i].Rows.Add(dr);
                        lodataset.AcceptChanges();

                        //////////////////////////////
                        DataRow dr1 = lodataset.Tables[i].NewRow();

                        for (int k = 0; k < lodataset.Tables[i].Columns.Count; k++)
                        {
                            if (lodataset.Tables[i].Columns[k].ColumnName.Contains("Investment/Household"))
                                dr1["Investment/Household"] = "Investment/Household";

                            if (lodataset.Tables[i].Columns[k].ColumnName.Contains("Investing Entity"))
                                dr1["Investing Entity"] = "Investing Entity";

                            if (lodataset.Tables[i].Columns[k].ColumnName.Contains("Proposed Amount"))
                                dr1[k] = 0;

                            if (lodataset.Tables[i].Columns[k].ColumnName.Contains("Confirmed Amount"))
                                dr1[k] = 0;

                            if (lodataset.Tables[i].Columns[k].ColumnName.Contains("TOTAL Commitment"))
                                dr1["TOTAL Commitment"] = 0;
                            if (lodataset.Tables[i].Columns[k].ColumnName.Contains("Notes"))
                                dr1["Notes"] = "Notes";

                            dr1["_BoldFlg"] = 9;
                            dr1["_UnderLineFlg"] = 9;

                        }


                        lodataset.Tables[i].Rows.Add(dr1);
                        lodataset.AcceptChanges();


                        lodataset.Tables[i].Merge(lodataset.Tables[i + 1]);
                        lodataset.AcceptChanges();


                        table.Merge(lodataset.Tables[i]);
                        DataRow dr3 = table.NewRow();
                        dr3[0] = "zzzzzzzzzzzzzzzzz";//                                                                                Close Date " + ddlclosedate.SelectedItem.Text;

                        dr3["_BoldFlg"] = 6;
                        dr3["_UnderLineFlg"] = 6;
                        dr3["Notes"] = "blankRow";
                        table.Rows.Add(dr3);
                        table.AcceptChanges();
                        table.TableName = "tbltest1";

                    }
                }

            }


        }

        TestDataset.Merge(table);
        TestDataset.AcceptChanges();
        return TestDataset;
    }

    private DataSet AddTotals(DataSet lodataset)
    {
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            DataRow dr = lodataset.Tables[0].NewRow();

            for (int j = 0; j < lodataset.Tables[0].Rows.Count; j++)
            {
                for (int k = 0; k < lodataset.Tables[0].Columns.Count; k++)
                {
                    if (Convert.ToString(lodataset.Tables[0].Rows[j]["_BoldFlg"]) == "4" && Convert.ToString(lodataset.Tables[0].Rows[j]["_UnderLineFlg"]) == "4")
                    {
                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("09/01/2011 Proposed Amount"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            if (Convert.ToString(lodataset.Tables[0].Rows[j][k]) != "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }


                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("09/01/2011 Confirmed"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            if (Convert.ToString(lodataset.Tables[0].Rows[j][k]) != "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("TOTAL Commitment"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            if (Convert.ToString(lodataset.Tables[0].Rows[j][k]) != "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }
                    }

                }
            }



            dr["_BoldFlg"] = 2;
            dr["_FundName"] = "Total";
            lodataset.Tables[0].Rows.Add(dr);
            lodataset.AcceptChanges();


        }
        return lodataset;
    }
    public void FillReportflag()
    {
        ddlReportGroupflag.Items.Clear();
        ddlReportgroupflag2.Items.Clear();
        ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("All", "null"));
        ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("All", "null"));

        ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("Yes", "1"));
        ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("Yes", "1"));

        ddlReportGroupflag.Items.Add(new System.Web.UI.WebControls.ListItem("No", "0"));
        ddlReportgroupflag2.Items.Add(new System.Web.UI.WebControls.ListItem("No", "0"));
        ddlReportGroupflag.SelectedValue = "1";
        ddlReportgroupflag2.SelectedValue = "null";

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
        ddlContact.Items.Clear();

        DataSet loDataset = clsDB.getDataSet("sp_r_Household_contact_list @Householdname ='" + ddlHousehold.SelectedItem + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlContact.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }
    public void AllocationGroup()
    {
        DB clsDB = new DB();
        ddlAllocationGroup.Items.Clear();
        ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Select", ""));
        DataSet loDataset = clsDB.getDataSet("SP_S_Advent_Allocation_Group  @Householdname ='" + ddlHousehold.SelectedItem + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"])));
        }

    }

    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        fillContact();
        AllocationGroup();
        fillHouseholdTitle();
    }
    protected void drpAllocationGroupTitle_SelectedIndexChanged(object sender, EventArgs e)
    {
        fillGroupAllocationTitle();
    }
    public void fillGroupAllocationTitle()
    {
        DB clsDB = new DB();
        drpAllocationGroupTitle.Items.Clear();
        if (!String.IsNullOrEmpty(ddlAllocationGroup.SelectedValue))
        {
            //drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", ""));
            //  Response.Write("SP_S_AllocationGroupTitle  @AllocationGroupName ='" + ddlAllocationGroup.SelectedValue.Replace("'", "''") + "'");
            DataSet loDataset = clsDB.getDataSet("SP_S_AllocationGroupTitle  @AllocationGroupName ='" + ddlAllocationGroup.SelectedValue.Replace("'", "''") + "'");
            for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
            {
                drpAllocationGroupTitle.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
            }
        }

    }
    public void fillHouseholdTitle()
    {
        DB clsDB = new DB();
        drpHouseHoldReportTitle.Items.Clear();
        if (!String.IsNullOrEmpty(ddlHousehold.SelectedValue))
        {
            drpHouseHoldReportTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", ""));
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

        lblError.Text = "";
        lblPartnership.Text = "";

        if (RadioButton1.Checked)
        {
            fbCheckExcel = false;
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
    void Page_PreInit(object sender, System.EventArgs args)
    {
        gvReport.SkinID = "gvReportSkin";
    }
    private DataSet GetReportData()
    {
        object ClosedDT = ddlclosedate.SelectedValue == "0" ? "null" : "'" + ddlclosedate.SelectedValue + "'";
        object AdvisorId = ddlAdvisor.SelectedValue == "0" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object AssociateId = ddlAssociate.SelectedValue == "0" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        string fundidlst = null;

        for (int i = 0; i < lstpartnership.Items.Count; i++)
        {
            if (lstpartnership.Items[i].Selected)
            {
                if (string.IsNullOrEmpty(fundidlst))
                {
                    fundidlst = lstpartnership.Items[i].Value;
                }
                else
                {
                    fundidlst = fundidlst + "," + lstpartnership.Items[i].Value;
                }
            }
        }

        string TypeList = null;

        for (int i = 0; i < lstType.Items.Count; i++)
        {
            if (lstType.Items[i].Selected)
            {
                if (string.IsNullOrEmpty(TypeList))
                {
                    TypeList = lstType.Items[i].Value;
                }
                else
                {
                    TypeList = TypeList + "," + lstType.Items[i].Value;
                }
            }
        }



        object TypeListTxt = TypeList == null || TypeList == "0" ? "null" : "'" + TypeList.ToString().Replace("'", "''").Trim() + "'";

        object FundIdListTxt = fundidlst == null || fundidlst == "0" ? "null" : "'" + fundidlst.ToString().Replace("'", "''").Trim() + "'";

        String lsSQL = "SP_S_NON_MARKETABLE_COMMITMENT_REPORT @ClosedDT = " + ClosedDT + ",@FundIdListTxt = " + FundIdListTxt + ",@FundTypeIdListTxt= " + TypeListTxt + ", @AdvisorId = " + AdvisorId + ", @AssociateId = " + AssociateId;// "select * from investment_report";// getFinalSp();
        DB clsDB = new DB();
        DataSet lodataset;
        lodataset = null;
        //  Response.Write(lsSQL);
        lodataset = clsDB.getDataSet(lsSQL);

        lodataset = AddRejectedRecomSection(lodataset);

        lodataset = AddTotals(lodataset);
        if (!Convert.ToString(FundIdListTxt).Contains(",") && Convert.ToString(FundIdListTxt) != "null")
            lblPartnership.Text = lstpartnership.SelectedItem.Text;

        //lodataset.Tables[0].Columns["TOTAL Commitment"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);
        //lodataset.Tables[0].Columns["Notes"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);

        //lodataset.Tables[0].Columns.Remove("_BoldFlg");
        //lodataset.Tables[0].Columns.Remove("_UnderLineFlg");
        //lodataset.Tables[0].Columns.Remove("_Anziano ID");
        //lodataset.Tables[0].Columns.Remove("_FundName"); 


        return lodataset;
    }

    public void Generatereport()
    {
        DataSet lodataset;
        lodataset = GetReportData();

        if (lodataset.Tables[0].Rows.Count < 2)
        {
            lblError.Text = "No Records found.";
            return;
        }

        //lodataset = AddTotals(lodataset);

        mvShowReport.ActiveViewIndex = 1;

        lodataset.Tables[0].Columns["TOTAL Commitment"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);
        lodataset.Tables[0].Columns["Notes"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);

        gvReport.AutoGenerateColumns = false;

        gvReport.Controls.Clear();
        gvReport.Columns.Clear();

        gvReport.DataSource = null;
        gvReport.DataBind();

        foreach (DataColumn dc in lodataset.Tables[0].Columns)
        {
            BoundField newboundfiled = new BoundField();
            newboundfiled.DataField = dc.ColumnName;

            newboundfiled.HeaderText = dc.ColumnName;
            if (dc.ColumnName.Contains("Proposed Amount"))
            {
                newboundfiled.HeaderText = "Proposed</br>Amount";
                //newboundfiled.HeaderStyle.Width = Unit.Point(44);
            }

            if (dc.ColumnName.Contains("Confirmed Amount"))
                newboundfiled.HeaderText = "Confirmed</br>Amount";

            if (dc.ColumnName.Contains("TOTAL Commitment"))
                newboundfiled.HeaderText = "TOTAL</br>Commitment";


            if (dc.ColumnName.Substring(0, 1) != "_")
            {
                gvReport.Columns.Add(newboundfiled);
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
            /*if (dc.ColumnName.Contains("% Assets"))
            {
                newboundfiled.HtmlEncode = false;
                newboundfiled.DataFormatString = "{0:#,###0.0;(#,###0.0)}";
                newboundfiled.HeaderStyle.HorizontalAlign = HorizontalAlign.Right;
                newboundfiled.HeaderStyle.Wrap = false;
                //newboundfiled.ItemStyle.CssClass = "ddcblk";
                newboundfiled.ItemStyle.HorizontalAlign = HorizontalAlign.Right;
            }*/
        }

        //Assing DataSource to GridView.
        /*
        for (int i = 0; i < lodataset.Tables[0].Rows.Count; i++)
        {
            if (lodataset.Tables[0].Rows[i]["_UnderlineFlg"].ToString().ToUpper() == "TRUE" && lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "TRUE")
            {
                lodataset.Tables[0].Rows[i][1] = lodataset.Tables[0].Rows[i][0].ToString();
            }

            if (lodataset.Tables[0].Rows[i]["_UnderlineFlg"].ToString() == "" && lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "TRUE")
            {
                if (i != lodataset.Tables[0].Rows.Count - 1)
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[1] = "";
                    lodataset.Tables[0].Rows.InsertAt(newCustomersRow, i + 1);
                }
            }
        }
        */

        gvReport.DataSource = lodataset;
        gvReport.DataBind();
        // gvReport.Columns[0].Visible = false;

        for (int k = 0; k < gvReport.Columns.Count; k++)
        {
            if (gvReport.HeaderRow.Cells[k].Text.Contains("TOTAL Commitment"))
            {
                gvReport.HeaderRow.Cells[k].HorizontalAlign = HorizontalAlign.Center;
            }
        }

        if (gvReport.Rows.Count > 0)
        {
            lblmessage.Text = "";
            lblmessage.Visible = false;
            gvReport.HeaderStyle.Font.Bold = true;
            gvReport.HeaderStyle.Font.Size = FontUnit.Point(14);
            gvReport.RowStyle.Font.Size = FontUnit.Point(12);
            //gvReport.Columns[1].HeaderStyle.Width = 130; // investment column
            //gvReport.Columns[2].HeaderStyle.Width = 200; // investment entity column
            //gvReport.Columns[5].HeaderStyle.Width = 300; // Notes column
            gvReport.Columns[5].HeaderStyle.HorizontalAlign = HorizontalAlign.Center; // Notes column
            gvReport.HeaderRow.Cells[3].Style.Add("padding-left", "15px");
            gvReport.HeaderRow.Cells[4].Style.Add("padding-left", "15px");
            gvReport.HeaderRow.Cells[5].Style.Add("padding-left", "15px");
            //gvReport.HeaderRow.Cells[6].Style.Add("padding-left", "15px");

            //gvReport.HeaderRow.BorderColor = Color.White;
            gvReport.HeaderRow.BackColor = System.Drawing.Color.LightGray;

            for (int j = 0; j < gvReport.Columns.Count; j++)
            {
                if (gvReport.Columns[j].HeaderText.Contains("Confirmed"))
                {
                    //gvReport.Columns[j].ItemStyle.ItemStyle.Border.BorderStyle = BorderStyle.Solid;
                    //gvReport.Columns[j].ItemStyle.BorderWidth = 2;
                    //gvReport.Columns[j].ItemStyle.BorderColor = System.Drawing.Color.Gray;
                }
            }

            for (int i = 0; i < lodataset.Tables[0].Rows.Count; i++)
            {
                if (lodataset.Tables[0].Rows[i]["_UnderlineFlg"].ToString().ToUpper() == "6" && lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "6")
                {
                    gvReport.Rows[i].Cells[5].Text = "";
                }

                if (lodataset.Tables[0].Rows[i]["_UnderlineFlg"].ToString().ToUpper() == "0" && lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "0")
                {
                    gvReport.Rows[i].BackColor = System.Drawing.Color.LightBlue;
                    gvReport.Rows[i].ForeColor = System.Drawing.Color.LightBlue;
                    gvReport.Rows[i].Cells[0].Font.Bold = true;
                    gvReport.Rows[i].Cells[0].Font.Size = 14;

                    gvReport.Rows[i].Cells[0].ForeColor = System.Drawing.Color.Black;
                }


                if (lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "2" || lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "4")
                {
                    gvReport.Rows[i].BackColor = System.Drawing.Color.Gray;
                    gvReport.Rows[i].Font.Bold = true;
                    gvReport.Rows[i].Font.Size = 12;
                    gvReport.Rows[i].ForeColor = System.Drawing.Color.Black;

                    if (lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "2")
                    {
                        gvReport.Rows[i].Cells[0].Text = "Total";
                    }
                }

                if (lodataset.Tables[0].Rows[i]["_UnderlineFlg"].ToString().ToUpper() == "0" && lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "0" && lodataset.Tables[0].Rows[i]["Notes"].ToString().ToUpper() == "REJECTED RECOMMENDATIONS HEADERROW")
                {
                    gvReport.Rows[i].Font.Bold = true;
                    gvReport.Rows[i].Font.Size = 16;
                    gvReport.Rows[i].BackColor = System.Drawing.Color.FromArgb(49, 132, 155);
                    gvReport.Rows[i].ForeColor = System.Drawing.Color.White;
                    gvReport.Rows[i].Cells[0].ForeColor = System.Drawing.Color.White;
                    //gvReport.Rows[i].Cells[0].ColumnSpan = 2;
                    gvReport.Rows[i].Cells[2].Text = "Close Date " + ddlclosedate.SelectedItem.Text;
                    gvReport.Rows[i].Cells[5].Text = "";
                }

                if (lodataset.Tables[0].Rows[i]["_UnderlineFlg"].ToString().ToUpper() == "9" && lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "9")
                {
                    gvReport.Rows[i].BackColor = System.Drawing.Color.LightGray;
                    gvReport.Rows[i].Font.Bold = true;
                    gvReport.Rows[i].Font.Size = 14;
                    gvReport.Rows[i].ForeColor = System.Drawing.Color.Black;


                    for (int k = 0; k < gvReport.Columns.Count; k++)
                    {

                        if (k == 2)
                        {
                            gvReport.Rows[i].Cells[k].Text = "Proposed</br>Amount";
                            //newboundfiled.HeaderStyle.Width = Unit.Point(44);
                        }

                        if (k == 3)
                            gvReport.Rows[i].Cells[k].Text = "Confirmed</br>Amount";

                        if (k == 4)
                        {
                            gvReport.Rows[i].Cells[k].Text = "TOTAL</br>Commitment";
                            gvReport.Rows[i].Cells[k].HorizontalAlign = HorizontalAlign.Center;
                        }
                    }


                }

                if (lodataset.Tables[0].Rows[i]["_UnderlineFlg"].ToString().ToUpper() == "3" && lodataset.Tables[0].Rows[i]["_BoldFlg"].ToString().ToUpper() == "3")
                {
                    gvReport.Rows[i].Cells[2].Text = "";
                    gvReport.Rows[i].Cells[3].Text = "";
                    gvReport.Rows[i].Cells[4].Text = "";
                    gvReport.Rows[i].Cells[5].Text = "";
                }


            }

            //////// new logic to add heders above header row of gridview


        }
        else
        {
            lblmessage.Text = "No Records found.";
            lblmessage.Visible = true;
            GenerateXLS.Enabled = false;
        }
    }

    protected void btnBack_Click(object sender, EventArgs e)
    {
        mvShowReport.ActiveViewIndex = 0;
        //txtAsofdate.Text = "";
        //txtpriorperiod.Text = "";


    }

    protected void BtnExport_Click(object sender, EventArgs e)
    {
        fbCheckExcel = true;
        Generatereport();
        generatesExcelsheets();
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

        //if (e.Row.RowType == DataControlRowType.Header && ddlclosedate.SelectedValue != "0")
        //{
        //    GridView oGridView = (GridView)sender;
        //    GridViewRow oGridViewRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
        //    oGridViewRow.BorderWidth = Unit.Pixel(0);

        //    String lstitle = "Non-Marketable Commitments";
        //    string closedate = ddlclosedate.SelectedValue;

        //    Table loTable = new Table();
        //    loTable.Width = Unit.Percentage(100);
        //    loTable.HorizontalAlign = HorizontalAlign.Center;
        //    TableCell oTableCell = new TableCell();

        //    TableRow loRow = new TableRow();
        //    TableCell loCell = new TableCell();

        //    loRow = new TableRow();
        //    loRow.Height = Unit.Pixel(25);
        //    loCell = new TableCell();
        //    loCell.Text = lstitle;
        //    loCell.CssClass = "familyname";
        //    loCell.Height = Unit.Pixel(25);
        //    loCell.ColumnSpan = gvReport.Columns.Count - 1;
        //    loCell.HorizontalAlign = HorizontalAlign.Center;
        //    loRow.Cells.Add(loCell);

        //    loTable.Rows.Add(loRow);


        //    // Investment Team Report type heading
        //         loRow = new TableRow();
        //        loRow.Height = Unit.Pixel(25);
        //        loCell = new TableCell();
        //        loCell.Text = "";
        //        loCell.CssClass = "familyname";
        //        loCell.Height = Unit.Pixel(25);
        //        loCell.ColumnSpan = gvReport.Columns.Count - 1;
        //        loCell.HorizontalAlign = HorizontalAlign.Center;
        //        loRow.Cells.Add(loCell);
        //        loTable.Rows.Add(loRow);

        //    //

        //    oTableCell.Controls.Add(loTable);
        //    oTableCell.Height = Unit.Pixel(25);
        //    oTableCell.ColumnSpan = gvReport.Columns.Count - 1;
        //    oTableCell.HorizontalAlign = HorizontalAlign.Center;
        //    oGridViewRow.CssClass = "ht25px";
        //    oGridViewRow.Cells.Add(oTableCell);
        //    oGridView.Controls[0].Controls.AddAt(0, oGridViewRow);


        //}
    }


    public void GetMultiRowHeader(GridViewRowEventArgs e, DataTable GetCels)
    {
        if (e.Row.RowType == DataControlRowType.Header)
        {
            GridViewRow row = default(GridViewRow);
            //IDictionaryEnumerator enumCels = GetCels.GetEnumerator();

            row = new GridViewRow(-1, -1, DataControlRowType.Header, DataControlRowState.Normal);

            //while (enumCels.MoveNext())
            //{
            //    string[] count = enumCels.Value.ToString().Split(Convert.ToChar(","));
            //    TableCell Cell = default(TableCell);
            //    Cell = new TableCell();
            //    Cell.RowSpan = Convert.ToInt16(count[2].ToString());
            //    Cell.ColumnSpan = Convert.ToInt16(count[1].ToString());
            //    Cell.Controls.Add(new LiteralControl(count[0].ToString()));
            //    Cell.HorizontalAlign = HorizontalAlign.Center;
            //    Cell.ForeColor = System.Drawing.Color.White;
            //    Cell.BackColor = System.Drawing.Color.FromArgb(49, 132, 155);
            //    row.Cells.Add(Cell);
            //}

            for (int j = 0; j < GetCels.Rows.Count; j++)
            {

                string[] count = GetCels.Rows[j]["value"].ToString().Split(Convert.ToChar(","));
                TableCell Cell = default(TableCell);
                Cell = new TableCell();
                Cell.RowSpan = Convert.ToInt16(count[2].ToString());
                Cell.ColumnSpan = Convert.ToInt16(count[1].ToString());
                Cell.Controls.Add(new LiteralControl(count[0].ToString()));
                Cell.HorizontalAlign = HorizontalAlign.Center;
                Cell.ForeColor = System.Drawing.Color.White;
                Cell.BackColor = System.Drawing.Color.FromArgb(49, 132, 155);
                row.Cells.Add(Cell);


            }

            e.Row.Parent.Controls.AddAt(0, row);
        }
    }


    ////
    protected void gvReport_RowCreated(object sender, GridViewRowEventArgs e)
    {
        if (1 == 2)
            if (e.Row.RowType == DataControlRowType.Header)
            {
                GridView HeaderGrid = (GridView)sender;
                GridViewRow HeaderGridRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);

                // get data
                object ClosedDT = ddlclosedate.SelectedValue == "0" ? "null" : "'" + ddlclosedate.SelectedValue + "'";
                string fundidlst = null;

                for (int i = 0; i < lstpartnership.Items.Count; i++)
                {
                    if (lstpartnership.Items[i].Selected)
                    {
                        if (string.IsNullOrEmpty(fundidlst))
                        {
                            fundidlst = lstpartnership.Items[i].Value;
                        }
                        else
                        {
                            fundidlst = fundidlst + "," + lstpartnership.Items[i].Value;
                        }
                    }
                }

                string TypeList = null;

                for (int i = 0; i < lstType.Items.Count; i++)
                {
                    if (lstType.Items[i].Selected)
                    {
                        if (string.IsNullOrEmpty(TypeList))
                        {
                            TypeList = lstType.Items[i].Value;
                        }
                        else
                        {
                            TypeList = TypeList + "," + lstType.Items[i].Value;
                        }
                    }
                }



                object TypeListTxt = TypeList == null || TypeList == "0" ? "null" : "'" + TypeList.ToString().Replace("'", "''").Trim() + "'";

                object FundIdListTxt = fundidlst == null ? "null" : "'" + fundidlst.ToString().Replace("'", "''").Trim() + "'";

                String lsSQL = "SP_S_NON_MARKETABLE_COMMITMENT_REPORT @ClosedDT = " + ClosedDT + ",@FundTypeIdListTxt=" + TypeListTxt + ",@FundIdListTxt = " + FundIdListTxt;// "select * from investment_report";// getFinalSp();
                DB clsDB = new DB();
                DataSet lodataset;
                lodataset = null;
                //  Response.Write(lsSQL);
                lodataset = clsDB.getDataSet(lsSQL);

                lodataset.Tables[0].Columns["TOTAL Commitment"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);
                lodataset.Tables[0].Columns["Notes"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);

                lodataset.Tables[0].Columns.Remove("_BoldFlg");
                lodataset.Tables[0].Columns.Remove("_UnderLineFlg");
                lodataset.Tables[0].Columns.Remove("_Anziano ID");
                lodataset.Tables[0].Columns.Remove("_FundName");

                // end of get data


                Table HeaderTable = new Table();
                TableCell HeaderCell = new TableCell();
                TableRow HeaderRow = new TableRow();

                for (int i = 0; i < lodataset.Tables[0].Columns.Count; i++)
                {
                    if (lodataset.Tables[0].Columns[i].ColumnName.Contains("Proposed"))
                    {
                        HeaderCell.Text = "Close Date 1111111";
                        HeaderCell.ColumnSpan = 2;
                        HeaderRow.Cells.Add(HeaderCell);
                        //HeaderGridRow.Cells.Add(HeaderCell);
                    }
                    else
                    {
                        HeaderCell.Text = "---";
                        HeaderCell.ColumnSpan = 1;
                        HeaderRow.Cells.Add(HeaderCell);
                        //HeaderGridRow.Cells.Add(HeaderCell);
                    }

                    HeaderTable.Rows.Add(HeaderRow);

                    //gvReport.Controls[0].Controls.AddAt(0, HeaderGridRow);
                    // gvReport.Controls[0].Controls.AddAt(0, HeaderRow);
                }

                //HeaderCell = new TableCell();
                //HeaderCell.Text = "Employee";
                //HeaderCell.ColumnSpan = 2;
                //HeaderGridRow.Cells.Add(HeaderCell);



            }
    }



    ////
    protected void gvReport_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        ////Everytime you want to add new rows header, you creat new formatcells variable
        //SortedList formatCells = new SortedList();
        ////Format cells format:"
        //// formatCells.Add(<Column number>, <Header Name,number of column to colspan, number of row to rowspan>)

        //formatCells.Add("1", "ROW SPAN,1,2");
        //formatCells.Add("2", "TopGroup,4,1");
        //SortedList formatcells2 = new SortedList();
        //formatcells2.Add("1", "Subgroup1,2,1");
        //formatcells2.Add("2", "Subgroup2,2,1");
        //GetMultiRowHeader(e, formatcells2);
        //GetMultiRowHeader(e, formatCells);


        if (e.Row.RowType == DataControlRowType.Header)
        {
            GridView HeaderGrid = (GridView)sender;
            GridViewRow HeaderGridRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);

            // get data
            object ClosedDT = ddlclosedate.SelectedValue == "0" ? "null" : "'" + ddlclosedate.SelectedValue + "'";
            string fundidlst = null;


            for (int i = 0; i < lstpartnership.Items.Count; i++)
            {
                if (lstpartnership.Items[i].Selected)
                {
                    if (string.IsNullOrEmpty(fundidlst))
                    {
                        fundidlst = lstpartnership.Items[i].Value;
                    }
                    else
                    {
                        fundidlst = fundidlst + "," + lstpartnership.Items[i].Value;
                    }
                }
            }

            string TypeList = null;

            for (int i = 0; i < lstType.Items.Count; i++)
            {
                if (lstType.Items[i].Selected)
                {
                    if (string.IsNullOrEmpty(TypeList))
                    {
                        TypeList = lstType.Items[i].Value;
                    }
                    else
                    {
                        TypeList = TypeList + "," + lstType.Items[i].Value;
                    }
                }
            }



            object TypeListTxt = TypeList == null || TypeList == "0" ? "null" : "'" + TypeList.ToString().Replace("'", "''").Trim() + "'";

            object FundIdListTxt = fundidlst == null || fundidlst == "0" ? "null" : "'" + fundidlst.ToString().Replace("'", "''").Trim() + "'";

            String lsSQL = "SP_S_NON_MARKETABLE_COMMITMENT_REPORT @ClosedDT = " + ClosedDT + ",@FundTypeIdListTxt=" + TypeListTxt + ",@FundIdListTxt = " + FundIdListTxt;// "select * from investment_report";// getFinalSp();
            DB clsDB = new DB();
            DataSet lodataset;
            lodataset = null;
            //  Response.Write(lsSQL);
            lodataset = clsDB.getDataSet(lsSQL);

            lodataset.Tables[0].Columns["TOTAL Commitment"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);
            lodataset.Tables[0].Columns["Notes"].SetOrdinal(lodataset.Tables[0].Columns.Count - 1);

            lodataset.Tables[0].Columns.Remove("_BoldFlg");
            lodataset.Tables[0].Columns.Remove("_UnderLineFlg");
            lodataset.Tables[0].Columns.Remove("_Anziano ID");
            lodataset.Tables[0].Columns.Remove("_FundName");

            // end of get data

            Table HeaderTable = new Table();
            TableCell HeaderCell = new TableCell();
            TableRow HeaderRow = new TableRow();

            //Format cells format:"
            // formatCells.Add(<Column number>, <Header Name,number of column to colspan, number of row to rowspan>)
            SortedList formatCells = new SortedList();
            //Hashtable formatCells = new Hashtable();



            DataTable table = new DataTable();
            DataRow row;
            table.Columns.Add("idx", typeof(string));
            table.Columns.Add("value", typeof(string));

            string blank = "";
            string closeDate = string.Empty;

            for (int i = 0; i < lodataset.Tables[0].Columns.Count; i++)
            {
                if (lodataset.Tables[0].Columns[i].ColumnName.Contains("Proposed"))
                {
                    closeDate = lodataset.Tables[0].Columns[i].ColumnName.Substring(0, 10);
                    formatCells.Add(i.ToString(), "Close Date <br/>" + closeDate + ",2,1");

                    //create new DataRow in our DataTable
                    row = table.NewRow();
                    row["idx"] = i.ToString();
                    row["value"] = "Close Date <br/>" + closeDate + ",2,1";
                    table.Rows.Add(row);

                    i++;
                }
                else
                {
                    formatCells.Add(i.ToString(), ",1,1");

                    row = table.NewRow();
                    row["idx"] = i.ToString();
                    row["value"] = ",1,1";
                    table.Rows.Add(row);
                }

            }

            GetMultiRowHeader(e, table);

        }
    }

    private void FillDropDownList()
    {
        BindAdvisor(ddlAdvisor);
        BindAssociate(ddlAssociate);
    }

    public void BindAdvisor(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_ADVISOR";
        clsGM.getListForBindDDL(ddl, sqlstr, "OwnerIdName", "OwnerId");



        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindAssociate(DropDownList ddl)
    {
        object OwnerId = ddlAdvisor.SelectedValue == "0" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        ddl.Items.Clear();

        sqlstr = "SP_S_ASSOCIATE @OwnerId=" + OwnerId;
        clsGM.getListForBindDDL(ddl, sqlstr, "Ssi_SecondaryOwnerIdName", "Ssi_SecondaryOwnerId");

        if (ddl.Items.Count == 1)
        {
            if (ddl.Items[0].Value == "0")
                ddl.Items.Remove(ddl.Items[0]);
        }
        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;


    }

    protected void ddlAdvisor_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblmessage.Text = "";
        BindAssociate(ddlAssociate);
    }
    protected void lstType_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblmessage.Text = "";
        BindPartnership();
    }
}
