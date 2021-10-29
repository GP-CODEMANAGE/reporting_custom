using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using Spire.Xls;

public partial class frmBillingAUM : System.Web.UI.Page
{
    #region darkblue3 RGB Color

    // string c_darkblue3 = AppLogic.GetParam(AppLogic.ConfigParam.c_darkblue3);

    int c_darkblue3_Red = 204;
    int c_darkblue3_Green = 255;
    int c_darkblue3_Blue = 255;


    #endregion


    #region lightgray RGB Color

    // string c_darkblue3 = AppLogic.GetParam(AppLogic.ConfigParam.c_darkblue3);

    int c_lightgray_Red = 216;
    int c_lightgray_Green = 216;
    int c_lightgray_Blue = 216;


    #endregion


    #region White RGB Color
    int c_white_Red = 255;
    int c_white_Green = 255;
    int c_white_Blue = 255;
    #endregion


    #region Lightblue RGB Color

    //string c_lightblue = AppLogic.GetParam(AppLogic.ConfigParam.c_lightblue);
    int c_lightblue_Red = 218;
    int c_lightblue_Green = 238;
    int c_lightblue_Blue = 243;
    //int c_lightblue_Red = Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.c_lightblue_Red));
    //int c_lightblue_Green = Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.c_lightblue_Green));
    //int c_lightblue_Blue = Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.c_lightblue_Blue));

    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {

        }
    }



    protected void btnGenerateReport_Click(object sender, EventArgs e)
    {
        try
        {
            lblError.Text = "";
            if (txtAsOfDate.Text != "")
                generatesExcelsheets();
            else
                lblError.Text = "Please select date.";
        }
        catch (Exception ex)
        {
            lblError.Text = "Error Occured while fetching report. Details: " + ex.Message;
        }
    }

    private string GetListBoxItem(ListBox lst)
    {
        string text = string.Empty;
        for (int i = 0; i < lst.Items.Count; i++)
        {
            if (lst.Items[i].Selected)
            {
                text = text + "|" + lst.Items[i].Value.Replace("'", "''");
            }
        }
        if (text != "")
            text = text.Remove(0, 1);
        else
            text = "0";

        return text;
    }
    public void generatesExcelsheets()
    {
        #region Spire License Code
        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion

        //  String lsSQL = "SP_R_Adventure_Report @UUID = '" + System.Guid.NewGuid().ToString() + "'";

        String lsSQL = "EXEC SP_R_BILLING_AUM_REPORT @AsOfdate = '" + txtAsOfDate.Text + "'";

        DB clsDB = new DB();
        DataSet lodataset;
        lodataset = null;
        lodataset = clsDB.getDataSet(lsSQL);

        DataSet loInsertblankRow = lodataset.Copy();
        lodataset.Tables[0].Clear();
        lodataset.Clear();
        lodataset = null;
        lodataset = loInsertblankRow.Clone();
        int liBlankCounter = 1;
        DataSet loInsertdataset = null;
        String lsFileNamforFinalXls = "BilllingAUM" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\BilllingAUM.xlsx");
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);
        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls.Replace("xlsx", "xml"));
        FileInfo loFile = null;
        // DataSet loInsertdataset = null;
        //Spire.License.LicenseProvider.SetLicenseKey("RXBfKi8BALJLc4hMAXEUQZAUbU2CX1R+LkUf5JIXi7aZue6hXI5ljcKLtYiQe0Y8fO6LNtmhAAol7LvqkdFm3IiJNsy7tKOrHCXBlMyJ8w00NrFqplYOV1aExL/fDDPJoRh31RzsMFiP/gM8i+uXonNjW9Q+5UGjYr0vMzWE8L6XLtTp65NswSYdO8kYV95m+QOOKfL8JvWOAZr5V4GCCPo4/ESVGW8O5Ciy3XkaM5JrzT+SF4gE+dyeu1DunZJMZf6gashLnOkQS44xLGqQvDTi+w6Yh1xp3sBlCfd0oeW9c2hCVzb8OsYSj5Zfl9oik1ia9y2z58y1fl+NfKKwJSiIkXY2dmwvm44zsyJ+CAhOhh63oiR3Ju5l1riG7ZO7AuGM8twEyUfjETAzfoLWnTtbh2u/KopSFHjpMh3f3NbgfPsvSiRE+UYWGk4gag3ZyLDdcxKnFyADu8NTkyrHnmT+gD1DKiOxbIOQ4S+tq1Ptkv7Lr0iVSgOriwn45p8J3D4mFirXXynOQ6c6RqdN72XGetut2WA7MX/c1YOJHAkwbAH3quB2rrIuOyJIf8H9caTppOhqS7Fj6XwoPxzzT/rE3P+bgjn89l6977+rsCebMLoZkNwAhatVhu+IdvHYhZOsnw64EfeccCVPzwsY9oVgOPmQSyf1+k7aE2AyaFeZmj+ZkIryU+lMjTyaBY/VBgkdhi8Fb05dw217vBoWaECSxu4D6H6ml4ymodegICj2pwFFA//tuwwLUsSVm0XPWfdG4KR6GZjiClE+eNmNJwH9tpmW6gVmqFCdE9h2b3U6FZX+DHwJju3oGlBubz0egnArXxCj/34xXBE9SSYwIyGgTCyLnrho2zZeIe7xZAp3XS2zIC2LJbO10AiOlgChpues7gutpsddbzyr7adW1d8e9L92b26LIifYFQtyX+rFFOUr88B425QINj09P3HzAB2PHXCAW5P6EaS8abqLTYldhl7J8cIUHA0HmWyGNBh1Wv522e4Wz+/XknxEEia3y8YhQmvmfKiTXSu1RdP7dLYgLDP4IHhbeIO+Wx1YSmQ86SZKmqBhwM6vgZnwmoQ4BWO6fXtZK/h0pyp4WlxvYLShuHXN+uuJnFhSKuDrG7S7Qt+yEMCWgjwF06Jx3k4lchKmzpSm8r5GKETQ+0prX9WynirB7GPx0iRrAJtS1aJK0nca3lX12PqZdw3FkN6b9yRg+fugaNZkfluwNVnXEXeifwJOXWRp0UmGfKUrqlNoQDxKEbVzSYHJn9czXC8shbJavgRwUmAOUNkYhTBqrj4BBxeBpB5km6R/zrfMnUPJ0mMVGkxVxE2ivvuaVzrDITWmBx2STeMkAs4E8frg2vBNXeAP22yo+ho9URfvDS7Itq8s3XJ4zuiLNByWYdtTFAKWD9SzqvTfpKKwifp+Sl4upfC6gGylh3Tzyc2KD4XdANT+rlZ2VfpBQlo7DHMPKPgNB37WP4OSNEq1viV8U9JmSrCybZFbqd8+v1h3ygnUOR5wusbC5Q==");
        //Spire.License.LicenseProvider.LoadLicense();
        try
        {
            for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Rows.Count; liBlankRow++)
            {


                //if (liBlankRow != 0 && (loInsertblankRow.Tables[0].Rows[liBlankRow]["_BoldFlg"].ToString() == "True" ))
                //{
                //    String lsdsd = loInsertblankRow.Tables[0].Rows[liBlankRow][0].ToString();


                //    if (!lsdsd.Contains("NET CHANGE %"))
                //    {
                //        // Response.Write(ddlAlignment.SelectedItem.ToString() + "<br>" + lsdsd + "<br>" + loInsertblankRow.Tables[0].Rows[liBlankRow]["_Ssi_Super_BoldFlg"].ToString()+"<br>--------------");


                //        // if (!String.IsNullOrEmpty(txtpriorperiod.Text) )
                //        // {


                //        DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                //        newCustomersRow[0] = "test";

                //        lodataset.Tables[0].Rows.Add(newCustomersRow);
                //        lodataset.AcceptChanges();
                //        liBlankCounter = liBlankCounter + 1;
                //        //  }
                //        // else if (lsdsd.Contains("TOTAL "))
                //        // {



                //        // Response.Write("<br>Ins: " + lsdsd.Contains("TOTAL ") + "<br>-----------------<br>");
                //        // DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                //        //  newCustomersRow[0] = "test";

                //        //   lodataset.Tables[0].Rows.Add(newCustomersRow);
                //        // lodataset.AcceptChanges();
                //        // liBlankCounter = liBlankCounter + 1;
                //        // }


                //    }

                //}
                lodataset.Tables[0].ImportRow(loInsertblankRow.Tables[0].Rows[liBlankRow]); lodataset.AcceptChanges();
            }


            loInsertdataset = lodataset.Copy();
            int liTtrow = 0;
            for (int liNewdataset = 0; liNewdataset < lodataset.Tables[0].Columns.Count; liNewdataset++)
            {
                if (!lodataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") && !lodataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
                {
                    liTtrow = liTtrow + 1;
                }

            }
            for (int liNewdataset = lodataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
            {

                if (lodataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || lodataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
                {
                    loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
                }

            }
            loInsertdataset.AcceptChanges();
            // loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[1]);
            // loInsertdataset.AcceptChanges();

            // Response.Write(strDirectory);
            //  string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strDirectory + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
            // string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + strDirectory + "';Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";    // change by abhi 11/09/2017
            // DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");



            loFile = new FileInfo(strDirectory1);
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
            //            //   Response.Write(lsFirstColumn);
            //        }
            //    }
            //}

            for (int liCounter = 0; liCounter < loInsertdataset.Tables[0].Rows.Count; liCounter++)
            {

                lsFirstColumn = "Insert into [Sheet1$] (";

                lsFieldvalue = "";
                for (int liColumns = 0; liColumns < loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    string strvalue = Convert.ToString(loInsertdataset.Tables[0].Rows[liCounter][liColumns]).Replace("'", "''");
                    if (strvalue == "")
                    {
                        strvalue = "        ";
                    }
                    lsFieldvalue += "'" + strvalue + "'";
                    if (liColumns < loInsertdataset.Tables[0].Columns.Count - 1)
                    {
                        lsFieldvalue = lsFieldvalue + ",";
                    }
                }
                lsFirstColumn = lsFirstColumn + lsFiled + ")" + " Values (" + lsFieldvalue + ")";
                //using (DbConnection connection = factory.CreateConnection())
                //{
                //    connection.ConnectionString = connectionString;

                //    using (DbCommand command = connection.CreateCommand())
                //    {
                //        //if (liCounter == 0 || liCounter == 2)
                //        //{
                //        //    connection.Open();
                //        //    command.CommandText = "INSERT INTO [Sheet1$] (id1) VALUES('')";
                //        //    command.ExecuteNonQuery();
                //        //    connection.Close();
                //        //}
                //        try
                //        {
                //            command.CommandText = lsFirstColumn;
                //            //  Response.Write(lsFirstColumn);
                //            connection.Open();
                //            command.ExecuteNonQuery();
                //            connection.Close();
                //        }
                //        catch
                //        {
                //            // Response.Write(lsFirstColumn);
                //            Response.End();
                //        }
                //    }
                //}
            }

            Workbook workbooknew = new Workbook();
            workbooknew.LoadFromFile(strDirectory);

            Worksheet sheetnew = workbooknew.Worksheets[0];
            sheetnew.InsertDataTable(loInsertdataset.Tables[0], true, 4, 1);

            workbooknew.SaveToFile(strDirectory);


        }
        catch (Exception ex)
        {
            Response.Write("Error in bind" + ex.ToString());
        }
        #region StyleUsing Spire.xls
        Workbook workbook = new Workbook();
        try
        {

            workbook.LoadFromFile(strDirectory);

            //Gets first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            //  Worksheet sheetCover = workbook.Worksheets[0];

            //  String lsfamilyName = ddlHousehold.SelectedItem.Text;
            // int liCommaCounter = lsfamilyName.IndexOf(",");
            // int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
            //if (liCommaCounter > 0 && liSpaceCounter > 0)
            //    lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
            //else
            //    lsfamilyName = lsfamilyName;
            //if (ddlAllocationGroup.SelectedValue != "")
            //{
            //    lsfamilyName = ddlAllocationGroup.SelectedItem.Text;
            //}
            //if (!String.IsNullOrEmpty(drpHouseHoldReportTitle.SelectedValue))
            //    lsfamilyName = drpHouseHoldReportTitle.SelectedValue;

            //if (!String.IsNullOrEmpty(drpAllocationGroupTitle.SelectedValue))
            //    lsfamilyName = drpAllocationGroupTitle.SelectedValue;

            //sheet.Range["A2"].Text = lsfamilyName;
            // sheet.Range["A4"].Text = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";
            sheet.Range["A2"].VerticalAlignment = VerticalAlignType.Center;
            //  if (ddlAlignment.SelectedItem.ToString() != "Horizontal")
            //   sheet.Range["A3"].Text = "ASSET DISTRIBUTION COMPARISON";
            sheet.Range["A3"].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A4"].VerticalAlignment = VerticalAlignType.Center;


            // sheetCover.Range["A21"].Text = lsfamilyName;
            // sheetCover.Range["A23"].Text = Convert.ToDateTime(txtAsofdate.Text).ToString("MMMM dd, yyyy") + "";
            //  sheetCover.Range[1, 1, 500, 1].ColumnWidth = 100;
            //    sheetCover.Range["A21"].RowHeight = 100;
            sheet.Range["A2"].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A3"].Text = "As of " + Convert.ToDateTime(txtAsOfDate.Text).ToString("MM/dd/yyyy") + "";
            //if (ddlAlignment.SelectedItem.ToString() != "Horizontal")
            // sheetCover.Range["K35"].Text = "Asset Distribution Comparison" + ": " + lsfamilyName;
            // else
            //  sheetCover.Range["K35"].Text = "Asset Distribution";
            sheet.GridLinesVisible = false;

            // Merge cells contained in the range.
            sheet.Range["A3:P3"].Merge();
            sheet.Range["A4:W4"].IsWrapText = true;
            sheet.Range["L4:W4"].AutoFitColumns();
            sheet.Range["L4"].ColumnWidth = 22;

            for (int liRemoveheader = 1; liRemoveheader < 24; liRemoveheader++)
            {
                sheet.Range[1, liRemoveheader].Text = "";
            }

            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 5;


                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {

                    //Header Setting                                                         
                    if (liCounter == 0)
                    {
                        sheet.Range[4, liColumns].Style.Font.FontName = "Calibri";
                        //sheet.Rang5, liColumns].Style.Font.Size = 9;
                        sheet.Range[4, liColumns].Style.Font.Size = 11;
                        sheet.Range[4, liColumns].RowHeight = 50;
                        sheet.Range[4, liColumns].VerticalAlignment = VerticalAlignType.Bottom;

                        sheet.Range[4, liColumns].Style.Font.IsBold = true;
                        sheet.Range[4, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(c_lightblue_Red, c_lightblue_Green, c_lightblue_Blue);
                        sheet.Range[4, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Center;

                        sheet.Range[4, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                        sheet.Range[4, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(0, 0, 0);
                        sheet.Range[4, liColumns].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
                        sheet.Range[4, liColumns].Style.Borders[BordersLineType.EdgeRight].Color = System.Drawing.Color.FromArgb(0, 0, 0);
                        sheet.Range[4, liColumns].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                        sheet.Range[4, liColumns].Style.Borders[BordersLineType.EdgeTop].Color = System.Drawing.Color.FromArgb(0, 0, 0);

                        //if (liColumns == 2)
                        //    sheet.Range[4, liColumns].Text = "Billable AUM " + Convert.ToDateTime(txtAsOfDate.Text).ToString("MM/dd/yyyy");
                        //if (liColumns == 4)
                        //    sheet.Range[4, liColumns].Text = "Annual Fee " + Convert.ToDateTime(txtAsOfDate.Text).ToString("MM/dd/yyyy");


                        // sheet.Range[lisrc, liColumns].Style.Interior = System.Drawing.Color.FromArgb(255, 255, 255);
                    }

                    if (lisrc != 4)
                    {
                        sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(c_white_Red, c_white_Green, c_white_Blue);
                        sheet.Range[lisrc, liColumns].Style.Font.FontName = "Calibri";
                        //sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 11;
                        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(0, 0, 0);
                        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
                        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeRight].Color = System.Drawing.Color.FromArgb(0, 0, 0);
                        //sheet.Range[1, 1, 500, 1].ColumnWidth = 25;
                        //sheet.Range[1, 2, 500, 2].ColumnWidth = 15;
                        //sheet.Range[1, 3, 500, 3].ColumnWidth = 22;
                        //sheet.Range[1, 4, 500, 4].ColumnWidth = 12;
                        //sheet.Range[1, 5, 500, 5].ColumnWidth = 12;
                        //sheet.Range[1, 6, 500, 6].ColumnWidth = 32;
                        //sheet.Range[1, 7, 500, 7].ColumnWidth = 19;
                        //sheet.Range[1, 8, 500, 8].ColumnWidth = 19;
                        //sheet.Range[1, 9, 500, 9].ColumnWidth = 20;
                        //sheet.Range[1, 10, 500, 10].ColumnWidth = 15;
                        //sheet.Range[1, 11, 500, 11].ColumnWidth = 15;
                        //sheet.Range[1, 12, 500, 12].ColumnWidth = 15;
                        //sheet.Range[1, 13, 500, 13].ColumnWidth = 15;
                        //sheet.Range[1, 14, 500, 14].ColumnWidth = 15;
                        //sheet.Range[1, 15, 500, 15].ColumnWidth = 40;
                        //sheet.Range[1, 16, 500, 16].ColumnWidth = 15;
                        //sheet.Range[1, 17, 500, 17].ColumnWidth = 15;
                        //sheet.Range[1, 18, 500, 18].ColumnWidth = 15;
                        //sheet.Range[1, 19, 500, 19].ColumnWidth = 15;
                        //sheet.Range[1, 20, 500, 20].ColumnWidth = 15;
                        //sheet.Range[1, 21, 500, 21].ColumnWidth = 15;
                        sheet.Range[1, 1, 500, 1].ColumnWidth = 25;
                        sheet.Range[1, 2, 500, 2].ColumnWidth = 25;
                        sheet.Range[1, 3, 500, 3].ColumnWidth = 22;
                        sheet.Range[1, 4, 500, 4].ColumnWidth = 15;
                        sheet.Range[1, 5, 500, 5].ColumnWidth = 22;
                        sheet.Range[1, 6, 500, 6].ColumnWidth = 12;
                        sheet.Range[1, 7, 500, 7].ColumnWidth = 12;
                        sheet.Range[1, 8, 500, 8].ColumnWidth = 12;
                        sheet.Range[1, 9, 500, 9].ColumnWidth = 32;
                        sheet.Range[1, 10, 500, 10].ColumnWidth = 19;
                        sheet.Range[1, 11, 500, 11].ColumnWidth = 19;
                        sheet.Range[1, 12, 500, 12].ColumnWidth = 20;
                        sheet.Range[1, 13, 500, 13].ColumnWidth = 15;
                        sheet.Range[1, 14, 500, 14].ColumnWidth = 15;
                        sheet.Range[1, 15, 500, 15].ColumnWidth = 15;
                        sheet.Range[1, 16, 500, 16].ColumnWidth = 15;
                        sheet.Range[1, 17, 500, 17].ColumnWidth = 15;
                        sheet.Range[1, 18, 500, 18].ColumnWidth = 40;
                        sheet.Range[1, 19, 500, 19].ColumnWidth = 15;
                        sheet.Range[1, 20, 500, 20].ColumnWidth = 15;
                        sheet.Range[1, 21, 500, 21].ColumnWidth = 15;
                        sheet.Range[1, 22, 500, 22].ColumnWidth = 15;
                        sheet.Range[1, 23, 500, 23].ColumnWidth = 15;
                        sheet.Range[1, 24, 500, 24].ColumnWidth = 15;
                    }



                    //sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(121, 121, 121);

                    if (liColumns == 1 || liColumns == 3 || (liColumns < 3 && liColumns < 9))
                        sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                    else
                        sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Right;

                    sheet.Range[lisrc, liColumns].VerticalAlignment = VerticalAlignType.Bottom;


                }
                if (liCounter == lodataset.Tables[0].Rows.Count - 1)
                {
                    sheet.Range[lisrc, 1, lisrc, 50].Style.Font.IsBold = true;

                    //// sheet.Range[lisrc - 1, 1].Text = " ";
                    //sheet.Range[lisrc, 1].Style.Interior.Color = System.Drawing.Color.FromArgb(218, 238, 243);
                    //sheet.Range[lisrc, 2].Style.Interior.Color = System.Drawing.Color.FromArgb(218, 238, 243);
                    //sheet.Range[lisrc, 3].Style.Interior.Color = System.Drawing.Color.FromArgb(218, 238, 243);
                    //sheet.Range[lisrc, 4].Style.Interior.Color = System.Drawing.Color.FromArgb(218, 238, 243);

                    for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                    {
                        // sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                        //  sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(216, 216, 216);
                        sheet.Range[lisrc, liColumns].Style.Font.FontName = "Calibri";
                        //sheet.Range[lisrc, liColumns].Style.Font.Size = 9;
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 11;
                        // sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;

                    }
                }
                //if (lodataset.Tables[0].Rows[liCounter]["_Ssi_UnderlineFlg"].ToString() != "True" && lodataset.Tables[0].Rows[liCounter]["_Ssi_Super_BoldFlg"].ToString() != "True")
                //{
                //    if (!String.IsNullOrEmpty(Convert.ToString(lodataset.Tables[0].Rows[liCounter][1])))
                //    {
                //        String abc = "          " + lodataset.Tables[0].Rows[liCounter][1].ToString();
                //        sheet.Range[lisrc, 1].Text = abc;
                //    }
                //}
                //if (lodataset.Tables[0].Rows[liCounter]["_Ssi_UnderlineFlg"].ToString() == "True")
                //{
                //    for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                //    {
                //        String abc = "          " + "          " + lodataset.Tables[0].Rows[liCounter][0].ToString();
                //        sheet.Range[lisrc, 1].Text = abc;
                //        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                //        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                //        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeTop].Color = System.Drawing.Color.FromArgb(0, 0, 0);
                //    }

                //}
                //if (lodataset.Tables[0].Rows[liCounter]["_Ssi_Super_BoldFlg"].ToString() == "True")
                //{
                //    for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                //    {
                //        //ExcelColors backKnownColor1 = (ExcelColors)(49);
                //        //  sheet.Range[lisrc, liColumns].Style.Interior.FillPattern = ExcelPatternType.Gradient;
                //        // sheet.Range[lisrc, liColumns].Style.Interior.Gradient.BackKnownColor = backKnownColor1;
                //        // sheet.Range[lisrc, liColumns].Style.Interior.Gradient.ForeKnownColor = backKnownColor1;
                //        //sheet.Range[lisrc, liColumns].Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical;
                //        //  sheet.Range[lisrc, liColumns].Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants4; 
                //        sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(51, 204, 204);
                //        sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";

                //        if (liColumns == 1)
                //        {
                //            //sheet.Range[lisrc, liColumns].Style.Font.Size = 9;
                //            sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                //        }
                //        else
                //        {
                //            //sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                //            sheet.Range[lisrc, liColumns].Style.Font.Size = 7;
                //        }


                //        sheet.Range[lisrc, liColumns].Style.Font.IsBold = true;
                //        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;

                //        sheet.Range[lisrc - 1, 1].Text = "";

                //        sheet.Range[lisrc, liColumns].VerticalAlignment = VerticalAlignType.Center;
                //        sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                //    }
                //}
                //if (lodataset.Tables[0].Rows[liCounter]["_Ssi_TabFlg"].ToString() == "True" && lodataset.Tables[0].Rows[liCounter]["_Ssi_UnderlineFlg"].ToString() != "True")
                //{

                //    String abc = "          " + "          " + lodataset.Tables[0].Rows[liCounter][1].ToString();
                //    sheet.Range[lisrc, 1].Text = abc;



                //}
                //if (lodataset.Tables[0].Rows[liCounter]["_ssi_greylineflg"].ToString() == "True")
                //{
                //    for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                //    {
                //        //sheet.Range[lisrc, liColumns].Style.Font.Color = System.Drawing.Color.FromArgb(165, 165, 165);
                //        sheet.Range[lisrc, liColumns].Style.Font.Color = System.Drawing.Color.FromArgb(99, 99, 99);
                //    }
                //}
                for (int liColumns = 2; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {

                    //  Response.Write("<br>String :"+sheet.Range[lisrc, liColumns].Text+" " + " Colums: " +liColumns+ "  "+loInsertdataset.Tables[0].Columns.Count);
                    if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && liColumns != loInsertdataset.Tables[0].Columns.Count)
                    {
                    }
                    if (liColumns == loInsertdataset.Tables[0].Columns.Count && !String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text))
                    {
                        try
                        { //sheet.Range[lisrc, liColumns].Text = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(sheet.Range[lisrc, liColumns].Text));
                        }
                        catch { }

                    }
                }
            }

            // sheet.Range[3, 1, 500, 1].ColumnWidth = 35   ;
            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 5;
                int liColumnHigeshWidth = 0;
                for (int liColumns = 2; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {

                    try
                    {
                        if (liColumns == 9 || liColumns == 12 || liColumns == 13 || liColumns == 14 || liColumns == 16 || liColumns == 18 || liColumns == 19 || liColumns == 20 || liColumns == 21 || liColumns == 22) // liColumns == 10 ||
                        {
                            //double a = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                            //sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                            //sheet.Range[lisrc, liColumns].NumberFormat = "$ #,##0";
                            // string val = sheet.Range[lisrc, liColumns].Text;
                            // sheet.Range[lisrc, liColumns].Text = "$ " + sheet.Range[lisrc, liColumns].Text;

                        }


                        //if (liColumns == 11 || liColumns == 23)
                            if (liColumns == 12 || liColumns == 24)
                            {
                            //if (liColumns == 3 && liCounter > lodataset.Tables[0].Rows.Count - 4)
                            //{
                            //    sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(191, 191, 191);

                            //}

                            sheet.Range[lisrc, liColumns].Text = sheet.Range[lisrc, liColumns].Text.Replace("%", "");
                            if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                                sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                            sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text) / 100);
                            //  sheet.Range[lisrc, liColumns].NumberFormat = "#,##0.00%_);\\(#,##0.00%\\)";



                        }

                        if (liColumns == 4)
                        {
                            //  sheet.Range[1, 4, 500, 4].Style.Font.IsBold = true;

                            //if (liCounter > lodataset.Tables[0].Rows.Count - 4)
                            //{
                            //    sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                            //    sheet.Range[lisrc, liColumns].NumberFormat = "$ #,##0";

                            //}
                            //else
                            //{
                            //    sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                            //    sheet.Range[lisrc, liColumns].NumberFormat = "#,##0;[RED]-#,##0";
                            //}
                        }


                        //if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && !sheet.Range[lisrc, liColumns].Text.Contains("%"))
                        //{
                        //    if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                        //        sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                        //    sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);

                        //    //sheet.Range[lisrc, liColumns].NumberFormat = "##0.00%";
                        //    sheet.Range[lisrc, liColumns].NumberFormat = "#,##0.00;[RED]-#,##0.00";

                        //}
                        //if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && sheet.Range[lisrc, liColumns].Text.Contains("%"))
                        //{
                        //    sheet.Range[lisrc, liColumns].Text = sheet.Range[lisrc, liColumns].Text.Replace("%", "");
                        //    if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                        //        sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                        //    sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text) / 100);
                        //    sheet.Range[lisrc, liColumns].NumberFormat = "#,##0.0%_);\\(#,##0.0%\\)";
                        //}
                        //if (!String.IsNullOrEmpty(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text))
                        //{
                        //    //  if (sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Contains("("))
                        //    //    sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Replace("(", "").Replace(")", ""))).ToString();
                        //    // sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberValue = Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text);
                        //    //sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "#,##0.0_);\\(#,##0.0\\)";
                        //    //  if (ddlAlignment.SelectedItem.ToString() == "Horizontal")
                        //    //    sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "#,##0.0_);\\(#,##0.0\\)";

                        //    //  else
                        //    if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                        //        sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                        //    sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);

                        //    // Response.Write("ll");
                        //    sheet.Range[lisrc, liColumns].NumberFormat = "#,##0.00;[RED]-#,##0.00";

                        //}
                    }
                    catch
                    {
                        //  Response.Write("<br>Error: " + lisrc + "  " + liColumns + " " + sheet.Range[lisrc, liColumns].Text);
                    }
                }
            }
            sheet.DeleteRow(1, 1);
            sheet.DeleteRow(2, 1);
            sheet.DeleteRow(1, 1);
            //  sheet.DeleteColumn(22, 1);
            sheet.Range[2, 1, lodataset.Tables[0].Rows.Count + 2, 24].RowHeight = 12.75;
            sheet.Range[2, 4, lodataset.Tables[0].Rows.Count + 2, 4].Style.HorizontalAlignment = HorizontalAlignType.Left;

            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                //  int lisrc = liCounter + 7;

                int lisrc = liCounter + 2;
                //string val1 = lodataset.Tables[0].Rows[liCounter][10].ToString();

                string val1 = lodataset.Tables[0].Rows[liCounter][11].ToString();

                if (val1 != "")
                {
                    //sheet.Range[lisrc, 11].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 11].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 11].NumberFormat = "$ #,##0";


                    sheet.Range[lisrc, 12].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 12].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 12].NumberFormat = "$ #,##0";






                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }

               // val1 = lodataset.Tables[0].Rows[liCounter][11].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][12].ToString();

                if (val1 != "")
                {
                    //sheet.Range[lisrc, 12].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 12].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 12].NumberFormat = "$ #,##0";


                    sheet.Range[lisrc, 13].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 13].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 13].NumberFormat = "$ #,##0";




                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }

                try
                {
                    //val1 = lodataset.Tables[0].Rows[liCounter][12].ToString();


                    val1 = lodataset.Tables[0].Rows[liCounter][13].ToString();


                    if (val1 != "")
                    {
                        //sheet.Range[lisrc, 13].Text = "";
                        //double d = Convert.ToDouble(val1);
                        //d = d / 100;

                        //sheet.Range[lisrc, 13].NumberValue = Convert.ToDouble(d);  //Convert.ToString(Convert.ToDouble(val));
                        //sheet.Range[lisrc, 13].NumberFormat = "#,##0.00%_);\\(#,##0.00%\\)";//"#,##0.00%_);(#,##0.00%)";

                        sheet.Range[lisrc, 14].Text = "";
                        double d = Convert.ToDouble(val1);
                        d = d / 100;

                        sheet.Range[lisrc, 14].NumberValue = Convert.ToDouble(d);  //Convert.ToString(Convert.ToDouble(val));
                        sheet.Range[lisrc, 14].NumberFormat = "#,##0.00%_);\\(#,##0.00%\\)";//"#,##0.00%_);(#,##0.00%)";







                        // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                    }
                }
                catch { }

               // val1 = lodataset.Tables[0].Rows[liCounter][13].ToString();
                val1 = lodataset.Tables[0].Rows[liCounter][14].ToString();

                if (val1 != "")
                {
                    //sheet.Range[lisrc, 14].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 14].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 14].NumberFormat = "$ #,##0";

                    sheet.Range[lisrc, 15].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 15].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 15].NumberFormat = "$ #,##0";





                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }

               // val1 = lodataset.Tables[0].Rows[liCounter][14].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][15].ToString();

                if (val1 != "")
                {
                    //sheet.Range[lisrc, 15].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 15].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 15].NumberFormat = "$ #,##0";


                    sheet.Range[lisrc, 16].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 16].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 16].NumberFormat = "$ #,##0";




                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }


                // Adjustment
               // val1 = lodataset.Tables[0].Rows[liCounter][15].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][16].ToString();


                if (val1 != "")
                {
                    //sheet.Range[lisrc, 16].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 16].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 16].NumberFormat = "$ #,##0";


                    sheet.Range[lisrc, 17].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 17].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 17].NumberFormat = "$ #,##0";



                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }

                //Adjusted Quarterly Fee
                //val1 = lodataset.Tables[0].Rows[liCounter][17].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][18].ToString();


                if (val1 != "")
                {
                    //sheet.Range[lisrc, 18].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 18].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 18].NumberFormat = "$ #,##0";


                    sheet.Range[lisrc, 19].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 19].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 19].NumberFormat = "$ #,##0";



                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }

                //Month1
                //val1 = lodataset.Tables[0].Rows[liCounter][19].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][20].ToString();
                

                if (val1 != "")
                {
                    //sheet.Range[lisrc, 20].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 20].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 20].NumberFormat = "$ #,##0";


                    sheet.Range[lisrc, 21].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 21].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 21].NumberFormat = "$ #,##0";



                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }
                //Month2
               // val1 = lodataset.Tables[0].Rows[liCounter][20].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][21].ToString();


                if (val1 != "")
                {
                    //sheet.Range[lisrc, 21].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 21].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 21].NumberFormat = "$ #,##0";

                    sheet.Range[lisrc, 22].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 22].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 22].NumberFormat = "$ #,##0";




                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }
                ////Month3
               // val1 = lodataset.Tables[0].Rows[liCounter][21].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][22].ToString();


                if (val1 != "")
                {
                    //sheet.Range[lisrc, 22].Text = "";
                    ////double d = Convert.ToDouble(val1);
                    //sheet.Range[lisrc, 22].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 22].NumberFormat = "$ #,##0";



                    sheet.Range[lisrc, 23].Text = "";
                    //double d = Convert.ToDouble(val1);
                    sheet.Range[lisrc, 23].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 23].NumberFormat = "$ #,##0";





                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }
               // val1 = lodataset.Tables[0].Rows[liCounter][22].ToString();

                val1 = lodataset.Tables[0].Rows[liCounter][23].ToString();

                if (val1 != "")
                {
                    //sheet.Range[lisrc, 23].Text = "";
                    //double d = Convert.ToDouble(val1);
                    //d = d / 100;
                    //sheet.Range[lisrc, 23].NumberValue = Convert.ToDouble(d);  //Convert.ToString(Convert.ToDouble(val));
                    //sheet.Range[lisrc, 23].NumberFormat = "#,##0.00%_);(#,##0.00%)";// "#,##0.00%_);(#,##0.00%)";


                    sheet.Range[lisrc, 24].Text = "";
                    double d = Convert.ToDouble(val1);
                    d = d / 100;
                    sheet.Range[lisrc, 24].NumberValue = Convert.ToDouble(d);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 24].NumberFormat = "#,##0.00%_);(#,##0.00%)";// "#,##0.00%_);(#,##0.00%)";



                    // sheet1.Range[lisrc, 6].HorizontalAlignment = HorizontalAlignType.Center;
                }
            }


            // sheet.DeleteRow(3, 1);
            //  sheet.DeleteRow(3);
            //Save workbook to disk
            // workbook.Save();
            //     workbook.SaveAsXml(strDirectory2);
            //     workbook = null;


            workbook.SaveToFile(strDirectory, ExcelVersion.Version2016);

        }
        catch (Exception e)
        {
            Response.Write("Error in Exxcel" + e.ToString());
        }
        //try
        //{
        //    XmlDocument xmlDoc = new XmlDocument();
        //    xmlDoc.Load(strDirectory2);
        //    XmlElement businessEntities = xmlDoc.DocumentElement;
        //    XmlNode loNode = businessEntities.LastChild;
        //    //   businessEntities.RemoveChild(loNode);       comment becaue of for spire error 
        //    foreach (XmlNode lxNode in businessEntities)
        //    {
        //        if (lxNode.Name == "ss:Worksheet")
        //        {
        //            foreach (XmlNode lxPagingNode in lxNode.ChildNodes)
        //            {
        //                if (lxPagingNode.Name == "x:WorksheetOptions")
        //                {
        //                    foreach (XmlNode lxPagingSetup in lxPagingNode.ChildNodes)
        //                    {
        //                        if (lxPagingSetup.Name == "x:PageSetup")
        //                        {
        //                            //  lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&0022Frutiger 55 Roman,Regular0022&8 Page &P of &N &R&0022Frutiger 55 Roman,italic0022&8  &KD8D8D8&D, &T";
        //                            try
        //                            {
        //                                if (!lxNode.Attributes[0].InnerText.ToLower().Contains("cover"))
        //                                    lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&\"Frutiger 55 Roman,Regular\"&8Page &P of &N&R&\"Frutiger 55 Roman,Italic\"&8&KD8D8D8&D,&T";
        //                                else
        //                                    lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&R&\"Frutiger 55 Roman,Italic\"&8&KD8D8D8&D,&T";

        //                            }
        //                            catch { }
        //                        }
        //                    }
        //                }

        //            }
        //        }

        //        if (lxNode.Name == "ss:Styles")
        //        {
        //            foreach (XmlNode lxNodes in lxNode.ChildNodes)
        //            {
        //                try
        //                {

        //                    foreach (XmlNode lxNodess in lxNodes.ChildNodes)
        //                    {
        //                        if (lxNodess.Name == "ss:Interior")
        //                        {
        //                            if (lxNodess.Attributes[0].InnerText == "#33CCCC")
        //                                lxNodess.Attributes[0].InnerText = "#B7DDE8";

        //                            if (lxNodess.Attributes[0].InnerText == "#C0C0C0")
        //                                lxNodess.Attributes[0].InnerText = "#D8D8D8";

        //                        }
        //                    }

        //                    foreach (XmlNode lxNodess in lxNodes.ChildNodes)
        //                    {
        //                        if (lxNodess.Name == "ss:Borders")
        //                        {
        //                            foreach (XmlNode lxNodessss in lxNodess.ChildNodes)
        //                            {
        //                                if (lxNodessss.Attributes["ss:Color"].InnerText == "#C0C0C0")
        //                                {
        //                                    //lxNodessss.Attributes["ss:Color"].InnerText = "#F2F2F2";
        //                                    lxNodessss.Attributes["ss:Color"].InnerText = "#D8D8D8";
        //                                }
        //                            }
        //                        }
        //                    }

        //                    foreach (XmlNode lxNodess in lxNodes.ChildNodes)
        //                    {
        //                        if (lxNodess.Name == "ss:Font")
        //                        {

        //                            if (lxNodess.Attributes["ss:Color"].InnerText == "#808080")
        //                            {
        //                                //lxNodessss.Attributes["ss:Color"].InnerText = "#F2F2F2"
        //                                lxNodess.Attributes["ss:Color"].InnerText = "#A5A5A5";
        //                            }

        //                        }
        //                    }
        //                }
        //                catch
        //                {
        //                }
        //            }
        //        }
        //    }

        //    xmlDoc.Save(strDirectory2);
        //    xmlDoc = null;
        //}
        //catch (Exception e)
        //{
        //    Response.Write("Error in XML" + e.ToString());
        //}

        //loFile = null;
        //loFile = new FileInfo(strDirectory);
        //loFile.Delete();
        //loFile = new FileInfo(strDirectory2);
        //loFile.CopyTo(strDirectory, true);
        //loFile = null;
        ////loFile = new FileInfo(strDirectory2);
        ////loFile.Delete();
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

        #region New xls to xlsx code
        //Workbook workbook1 = new Workbook();
        //workbook1.LoadFromXml(strDirectory2);
        //workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);

        Workbook workbook1 = new Workbook();
        workbook1.LoadFromFile(strDirectory);
        Worksheet sheet1 = workbook1.Worksheets[0];
        sheet1.Range[1, 1, 1, 24].Style.Color = System.Drawing.Color.FromArgb(c_darkblue3_Red, c_darkblue3_Green, c_darkblue3_Blue);
        workbook1.SaveToFile(strDirectory, ExcelVersion.Version2010);

        #region PageSetup
        Workbook workbook2 = new Workbook();
        workbook2.LoadFromFile(strDirectory);
        Worksheet sheet2 = workbook2.Worksheets[0];
        workbook2.SaveToFile(strDirectory, ExcelVersion.Version2016);

        var setup = sheet2.PageSetup;
        setup.FitToPagesWide = 1;
        //setup.FitToPagesTall = 1;
        setup.IsFitToPage = true;
        setup.PaperSize = PaperSizeType.PaperA4;
        setup.Orientation = PageOrientationType.Landscape;
        setup.FitToPagesWide = 1;
        setup.FitToPagesTall = 0;
        setup.CenterHorizontally = true;
        setup.CenterVertically = false;
        workbook2.SaveToFile(strDirectory, ExcelVersion.Version2016);
        #endregion
        loFile = new FileInfo(strDirectory2);
        loFile.Delete();
        loFile = null;

        lsFileNamforFinalXls = "/ExcelTemplate/TempFolder/" + lsFileNamforFinalXls;
        #endregion

        Response.Write("<script>");
        // lsFileNamforFinalXls = "./ExcelTemplate/" + lsFileNamforFinalXls;
        Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        Response.Write("</script>");

        string baseUrl = Request.Url.GetLeftPart(UriPartial.Authority);
        Response.Redirect(baseUrl + lsFileNamforFinalXls);
    }
}