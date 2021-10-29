using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Security.Principal;
using System.Data.SqlClient;
using System.Collections;
//using CrmSdk;
using System.IO;
using Spire.Xls;
using System.Data.Common;
using System.Xml;

public partial class NonMarketableSummaryReport : System.Web.UI.Page 
{
    public StreamWriter sw = null;
    bool bProceed = true;
    string strDescription;
    string greshamquery;
    int totalCount = 0;
    int successcount = 0;
    // int successcount = 0;
    GeneralMethods clsGM = new GeneralMethods();

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            //BindSecondaryOwner();
            //BindHousehold();
            BindType();
            BindPartnership();
            //BindGridView("'" + ddlHousehold.SelectedValue + "'");

        }
    }


    public void BindType()
    {
        string sqlstr = "SP_S_GRESHAM_FUND_TYPE  @TypeIdNmb=2";
        clsGM.getListForBindListBox(lstType, sqlstr, "TypeNametxt", "TypeIdNmb");

        lstType.Items.Insert(0, "All");
        lstType.Items[0].Value = "3,9";
        lstType.SelectedIndex = 0;
    }

    public void BindPartnership()
    {
        lstbxPartnership.Items.Clear();

        string strType = lstType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'";

        string sqlstr = "SP_S_GRESHAM_NON_MARKETABLE_FUND @TypeListTxt=" + strType;
        clsGM.getListForBindListBox(lstbxPartnership, sqlstr, "FundName", "FundId");

        lstbxPartnership.Items.Insert(0, "All");
        lstbxPartnership.Items[0].Value = "0";
        lstbxPartnership.SelectedIndex = 0;
    }

    private void BindDropdown(DropDownList ddl, string sqlstr, string TextField, string ValueField)
    {
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);

        ddl.DataTextField = TextField;
        ddl.DataValueField = ValueField;

        ddl.DataSource = ds;
        ddl.DataBind();

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
    }

   

    private DataSet GetReportData()
    {
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery = string.Empty;
        string PartnershipIDListTxt = string.Empty; 

        object AsofDate = txtAsOfDate.Text.Trim().Replace("'", "") == "" ? "null" :"'"+ txtAsOfDate.Text.Trim().Replace("'", "")+"'";

        if (lstType.SelectedValue == "9" && (lstbxPartnership.SelectedValue == "" || lstbxPartnership.SelectedValue == "0"))
        {
            PartnershipIDListTxt = "'" + NotItemsFromListBox(lstbxPartnership) + "'";
        }
        else if (lstType.SelectedValue == "3" && (lstbxPartnership.SelectedValue == "" || lstbxPartnership.SelectedValue == "0"))
        {
            PartnershipIDListTxt = "'" + NotItemsFromListBox(lstbxPartnership) + "'";
        }
        else
        {
            PartnershipIDListTxt = lstbxPartnership.SelectedValue == "0" || lstbxPartnership.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstbxPartnership) + "'";
        }


        try
        {
            greshamquery = "exec [SP_R_NON_MARKETABLE_SUMMARY_REPORT] @AsofDate=" + AsofDate
                                                          + ",@PartnershipIDListTxt=" + PartnershipIDListTxt;
                                                     
            //Response.Write(greshamquery);
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);

        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {

            //lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Detail.InnerText;
        }
        catch (Exception exc)
        {
            //lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Message;
        }

        return ds_gresham;
    }

    protected void gvList_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        //if (e.Row.RowType == DataControlRowType.Header)
        //{
        //    //Find the checkbox control in header and add an attribute
        //    ((CheckBox)e.Row.FindControl("chkbxNCSelectAll")).Attributes.Add("onclick", "javascript:SelectAll('" +
        //            ((CheckBox)e.Row.FindControl("chkbxNCSelectAll")).ClientID + "')");
        //}


        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            DropDownList ddlReportTrackerStatus1 = (DropDownList)e.Row.FindControl("ddlStatus");
            CheckBox chkbxBillingHandedOff1 = (CheckBox)e.Row.FindControl("chkbBillngHandedOff");

            // chkbxNC1.Attributes.Add("onclick", "EnableDisable('" + chkbxNC1.ClientID + "','" + txtCAUpdateValue.ClientID + "')");
            chkbxBillingHandedOff1.Checked = e.Row.Cells[1].Text.ToLower() == "true" ? true : false;
            ddlReportTrackerStatus1.SelectedValue = e.Row.Cells[2].Text == "&nbsp;" ? "0" : e.Row.Cells[2].Text;
        }
    }



    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
      //  CrmService service = null;
        //lblMessage.Text = "";

        try
        {
          //  service = GetCrmService(crmServerUrl, orgName);
            strDescription = "Crm Service starts successfully";
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
            // sw.WriteLine("step 1 ");
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            //lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            //lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }

       
    }


    /// <summary>
    /// Set up the CRM Service.
    /// </summary>
    /// <param name="organizationName">My Organization</param>
    /// <returns>CrmService configured with AD Authentication</returns>
    //public static CrmService GetCrmService(string crmServerUrl, string organizationName)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    // string username = WindowsIdentity.GetCurrent().Name;



    //    CrmService service = new CrmService();

    //    if (crmServerUrl != null &&
    //        crmServerUrl.Length > 0)
    //    {
    //        UriBuilder builder = new UriBuilder(crmServerUrl);
    //        builder.Path = "//MSCRMServices//2007//CrmService.asmx";
    //        service.Url = builder.Uri.ToString();
    //    }

    //    service.CrmAuthenticationTokenValue = token;
    //    service.Credentials = System.Net.CredentialCache.DefaultCredentials;

    //    //////////////////////////// impersonate service to crm user /////////////////////////////

    //    // WhoAmIRequest userRequest = new WhoAmIRequest();
    //    // Execute the request.
    //    // WhoAmIResponse user = (WhoAmIResponse)service.Execute(userRequest);
    //    // string currentuser = user.UserId.ToString();


    //    //string currentuser = "62DE1F95-8203-DE11-A38C-001D09665E8F";
    //    //token.CallerId = new Guid(currentuser);

    //    return service;
    //}
   
    protected void btnSearch_Click(object sender, EventArgs e)
    {
        
    }
    protected void btnBack_Click(object sender, EventArgs e)
    {
       
    }


    protected void btnExportToExcel_Click(object sender, EventArgs e)
    {
        generatesExcelsheets();
    }

    public void generatesExcelsheets()
    {

        #region Spire License Code
        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion


        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        Type tp = this.GetType();

        if (txtAsOfDate.Text == "")
        {
            sb.Append("\n<script type=text/javascript>\n");
            sb.Append("\n alert('Please select Asofdate');");
            sb.Append("</script>");
            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
            return;
        }


        DataSet lodataset = GetReportData();
        
        if (lodataset.Tables[0].Rows.Count < 1)
        {
            lblError.Text = "No Records found.";
            return;
        }


        DataSet loInsertdataset = lodataset.Copy();


        DateTime DtAsOfDate = Convert.ToDateTime(txtAsOfDate.Text.Replace("'", "").Trim());

        DateTime lastDay = new DateTime(DtAsOfDate.Year, DtAsOfDate.Month, 1);// First day of current month
        lastDay = lastDay.AddDays(-1);//get last day of last month

        DateTime DtPriorDate = lastDay;//DtAsOfDate.AddMonths(-1).AddDays(1);// Commented by Rohit

        string strAsOfDate = txtAsOfDate.Text.Replace("'", "").Trim();
        string strPriorDate = DtPriorDate.ToShortDateString();

        //if (Convert.ToString(ViewState["MultipleSelect"]) != "true")
        //{
        //    loInsertdataset.Tables[0].Columns.Remove("Partnership");

        //}


        //loInsertdataset.Tables[0].Columns.Remove("ssi_type");
        //loInsertdataset.Tables[0].Columns.Remove("AssociateID");
        //loInsertdataset.Tables[0].Columns.Remove("Ssi_HouseholdId");
        //loInsertdataset.Tables[0].Columns.Remove("RecipientID");
        //loInsertdataset.Tables[0].Columns.Remove("Send ViaID");
        //loInsertdataset.Tables[0].Columns.Remove("CA UpdateID");
        //loInsertdataset.Tables[0].Columns.Remove("StatusID");
        //loInsertdataset.Tables[0].Columns.Remove("Internal Billing ContactID");


        //loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[16]);
        //loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[17]);
        //loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[18]);
        //loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[19]);
        //loInsertdataset.Tables[0].Columns.Remove("_BoldFlg");
        //loInsertdataset.Tables[0].Columns.Remove("_OrderByColumn");
        //loInsertdataset.Tables[0].Columns.Remove("FundID");

        int liTtrow = 0;

        loInsertdataset.AcceptChanges();

        String lsFileNamforFinalXls ="NonMarketableSummaryReport"+ System.DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\NonMarketableSummaryReport.xlsx");
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
        sheetnew.InsertDataTable(loInsertdataset.Tables[0], true, 5, 1);

        workbooknew.SaveToFile(strDirectory);

        if (1 == 1)
        {
            #region StyleUsing Spire.xls
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(strDirectory);

            //Gets first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            //Worksheet sheetCover = workbook.Worksheets[0];
            //sheet.PageSetup.TopMargin = 0.25;
            sheet.Range["A2"].Text = "Non-Marketable Summary Report";
            sheet.Range["A2"].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A2"].Style.Font.Size = 20;
            sheet.Range["A2"].RowHeight = 30;

            sheet.Range["A3"].Text = txtAsOfDate.Text;
            sheet.Range["A3"].Style.Font.IsItalic = true;
            //if (!Convert.ToString(FundIdListTxt).Contains(",") && Convert.ToString(FundIdListTxt) != "null")
            //    sheet.Range["A3"].Text = lstpartnership.SelectedItem.Text;

            sheet.Range["A3"].VerticalAlignment = VerticalAlignType.Center;
            sheet.Range["A3"].Style.Font.FontName = "Frutiger 55 Roman";
            sheet.Range["A3"].Style.Font.IsBold = false;
            sheet.Range["A3"].Style.Font.Size = 14;
            sheet.Range["A3"].RowHeight = 20;


            sheet.GridLinesVisible = false;
            //remove header
            for (int liRemoveheader = 1; liRemoveheader < 23; liRemoveheader++)
            {
                sheet.Range[1, liRemoveheader].Text = "";
            }

            sheet.Range[5, 1, 5, loInsertdataset.Tables[0].Columns.Count].Style.Interior.Color = System.Drawing.Color.FromArgb(165, 165, 165);
            sheet.Range[5, 1, 5, loInsertdataset.Tables[0].Columns.Count].RowHeight = 44.10;
          //  sheet.Range[5, 1, 5, loInsertdataset.Tables[0].Columns.Count].IsWrapText = false;
           
            //lodataset = loInsertdataset.Copy();// all data with blank rows and _columns

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
                        sheet.Range[5, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                        sheet.Range[5, liColumns].Style.Font.Size = 9;
                        sheet.Range[5, liColumns].RowHeight = 35;
                        sheet.Range[5, liColumns].VerticalAlignment = VerticalAlignType.Center;
                        sheet.Range[5, liColumns].Style.Font.IsBold = true;
                        sheet.Range[5, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                    }

                    //sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                    //sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                    //sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                    //sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                    //sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color =System.Drawing.Color.FromArgb(216, 216, 216);

                    if (liColumns !=1)
                        sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Right;
                    sheet.Range[lisrc, liColumns].VerticalAlignment = VerticalAlignType.Center;
                }

            }

            //sheet.Range[6, 1, 500, 1].ColumnWidth = 35;
            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 6;
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
                            sheet.Range[lisrc, liColumns].NumberFormat = "#,##0_);[Black]\\($#,##0\\)";
                        }
                        if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && sheet.Range[lisrc, liColumns].Text.Contains("%"))
                        {
                            sheet.Range[lisrc, liColumns].Text = sheet.Range[lisrc, liColumns].Text.Replace("%", "");
                            if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                                sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                            sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text) / 100);
                            sheet.Range[lisrc, liColumns].NumberFormat = "#,##0_);[Black]\\($#,##0\\)";// "$#,##0.0%_);\\($#,##0.0%\\)";
                        }
                        if (!String.IsNullOrEmpty(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text))
                        {
                            if (sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Contains("("))
                                sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Replace("(", "").Replace(")", ""))).ToString();
                            sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberValue = Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text);
                            //sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "#,##0.0_);\\(#,##0.0\\)";
                            

                           
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

            sheet.Range[6, 1, sheet.Rows.Length, sheet.Columns.Length].Style.Font.Size = 12;
            sheet.Range[6, 1, sheet.Rows.Length, sheet.Columns.Length].Style.Font.FontName = "Frutiger 55 Roman";
            sheet.Range[6, 1, sheet.Rows.Length, sheet.Columns.Length].RowHeight = 19.50;
            

            /* ---------------NEW LOGIC TEST-------------*/
            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 7;
                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    //Header Setting           
                    if (liCounter == 0)
                    {
                        sheet.Range[5, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                        sheet.Range[5, liColumns].Style.Font.Size = 14;
                        sheet.Range[5, liColumns].RowHeight = 55;
                        sheet.Range[5, liColumns].VerticalAlignment = VerticalAlignType.Center;
                        sheet.Range[5, liColumns].Style.Font.IsBold = false;
                        sheet.Range[5, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Left;
                        sheet.Range[5, liColumns].IsWrapText = true;
                        //sheet.Range[6, liColumns].Style.Color = System.Drawing.Color.FromArgb(216, 216, 216);
                    }

                    if (liColumns > 1)
                    {
                        sheet.Range[5, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Right;
                    }

                }
            }

          

            for (int k = 1; k <= sheet.Columns.Length; k++)
            {
                if (sheet.Range[5, k].Value == "Partnership")
                {
                    sheet.Range[5, k, sheet.Rows.Length, k].ColumnWidth = 20.86;
                    sheet.Range[5, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Left;
                }

                if (sheet.Range[5, k].Value.Contains("Partnership Market Value as of"))
                {
                    sheet.Range[5, k, sheet.Rows.Length, k].ColumnWidth = 20.86;
                    sheet.Range[5, k, 5, k].HorizontalAlignment = HorizontalAlignType.Right;
                }

                if (sheet.Range[5, k].Value == "Transactions")
                {
                    sheet.Range[5, k, sheet.Rows.Length, k].ColumnWidth = 25;
                    sheet.Range[5, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Right;
                }

                if (sheet.Range[5, k].Value.Contains("Partnership Market Value as of"))
                {
                    sheet.Range[5, k, sheet.Rows.Length, k].ColumnWidth = 20.86;
                    sheet.Range[5, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Right;
                }

                if (sheet.Range[5, k].Value.Contains("Commitment Value as of"))
                {
                    sheet.Range[5, k, sheet.Rows.Length, k].ColumnWidth = 20.86;
                    sheet.Range[5, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Right;
                    sheet.Range[5, k, sheet.Rows.Length, k].IsWrapText = true;
                }

                if (sheet.Range[5, k].Value == "MV as Percentage of Commitment")
                {
                    sheet.Range[5, k, sheet.Rows.Length, k].ColumnWidth = 20.86;
                    sheet.Range[5, k, sheet.Rows.Length, k].HorizontalAlignment = HorizontalAlignType.Right;
                }
            }

            sheet.Range["B5"].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["C5"].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["D5"].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["E5"].HorizontalAlignment = HorizontalAlignType.Center;
            sheet.Range["F5"].HorizontalAlignment = HorizontalAlignType.Center;

            sheet.Range[6, 1, 500, 1].ColumnWidth = 41.57;
            //sheet.Range[6, 2, 500, 2].ColumnWidth = 52;
            //sheet.Range[6, 3, 500, 3].ColumnWidth = 20;
            //sheet.Range[6, 4, 500, 4].ColumnWidth = 20;
            //sheet.Range[6, 5, 500, 5].ColumnWidth = 35;
            //sheet.Range[6, 5, 500, 5].IsWrapText = true;

            //sheet.Range[6, 6, 500, 6].ColumnWidth = 35;
            //sheet.Range[6, 6, 500, 6].IsWrapText = true;
            for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
            {
                int lisrc = liCounter + 6;
                string val1 = loInsertdataset.Tables[0].Rows[liCounter][1].ToString();
                if (val1 != "")
                {
                    sheet.Range[lisrc, 2].Text = "";
                    sheet.Range[lisrc, 2].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 2].NumberFormat = "#,##0_);[Black](#,##0)";
                }
                val1 = loInsertdataset.Tables[0].Rows[liCounter][2].ToString();
                if (val1 != "")
                {
                    sheet.Range[lisrc, 3].Text = "";
                    sheet.Range[lisrc, 3].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 3].NumberFormat = "#,##0_);[Black](#,##0)";
                }
                val1 = loInsertdataset.Tables[0].Rows[liCounter][3].ToString();
                if (val1 != "")
                {
                    sheet.Range[lisrc, 4].Text = "";
                    //  int a = (Convert.ToInt32(val1)) / 100;
                    sheet.Range[lisrc, 4].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 4].NumberFormat = "#,##0_);[Black](#,##0)";
                }
                val1 = loInsertdataset.Tables[0].Rows[liCounter][4].ToString();
                if (val1 != "")
                {
                    sheet.Range[lisrc, 5].Text = "";
                    sheet.Range[lisrc, 5].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                    sheet.Range[lisrc, 5].NumberFormat = "#,##0_);[Black](#,##0)";
                }
            }

            workbook.SaveToFile(strDirectory, ExcelVersion.Version2016);


            /**/
            #region not used
            ////Save workbook to disk
            //// workbook.Save();
            //workbook.SaveAsXml(strDirectory2);
            //workbook = null;
            //XmlDocument xmlDoc = new XmlDocument();
            //xmlDoc.Load(strDirectory2);
            //XmlElement businessEntities = xmlDoc.DocumentElement;
            //XmlNode loNode = businessEntities.LastChild;
            //XmlNode loNode1 = businessEntities.FirstChild;
            ////   businessEntities.RemoveChild(loNode);   comment becaue of for spire error 


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

        Workbook workbook1 = new Workbook();
        //  workbook1.LoadFromFile(strDirectory.Replace("xls", "xlsx"));
        workbook1.LoadFromFile(strDirectory);
        Worksheet sheet1 = workbook1.Worksheets[0];
        sheet1.Range[5, 1, 5, 6].Style.Color = System.Drawing.Color.FromArgb(183, 221, 232);
        sheet1.Range[6, 1, 1000, 1].IsWrapText = false;
        workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);


        loFile = new FileInfo(strDirectory2);
        loFile.Delete();
        loFile = null;
        lsFileNamforFinalXls = "/ExcelTemplate/TempFolder/" + lsFileNamforFinalXls;
        #endregion



        Response.Write("<script>");
      //  lsFileNamforFinalXls = "./ExcelTemplate/" + lsFileNamforFinalXls;
        Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        Response.Write("</script>");
        string baseUrl = Request.Url.GetLeftPart(UriPartial.Authority);
        Response.Redirect(baseUrl + lsFileNamforFinalXls);
     
        //Response.Redirect("report.xls");
    }
    protected void lstType_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        BindPartnership();
    }

    private string NotItemsFromListBox(ListBox lstBox)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 0; i < lstBox.Items.Count; i++)
            {
                if (lstBox.Items[i].Selected == false)
                {
                    lstselecteditems = lstselecteditems + "," + lstBox.Items[i].Value;
                    //insert command
                }
            }
            if (lstselecteditems != "")
            {
                lstselecteditems = lstselecteditems.Substring(1);
            }
        }


        return lstselecteditems;
    }
}
