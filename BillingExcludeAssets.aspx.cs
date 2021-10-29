using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
//using CrmSdk;
using System.Globalization;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.ServiceModel.Description;
using System.Configuration;
using Microsoft.Xrm.Sdk.Query;
using System.Security.Principal;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class BillingExcludeAssets : System.Web.UI.Page
{

    string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
    string strDescription;
    string[] SourceFileArray;
    GeneralMethods GM = new GeneralMethods();
    //  public DataTable dtData;
    // Billing
    // Flag=1 Assetlevel _+veAmount
    // Flag=2 Assetlevel _-veAmount
    // Flag=3 Assetlevel _-Exclude 

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            // DataTable dtData = new DataTable();
            DateTime dtAsOfDate = DateTime.Now;
            DateTime lastDay = new DateTime(dtAsOfDate.Year, dtAsOfDate.Month, 1); //1st Day of Current Month
            lastDay = lastDay.AddDays(-1);  //last date of previous month

            txtAUMDate.Text = lastDay.ToString("MM/dd/yyyy");
            DateTime date = DateTime.Now;
            DateTime quarterEnd = NearestQuarterEnd(date);
            txtAUMDate.Text = quarterEnd.ToShortDateString();
            this.Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "selectMonths('" + txtAUMDate.Text + "')", true);
            BindAdvisors();
            BindHouseHold();
            BindBillingName();
            //   BindGrideview();
            //DataSet dsData = getDataSet("EXEC SP_R_BILLING @BillingForUUID = '7E1DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'");
            //dtData = dsData.Tables[0];

            //DataTable dtTempData = dtData.Clone();
            //ViewState["dtTempData"] = dtData;

            //dtData.Columns.Add("BillingExtra");
            //dtData.Columns.Add("AUMExtra");
            //dtData.Columns.Add("BillingFeePctExtra");
            //foreach (DataRow row in dtData.Rows)
            //{
            //    row["BillingExtra"] = row["FinalBillingMarketValue"].ToString();
            //    row["AUMExtra"] = row["FinalAUMMarketValue"].ToString();
            //    row["BillingFeePctExtra"] = row["BillingFeePct"].ToString();
            //}


            //BindGrideview(dtData);
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

    # region Bind all dropdown
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
        // "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
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

    protected void ddlAdvisor_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindHouseHold();
        BindBillingName();

        lblAUMDate.Visible = false;
        lblBillingFor.Visible = false;
        gvBilling.Visible = false;
        btnSave.Visible = false;
        btnSave0.Visible = false;
        btnSaveGenrate.Visible = false;
        btnSaveGenrate0.Visible = false;
        btnSaveGenratewithNonGA.Visible = false;
        btnSaveGenratewithNonGA0.Visible = false;
        //lblSavePopUp.Visible=tr


    }

    protected void ddlHH_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessageShow.Text = "";
        lblMessage.Text = "";
        BindBillingName();

        lblAUMDate.Visible = false;
        lblBillingFor.Visible = false;
        gvBilling.Visible = false;
        btnSave.Visible = false;
        btnSave0.Visible = false;
        btnSaveGenrate.Visible = false;
        btnSaveGenrate0.Visible = false;
        btnSaveGenratewithNonGA.Visible = false;
        btnSaveGenratewithNonGA0.Visible = false;
    }
    #endregion

    # region Submit button Not in use
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        int i;
        //DataSet dsData = getDataSet("EXEC SP_R_BILLING @BillingForUUID = '7E1DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'");
        ////dtData = dsData.Tables[0];
        //BindGrideview(dsData.Tables[0]);

        BindGrideview();

    }
    #endregion

    public void BindGrideview()
    {
        lblMessage.Text = "";
        //lblSavePopUp.Text = "";
        string Query = null;
        string billingVal = ddlBillFor.SelectedValue.ToString();
        // string billingVal = ddlBillFor.SelectedItem.Text.ToString();
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
            txtAUMDate.Text = dDate;
        }
        if (billingVal != "ALL")
        {

            //    if (dDate != "")
            //    {

            int len = billingVal.IndexOf("|");
            billingVal = billingVal.Substring(0, len);

            #region DropDowns 
            DataSet dsBillingDrodowns = new DataSet();
            Query = "EXEC sp_s_BillingDropDownList @DropDownType = 1";
            DataSet dsAdminCategory = getDataSet(Query);
            DataTable dtCopy = new DataTable();
            dtCopy = dsAdminCategory.Tables[0].Copy();
            dtCopy.TableName = "NewCopy";
            dsBillingDrodowns.Tables.Add(dtCopy);


            Query = "EXEC sp_s_BillingDropDownList @DropDownType = 2";
            DataSet dsNonGaServiceCategory = getDataSet(Query);
            DataTable dtCopy1 = dsNonGaServiceCategory.Tables[0].Copy();
            dsBillingDrodowns.Tables.Add(dtCopy1);


            ViewState["BillingDropDowns"] = dsBillingDrodowns;
            #endregion


            if (dDate != "")
                Query = "EXEC SP_R_BILLING @ReportFlg = 1, @BillingForUUID = '" + billingVal + "' ,@AsOfDate = '" + dDate + "'";
            else
                Query = "EXEC SP_R_BILLING @ReportFlg = 1, @BillingForUUID = '7E1DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'";


            DataSet dsData = getDataSet(Query);
            if (dsData.Tables.Count > 0)
            {
                if (dsData.Tables[0].Rows.Count > 0)
                {
                    DataTable dtData = dsData.Tables[0];

                    if (dtData.Rows.Count > 0)
                    {
                        btnSave.Visible = true;
                        btnSaveGenrate.Visible = true;
                        gvBilling.Visible = true;
                        btnSave.Visible = true;
                        btnSaveGenrate.Visible = true;
                        btnSave0.Visible = true;
                        btnSaveGenrate0.Visible = true;
                        btnSaveGenratewithNonGA.Visible = true;
                        btnSaveGenratewithNonGA0.Visible = true;
                        lblAUMDate.Visible = true;
                        lblBillingFor.Visible = true;
                        lblBillingFor.Text = ddlBillFor.SelectedItem.Text;
                        lblAUMDate.Text = txtAUMDate.Text;
                    }
                    else
                    {
                        btnSave.Visible = false;

                        btnSaveGenrate.Visible = false;
                        btnSaveGenrate0.Visible = false;
                        btnSaveGenratewithNonGA.Visible = false;
                        btnSaveGenratewithNonGA0.Visible = false;
                        btnSave0.Visible = false;
                        gvBilling.Visible = false;
                        lblBillingFor.Visible = false;
                        lblAUMDate.Visible = false;
                    }

                    DataTable dtTempData = dtData.Clone();
                    //ViewState["dtTempData"] = dtData;

                    dtData.Columns.Add("BillingExtra");
                    dtData.Columns.Add("AUMExtra");
                    dtData.Columns.Add("BillingFeePctExtra");
                    foreach (DataRow row in dtData.Rows)
                    {
                        row["BillingExtra"] = row["FinalBillingMarketValue"].ToString();
                        row["AUMExtra"] = row["FinalAUMMarketValue"].ToString();
                        row["BillingFeePctExtra"] = row["BillingFeePct"].ToString();
                    }


                    // gvBilling.Columns[9].Visible = true;
                    gvBilling.Columns[10].Visible = true;
                    gvBilling.Columns[11].Visible = true;
                    gvBilling.Columns[12].Visible = true;
                    gvBilling.Columns[13].Visible = true;
                    gvBilling.Columns[14].Visible = true;
                    gvBilling.Columns[15].Visible = true;
                    gvBilling.Columns[16].Visible = true;
                    gvBilling.Columns[17].Visible = true;

                    gvBilling.Columns[18].Visible = true;
                    gvBilling.Columns[19].Visible = true;
                    gvBilling.Columns[20].Visible = true;
                    gvBilling.Columns[21].Visible = true;
                    gvBilling.Columns[22].Visible = true;
                    gvBilling.Columns[23].Visible = true;
                    gvBilling.Columns[24].Visible = true;
                    gvBilling.Columns[25].Visible = true;
                    gvBilling.Columns[26].Visible = true;
                    gvBilling.Columns[27].Visible = true;
                    gvBilling.Columns[28].Visible = true;
                    gvBilling.Columns[29].Visible = true;
                    gvBilling.Columns[30].Visible = true;
                    gvBilling.Columns[31].Visible = true;
                    gvBilling.Columns[32].Visible = true;
                    gvBilling.Columns[33].Visible = true;
                    gvBilling.Columns[34].Visible = true;
                    gvBilling.Columns[35].Visible = true;
                    gvBilling.Columns[36].Visible = true;

                    gvBilling.Columns[37].Visible = true;
                    gvBilling.Columns[38].Visible = true;
                    gvBilling.Columns[39].Visible = true;
                    gvBilling.Columns[40].Visible = true;
                    gvBilling.Columns[41].Visible = true;
                    gvBilling.Columns[42].Visible = true;
                    gvBilling.Columns[43].Visible = true;
                    gvBilling.Columns[44].Visible = true;
                    gvBilling.Columns[45].Visible = true;
                    gvBilling.Columns[46].Visible = true;
                    gvBilling.Columns[47].Visible = true;
                    gvBilling.Columns[48].Visible = true;
                    gvBilling.Columns[49].Visible = true;
                    gvBilling.Columns[50].Visible = true;
                    gvBilling.Columns[51].Visible = true;


                    gvBilling.Columns[52].Visible = true;
                    gvBilling.Columns[53].Visible = true;
                    gvBilling.Columns[54].Visible = true;
                    gvBilling.Columns[55].Visible = true;
                    gvBilling.Columns[56].Visible = true;
                    gvBilling.Columns[57].Visible = true;
                    gvBilling.Columns[58].Visible = true;
                    gvBilling.Columns[59].Visible = true;
                    gvBilling.Columns[60].Visible = true;
                    gvBilling.Columns[61].Visible = true;
                    gvBilling.Columns[62].Visible = true;
                    gvBilling.Columns[63].Visible = true;
                    gvBilling.DataSource = dtData;

                    gvBilling.DataBind();

                    //txtBilling.
                    // gvBilling.Columns[9].Visible = false;

                    gvBilling.Columns[10].Visible = false;
                    gvBilling.Columns[11].Visible = false;
                    gvBilling.Columns[12].Visible = false;
                    gvBilling.Columns[13].Visible = false;
                    gvBilling.Columns[14].Visible = false;
                    gvBilling.Columns[15].Visible = false;
                    gvBilling.Columns[16].Visible = false;
                    gvBilling.Columns[17].Visible = false;

                    gvBilling.Columns[18].Visible = false;
                    gvBilling.Columns[19].Visible = false;
                    gvBilling.Columns[20].Visible = false;
                    gvBilling.Columns[21].Visible = false;
                    gvBilling.Columns[22].Visible = false;
                    gvBilling.Columns[23].Visible = false;
                    gvBilling.Columns[24].Visible = false;
                    gvBilling.Columns[25].Visible = false;
                    gvBilling.Columns[26].Visible = false;
                    gvBilling.Columns[27].Visible = false;
                    gvBilling.Columns[28].Visible = false;
                    gvBilling.Columns[29].Visible = false;
                    gvBilling.Columns[30].Visible = false;
                    gvBilling.Columns[31].Visible = false;
                    gvBilling.Columns[32].Visible = false;
                    gvBilling.Columns[33].Visible = false;
                    gvBilling.Columns[34].Visible = false;
                    gvBilling.Columns[35].Visible = false;
                    gvBilling.Columns[36].Visible = false;
                    gvBilling.Columns[37].Visible = false;
                    gvBilling.Columns[38].Visible = false;
                    gvBilling.Columns[39].Visible = false;
                    gvBilling.Columns[40].Visible = false;
                    gvBilling.Columns[41].Visible = false;
                    gvBilling.Columns[42].Visible = false;
                    gvBilling.Columns[43].Visible = false;
                    gvBilling.Columns[44].Visible = false;
                    gvBilling.Columns[45].Visible = false;
                    gvBilling.Columns[46].Visible = false;
                    gvBilling.Columns[47].Visible = false;
                    gvBilling.Columns[48].Visible = false;
                    gvBilling.Columns[49].Visible = false;
                    gvBilling.Columns[50].Visible = false;
                    gvBilling.Columns[51].Visible = false;

                    gvBilling.Columns[55].Visible = false;
                    gvBilling.Columns[56].Visible = false;
                    gvBilling.Columns[57].Visible = false;
                    gvBilling.Columns[58].Visible = false;
                    gvBilling.Columns[59].Visible = false;
                    gvBilling.Columns[60].Visible = false;
                    gvBilling.Columns[61].Visible = false;
                    gvBilling.Columns[62].Visible = false;
                    gvBilling.Columns[63].Visible = false;


                    //string hed = gvBilling.Columns[10].HeaderText.ToString();

                    foreach (GridViewRow grow in gvBilling.Rows)
                    {
                        string AssetLevelFlgNew = grow.Cells[16].Text;
                        string TotalAssetLevelFlgNew = grow.Cells[17].Text;
                        string ActualVal = grow.Cells[4].Text;

                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox TotalAUMText = (TextBox)grow.FindControl("txtAUM");
                            TextBox TotalBillingText = (TextBox)grow.FindControl("txtBilling");
                            TotalAUMText.Enabled = true;
                            TotalBillingText.Enabled = true;
                            grow.BackColor = System.Drawing.Color.LightGray;

                        }
                        if (TotalAssetLevelFlgNew == "True")
                        {
                            TextBox TotalAUMText = (TextBox)grow.FindControl("txtAUM");
                            TextBox TotalBillingText = (TextBox)grow.FindControl("txtBilling");
                            TotalAUMText.Enabled = false;
                            TotalBillingText.Enabled = false;

                            CheckBox cbBilling = (CheckBox)grow.FindControl("cbExcludeBilling");
                            cbBilling.Visible = false;
                            CheckBox cbAum = (CheckBox)grow.FindControl("cbExcludeAum");
                            cbAum.Visible = false;

                            int Rowindex = grow.RowIndex;

                            decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - Convert.ToDecimal(TotalBillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            grow.Cells[6].Text = currencyFormat(val.ToString());

                            val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - Convert.ToDecimal(TotalAUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            grow.Cells[8].Text = currencyFormat(val.ToString());


                        }
                        string BillingExcludeFlg = grow.Cells[23].Text;
                        string AUMExcludeFlg = grow.Cells[24].Text;


                        if (BillingExcludeFlg == "True")
                        {
                            TextBox TotalBillingText = (TextBox)grow.FindControl("txtBilling");
                            TotalBillingText.Enabled = false;
                            TextBox TotalBillingPerText = (TextBox)grow.FindControl("txtBillPer");
                            TotalBillingPerText.Enabled = false;

                        }
                        if (AUMExcludeFlg == "True")
                        {
                            TextBox TotalBillingText = (TextBox)grow.FindControl("txtAUM");
                            TotalBillingText.Enabled = false;
                        }
                        string ACLevelBillingExceptionFlg = grow.Cells[14].Text;
                        string ACLevelAUMExceptionFlg = grow.Cells[15].Text;

                        if (ACLevelBillingExceptionFlg == "True" && AssetLevelFlgNew != "True")
                        {
                            //if (grow.Cells[19].Text != "3")
                            //{
                            TextBox TotalBillingText = (TextBox)grow.FindControl("txtBilling");
                            TotalBillingText.Enabled = false;

                            TextBox TotalBillingPerText = (TextBox)grow.FindControl("txtBillPer");
                            TotalBillingPerText.Enabled = false;

                            CheckBox cbBilling = (CheckBox)grow.Cells[5].FindControl("cbExcludeBilling");
                            cbBilling.Enabled = false;
                            grow.Cells[42].Text = "3";
                            //}
                        }
                        if (ACLevelBillingExceptionFlg == "True" && AssetLevelFlgNew == "True")
                        {
                            grow.Cells[42].Text = "3";
                        }

                        if (ACLevelAUMExceptionFlg == "True" && AssetLevelFlgNew != "True")
                        {
                            TextBox TotalAUMText = (TextBox)grow.FindControl("txtAUM");
                            TotalAUMText.Enabled = false;

                            CheckBox cbAum = (CheckBox)grow.FindControl("cbExcludeAum");
                            cbAum.Enabled = false;
                            grow.Cells[43].Text = "3";
                        }
                        if (ACLevelAUMExceptionFlg == "True" && AssetLevelFlgNew == "True")
                        {
                            grow.Cells[43].Text = "3";
                        }

                        TextBox TotalAUMTText = (TextBox)grow.FindControl("txtAUM");
                        string value = TotalAUMTText.Text;

                        decimal ul;
                        if (decimal.TryParse(value, out ul))
                        {

                            TotalAUMTText.TextChanged -= txtAUM_TextChanged;
                            TotalAUMTText.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                            TotalAUMTText.TextChanged += txtAUM_TextChanged;
                        }

                        TextBox TotalBillingTxt = (TextBox)grow.FindControl("txtBilling");
                        value = TotalBillingTxt.Text;
                        if (decimal.TryParse(value, out ul))
                        {

                            TotalBillingTxt.TextChanged -= txtBilling_TextChanged;
                            TotalBillingTxt.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                            TotalBillingTxt.TextChanged += txtBilling_TextChanged;
                        }

                        string PerValue = grow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                        string BillingType = grow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                        string AUMType = grow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");

                        //BillingType=BillingType.Replace("&nbsp", "").Replace(";","");
                        //AUMType = AUMType.Replace("&nbsp", "").Replace(";", "");
                        //PerValue = PerValue.Replace("&nbsp", "").Replace(";", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        // if (BillingType != "1" && BillingType != "")

                        if (BillingType != "" && BillingType != "5") // for billing Yellow color
                        {
                            int index = grow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        else if (BillingType == "5") // for Billing red color
                        {
                            int index = grow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }


                        //  if (AUMType != "1" && AUMType != "")
                        if (AUMType != "" && AUMType != "5")    // for AUM yellow color
                        {
                            int index = grow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        else if (AUMType == "5")                // for AUM red color
                        {
                            int index = grow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = grow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        string diffflag = grow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            grow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }
                    }

                    //ViewState["dtData"] = dtData;
                    //    }
                    //    else
                    //    {
                    //        lblMessage.Text = "Please Select Date";
                    //    }
                    //}
                    //else
                    //{

                    //    lblMessage.Text = "Please Select Billing";
                    //}
                }

                else
                {
                    lblMessageShow.Visible = true;
                    lblMessageShow.Text = "No Records Found";
                    gvBilling.Visible = false;


                    //lblAUMDate.Visible = false;
                    //lblBillingFor.Visible = false;
                }

            }
            else
            {
                lblMessageShow.Visible = true;
                lblMessageShow.Text = "No Records Found";
                gvBilling.Visible = false;
                //lblAUMDate.Visible = false;
                //lblBillingFor.Visible = false;
            }

        }
        else
        {
            lblMessage.Text = "Please Select Billing";
        }

    }

    public DataSet getDataSet(string vSqlQuery)
    {
        DataSet ds = new DataSet();
        try
        {
            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter dagersham = new SqlDataAdapter();
            DataSet ds_gresham = new DataSet();

            dagersham = new SqlDataAdapter(vSqlQuery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds);//Fill Dataset 
        }
        catch (Exception ex)
        {

            lblMessage.ForeColor = System.Drawing.Color.Red;
            //lblMessage.Text = "Error in getting dataset value" + ex.Message;
        }
        return ds;

    }

    #region Checkbox and textboks changes

    protected void cbExcludeBilling_CheckedChanged(object sender, EventArgs e)
    {
        try
        {
            lblMessage.Text = "";
            CheckBox ck1 = (CheckBox)sender;
            GridViewRow grow = (GridViewRow)ck1.NamingContainer;
            string IdNumb1 = grow.Cells[22].Text;
            string IdNumbTotalAsset = "", TotalVal = null;
            int rowindex = -1;
            string BillingExtra = grow.Cells[25].Text;
            int roindex = grow.RowIndex;
            //int a= e.Row.RowIndex;
            //DataTable dtData = (DataTable)ViewState["dtData"];

            string Actual = grow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            //foreach (GridViewRow row in gvBilling.Rows) 
            //{
            //CheckBox chk = (CheckBox)row.Cells[6].FindControl("cbExcludeBilling");
            if (ck1.Checked)
            {
                TextBox tx = (TextBox)grow.FindControl("txtBilling");
                //   string BillingNew = tx.Text;
                string IdNumb = grow.Cells[22].Text;
                //string AccountID = grow.Cells[10].Text;
                //string SecurityID = grow.Cells[11].Text;
                string AssetClassID = grow.Cells[12].Text;
                string BillingID = grow.Cells[13].Text;
                string ACLevelBillingExceptionFlg = grow.Cells[14].Text;
                string ACLevelAUMExceptionFlg = grow.Cells[15].Text;

                string AssetLevelFlg = grow.Cells[16].Text;
                string TotalLevelFlg = grow.Cells[17].Text;
                string BillingExceptionId = grow.Cells[18].Text;
                string BillingFeeExceptionId = grow.Cells[19].Text;

                string BillingExceptionType = grow.Cells[20].Text;
                string AUMExceptionType = grow.Cells[21].Text;



                //string PositiveAum = grow.Cells[46].Text;
                //string PositiveBilling = grow.Cells[42].Text;
                //grow.Cells[46].Text = "";
                grow.Cells[44].Text = "";
                //     string AUMExceptionType = row.Cells[21].Text; 
                if (AssetLevelFlg == "True")    //  for Total Asset class ID   set checkbox value = true
                {

                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string ActualVal = gvrow.Cells[4].Text;
                        if (AssetClassID == AssetClassIDNew)
                        {
                            //  IdNumbTotalAsset = gvrow.Cells[21].Text;
                            rowindex = gvrow.RowIndex;
                            TextBox Billingtxt = (TextBox)gvrow.FindControl("txtBilling");
                            Billingtxt.Text = "0";
                            Billingtxt.Enabled = false;

                            TextBox txtNGAA = (TextBox)gvrow.FindControl("txtNGAA");
                            txtNGAA.Text = currencyFormat(ActualVal);

                            TextBox txtPer = (TextBox)gvrow.FindControl("txtBillPer");
                            txtPer.Enabled = false;


                            CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                            // cbBilling.Checked = true;
                            cbBilling.Enabled = false;
                            gvrow.Cells[42].Text = "3";

                            gvrow.Cells[44].Text = "";

                            //   gvrow.Cells[46].Text = "";   // positive AUM
                            gvrow.Cells[44].Text = "";  // positive billing
                            gvrow.Cells[32].Text = "";  //sub Billing 
                            //   gvrow.Cells[33].Text = "";   // sun AUM

                            if (AssetLevelFlgNew == "True")
                            {
                                cbBilling.Enabled = true;

                                gvrow.Cells[44].Text = "";
                            }
                            //gvrow.Cells[43].Text = "Asset";
                            //gvrow.Cells[45].Text = "Exclude";
                            //gvrow.Cells[13].Text ="True";
                        }

                    }






                    decimal TotalBilling = 0, TotalAUM = 0, TotalNGAA = 0;
                    int TotalRowIndex = 0;
                    string vActualValue = null;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;
                        string ActualVal = gvrow.Cells[4].Text;

                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalAUM = TotalAUM + AUMVal;

                            TextBox txtNGAA = (TextBox)gvrow.FindControl("txtNGAA");
                            decimal NGAAVal = Convert.ToDecimal(txtNGAA.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalNGAA = TotalNGAA + NGAAVal;

                            if (BillingVal > 0)
                            {
                                TotalBilling = TotalBilling + BillingVal;

                            }
                            else
                            {
                                TotalBilling = TotalBilling + BillingVal;
                            }
                        }

                        if (totalBilling != "False")
                        {
                            TotalRowIndex = gvrow.RowIndex;

                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[6].Text = val.ToString();

                        }



                        //string totalBilling0 = grow.Cells[15].Text;
                        //string totalBilling1 = grow.Cells[16].Text;
                        //string totalBilling2 = grow.Cells[17].Text;
                        //TextBox BilText = (TextBox)gvrow.FindControl("txtBilling");
                        //TotalBilling = TotalBilling + Convert.ToDecimal(BilText.Text);
                        //int id = gvrow.RowIndex;



                        #region  Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        vActualValue = gvrow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""); // get actual Value

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            // gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{

                        if (PerValue != "1" && PerValue != "")
                        {
                            int index1 = gvrow.RowIndex;
                            CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                            if (!cbBilling.Checked)
                            {
                                TextBox text1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBillPer");
                                text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }


                        int index = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            Billingtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (val > 0 && Convert.ToDecimal(vActualValue) < val)
                        {
                            if (Convert.ToDecimal(vActualValue) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }



                        //}
                        //  int index = gvrow.RowIndex;
                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            aumtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        //if (gvrow.Cells[17].Text == "True")
                        //{
                        //    string colval = gvrow.Cells[6].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "");
                        //    if (Convert.ToDecimal(colval) != 0)
                        //        Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    else
                        //        Billingtext1.BackColor = System.Drawing.Color.White;

                        //    if (Convert.ToDecimal(gvrow.Cells[8].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "")) != 0)
                        //        aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    else
                        //        aumtext1.BackColor = System.Drawing.Color.White;

                        //}

                        if (gvrow.Cells[17].Text == "True")
                        {
                            //int index1 = gvrow.RowIndex;

                            //TextBox Billingtex = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                            //decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                            //if (Convert.ToDecimal(vActualValue) != value)
                            //    Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            //else
                            //    Billingtex.BackColor = System.Drawing.Color.White;


                            //TextBox aumtex = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                            //decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                            //if (Convert.ToDecimal(vActualValue) != aumval1)
                            //    aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            //else
                            //    aumtex.BackColor = System.Drawing.Color.White;

                        }



                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");


                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        //if (PerValue != "1" && PerValue != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}


                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }


                        #endregion

                    }



                    TextBox tex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    tex.Text = TotalBilling.ToString();
                    addFormat(tex, "Billing");

                    TextBox texNGAA = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtNGAA");
                    texNGAA.Text = currencyFormat(TotalNGAA.ToString());

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowIndex].Cells[6].Text = currencyFormat(Totalval.ToString());


                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowIndex].Cells[8].Text = currencyFormat(Totalval.ToString());



                    #region totalTextcolor
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(vActualValue) < value)
                    {
                        if (Convert.ToDecimal(vActualValue) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }




                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(vActualValue) < aumval1)
                    {
                        if (Convert.ToDecimal(vActualValue) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    #endregion


                    //for (int i = 0; i < dtData.Rows.Count; i++)
                    //{
                    //    string ID = dtData.Rows[i]["sas_assetclassid"].ToString();
                    //    if (ID == AssetClassID)
                    //    {
                    //        dtData.Rows[i]["BillingExcludeFlg"] = "True";
                    //    }
                    //}




                }
                else    // For single record 
                {



                    //for (int i = 0; i < dtData.Rows.Count; i++)   // for single ID  set checkbox value = true
                    //{
                    //    if (dtData.Rows[i]["Idnmb"].ToString() == IdNumb1)
                    //    {
                    //        dtData.Rows[i]["BillingExcludeFlg"] = "True";
                    //        break;
                    //    }
                    //}

                    //string AssetClassID = grow.Cells[11].Text;
                    //string AssetLevelFlg = grow.Cells[15].Text;

                    roindex = grow.RowIndex;
                    TextBox Billingtxt = (TextBox)grow.FindControl("txtBilling");
                    Billingtxt.Text = "0";
                    Billingtxt.Enabled = false;

                    TextBox txtNGAA = (TextBox)grow.FindControl("txtNGAA");
                    txtNGAA.Text = currencyFormat(Actual);

                    TextBox txtPer = (TextBox)grow.FindControl("txtBillPer");
                    txtPer.Enabled = false;

                    grow.Cells[44].Text = "";
                    grow.Cells[32].Text = "";

                    int TotalRowIndex = 0;
                    decimal newBilling = 0, TotalValue = 0, TotalVNGAA = 0; ;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {

                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;
                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew != "True" && totalBilling != "True")
                        {

                            TextBox Billingtx = (TextBox)gvrow.FindControl("txtBilling");
                            decimal val = Convert.ToDecimal(Billingtx.Text.Replace("$", "").Replace(" ", "").Replace("(", "-").Replace(")", ""));
                            TotalValue = TotalValue + val;

                            TextBox NGAAtx = (TextBox)gvrow.FindControl("txtNGAA");
                            decimal valNGAA = Convert.ToDecimal(NGAAtx.Text.Replace("$", "").Replace(" ", "").Replace("(", "-").Replace(")", ""));
                            TotalVNGAA = TotalVNGAA + valNGAA;

                        }

                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True" && totalBilling != "True")
                        {
                            TotalRowIndex = gvrow.RowIndex;
                        }


                    }
                    TextBox tet = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    tet.Text = TotalValue.ToString();
                    addFormat(tet, "Billing");

                    TextBox tetNGAA = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtNGAA");
                    tetNGAA.Text = currencyFormat(TotalVNGAA.ToString());





                    decimal TotalBilling = 0, TotalAUM = 0, TotalNGAA = 0;
                    int TotalRowInde = 0;
                    string vActualValue1 = null;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;


                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalAUM = TotalAUM + AUMVal;

                            TextBox NGAAText = (TextBox)gvrow.FindControl("txtNGAA");
                            decimal NGAAVal = Convert.ToDecimal(NGAAText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalNGAA = TotalNGAA + NGAAVal;

                            if (BillingVal > 0)
                            {
                                TotalBilling = TotalBilling + BillingVal;
                            }
                            else
                            {
                                TotalBilling = TotalBilling + BillingVal;
                            }
                        }

                        if (totalBilling != "False")
                        {
                            TotalRowInde = gvrow.RowIndex;
                            //string ActualVal = gvrow.Cells[4].Text;
                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[6].Text = val.ToString();
                        }
                        //string totalBilling0 = grow.Cells[15].Text;
                        //string totalBilling1 = grow.Cells[16].Text;
                        //string totalBilling2 = grow.Cells[17].Text;
                        //TextBox BilText = (TextBox)gvrow.FindControl("txtBilling");
                        //TotalBilling = TotalBilling + Convert.ToDecimal(BilText.Text);
                        //int id = gvrow.RowIndex;




                        #region  Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    if (roindex != index)
                        //    {
                        //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    }
                        //}

                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeAum");
                        //    if (!cbBilling.Checked)
                        //    {
                        //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    }

                        //}

                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = gvrow.RowIndex;
                            CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                            if (!cbBilling.Checked)
                            {
                                TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                                text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }

                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }


                        vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        int index1 = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            Billingtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }


                        //}
                        //  int index = gvrow.RowIndex;
                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            aumtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");

                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        //if (gvrow.Cells[17].Text == "True")
                        //{
                        //    if (Convert.ToDecimal(gvrow.Cells[6].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "")) != 0)
                        //        Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    else
                        //        Billingtext1.BackColor = System.Drawing.Color.White;


                        //    if (Convert.ToDecimal(gvrow.Cells[8].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "")) != 0)
                        //        aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    else
                        //        aumtext1.BackColor = System.Drawing.Color.White;
                        //}




                        #endregion


                    }

                    TextBox text = (TextBox)gvBilling.Rows[TotalRowInde].FindControl("txtBilling");
                    text.Text = TotalBilling.ToString();
                    addFormat(text, "Billing");


                    TextBox textNGAA = (TextBox)gvBilling.Rows[TotalRowInde].FindControl("txtNGAA");
                    textNGAA.Text = currencyFormat(TotalNGAA.ToString());


                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowInde].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowInde].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    string TotalActual = gvBilling.Rows[TotalRowInde].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowInde].Cells[6].Text = currencyFormat(Totalval.ToString());

                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowInde].Cells[8].Text = currencyFormat(Totalval.ToString());

                    #region totalTextcolor
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowInde].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }


                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowInde].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");

                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }


                    #endregion

                }

            }
            else
            {
                string AssetClassID = grow.Cells[12].Text;
                string AssetLevelFlg = grow.Cells[16].Text;
                string FinalNGAAMarketValue = grow.Cells[57].Text;
                if (AssetLevelFlg == "True")                 //  for Total Asset class ID  set checkbox value = false
                {
                    decimal AssetTotalNGAA = 0;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string BillingExceptionID = gvrow.Cells[18].Text;

                        if (AssetClassID == AssetClassIDNew)
                        {
                            //  IdNumbTotalAsset = gvrow.Cells[21].Text;
                            rowindex = gvrow.RowIndex;
                            TextBox Billingtxt = (TextBox)gvrow.FindControl("txtBilling");

                            string val = gvrow.Cells[4].Text;
                            Billingtxt.Text = val;

                            Billingtxt.Enabled = true;

                            TextBox txtPer = (TextBox)gvrow.FindControl("txtBillPer");
                            txtPer.Enabled = true;


                            //TextBox txtNGAA = (TextBox)gvrow.FindControl("txtNGAA");
                            //txtNGAA.Text = currencyFormat(FinalNGAAMarketValue);

                            CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");



                            if (gvrow.Cells[23].Text != "True")
                            {
                                cbBilling.Checked = false;
                                cbBilling.Enabled = true;

                            }
                            else
                            {
                                Billingtxt.Enabled = false;
                                Billingtxt.Text = gvrow.Cells[25].Text;
                                cbBilling.Checked = true;
                                cbBilling.Enabled = true;
                            }

                            if (AssetLevelFlgNew == "True")
                            {
                                cbBilling.Checked = false;
                                Billingtxt.Enabled = true;

                                Billingtxt.Text = gvrow.Cells[4].Text;

                            }
                            if (cbBilling.Checked)
                            {
                                TextBox txtNGAA = (TextBox)gvrow.FindControl("txtNGAA");
                                txtNGAA.Text = currencyFormat(val);

                                cbBilling.Checked = true;
                                AssetTotalNGAA = AssetTotalNGAA + Convert.ToDecimal(val.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            }
                            else
                            {
                                //if (AssetLevelFlgNew != "True")
                                //{
                                TextBox txtNGAA = (TextBox)gvrow.FindControl("txtNGAA");
                                txtNGAA.Text = currencyFormat("0");
                                //  AssetTotalNGAA = AssetTotalNGAA + val;
                                //}
                            }
                            // if (BillingExceptionID != "" && AssetLevelFlgNew != "True")
                            gvrow.Cells[42].Text = "";

                            //gvrow.Cells[13].Text = "False";
                            gvrow.Cells[25].Text = "0";

                            addFormat(Billingtxt, "Billing");
                        }
                    }

                    TextBox txtAssetNGAA1 = (TextBox)grow.FindControl("txtNGAA");
                    txtAssetNGAA1.Text = currencyFormat(AssetTotalNGAA.ToString());


                    decimal TotalBilling = 0, TotalAUM = 0, TotalNGAA = 0;
                    int TotalRowIndex = 0;
                    string vActualValue1 = null;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;


                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalAUM = TotalAUM + AUMVal;

                            TextBox NGAAText = (TextBox)gvrow.FindControl("txtNGAA");
                            decimal NGAAal = Convert.ToDecimal(NGAAText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalNGAA = TotalNGAA + NGAAal;

                            if (BillingVal > 0)
                            {
                                TotalBilling = TotalBilling + BillingVal;
                            }
                            else
                            {
                                TotalBilling = TotalBilling + BillingVal;
                            }
                        }
                        if (totalBilling != "False")
                        {
                            TotalRowIndex = gvrow.RowIndex;
                            // string ActualVal = gvrow.Cells[4].Text;
                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[6].Text = val.ToString();
                        }
                        //string totalBilling0 = grow.Cells[15].Text;
                        //string totalBilling1 = grow.Cells[16].Text;
                        //string totalBilling2 = grow.Cells[17].Text;
                        //TextBox BilText = (TextBox)gvrow.FindControl("txtBilling");
                        //TotalBilling = TotalBilling + Convert.ToDecimal(BilText.Text);
                        //int id = gvrow.RowIndex;



                        #region  Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = gvrow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }



                        vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        int index1 = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            Billingtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        //}
                        //  int index = gvrow.RowIndex;


                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            aumtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }



                        #endregion


                    }

                    TextBox txt = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    txt.Text = TotalBilling.ToString();
                    addFormat(txt, "Billing");

                    TextBox txtNGAA1 = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtNGAA");
                    txtNGAA1.Text = currencyFormat(TotalNGAA.ToString());

                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowIndex].Cells[6].Text = currencyFormat(Totalval.ToString());


                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowIndex].Cells[8].Text = currencyFormat(Totalval.ToString());


                    #region total color
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                    if (Convert.ToDecimal(vActualValue1) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;


                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }




                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                    if (Convert.ToDecimal(vActualValue1) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    #endregion

                    //for (int i = 0; i < dtData.Rows.Count; i++)
                    //{
                    //    string ID = dtData.Rows[i]["sas_assetclassid"].ToString();
                    //    if (ID == AssetClassID)
                    //    {
                    //        dtData.Rows[i]["BillingExcludeFlg"] = "False";
                    //    }
                    //}
                }
                else  // for Single record  checkbox false
                {
                    roindex = grow.RowIndex;
                    TextBox Billingtxt = (TextBox)grow.FindControl("txtBilling");
                    Billingtxt.Enabled = true;
                    Billingtxt.Text = grow.Cells[4].Text;
                    addFormat(Billingtxt, "Billing");

                    TextBox txtNGAA = (TextBox)grow.FindControl("txtNGAA");
                    //    txtNGAA.Text = currencyFormat(grow.Cells[57].Text);
                    txtNGAA.Text = currencyFormat("0");
                    // addFormat(txtNGAA, "Billing");

                    TextBox txtPer = (TextBox)grow.FindControl("txtBillPer");
                    txtPer.Enabled = true;

                    //addFormat(Billingtxt, "Billing");

                    decimal newBilling = 0, newGAA = 0;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string TotalValNew = gvrow.Cells[17].Text;

                        //if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True")
                        //{
                        //    // IdNumbTotalAsset = gvrow.Cells[21].Text;
                        //    rowindex = gvrow.RowIndex;
                        //    TextBox Billingtx = (TextBox)gvrow.FindControl("txtBilling");
                        //    TotalVal = Billingtx.Text;
                        //    break;
                        //}


                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "False" && TotalValNew == "False")
                        {
                            TextBox tx = (TextBox)gvrow.FindControl("txtBilling");
                            decimal val = Convert.ToDecimal(tx.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            newBilling = newBilling + val;

                            TextBox txNGAA = (TextBox)gvrow.FindControl("txtNGAA");
                            decimal valNGAA = Convert.ToDecimal(txNGAA.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            newGAA = newGAA + valNGAA;
                        }
                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True" && TotalValNew == "False")
                        {
                            rowindex = gvrow.RowIndex;
                        }
                    }

                    TextBox tet = (TextBox)gvBilling.Rows[rowindex].FindControl("txtBilling");
                    tet.Text = newBilling.ToString();
                    addFormat(tet, "Billing");

                    TextBox tetNGAA = (TextBox)gvBilling.Rows[rowindex].FindControl("txtNGAA");
                    tetNGAA.Text = currencyFormat(newGAA.ToString());

                    decimal TotalBilling = 0, TotalAUM = 0, TotalNGAA = 0;
                    int TotalRowIndex = 0;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;


                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalBilling = TotalBilling + BillingVal;

                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalAUM = TotalAUM + AUMVal;


                            TextBox NGAAText = (TextBox)gvrow.FindControl("txtNGAA");
                            decimal NGAAVal = Convert.ToDecimal(NGAAText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalNGAA = TotalNGAA + NGAAVal;
                        }

                        if (totalBilling != "False")
                        {
                            TotalRowIndex = gvrow.RowIndex;
                            //string ActualVal = gvrow.Cells[4].Text;
                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[6].Text = val.ToString();
                        }


                        #region  Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeAum");
                        //    if (!cbBilling.Checked)
                        //    {
                        //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    }
                        //}

                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = gvrow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }
                        string vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        int index1 = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            Billingtext1.BackColor = System.Drawing.Color.White;


                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        //}
                        //  int index = gvrow.RowIndex;
                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            aumtext1.BackColor = System.Drawing.Color.White;


                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        //if (gvrow.Cells[17].Text == "True")
                        //{
                        //    if (Convert.ToDecimal(gvrow.Cells[6].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "")) != 0)
                        //        Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    else
                        //        Billingtext1.BackColor = System.Drawing.Color.White;


                        //    if (Convert.ToDecimal(gvrow.Cells[8].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "")) != 0)
                        //        aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    else
                        //        aumtext1.BackColor = System.Drawing.Color.White;
                        //}



                        #endregion

                    }

                    tet = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    tet.Text = TotalBilling.ToString();
                    addFormat(tet, "Billing");


                    tet = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtNGAA");
                    tet.Text = currencyFormat(TotalNGAA.ToString());


                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowIndex].Cells[6].Text = currencyFormat(Totalval.ToString());

                    //string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowIndex].Cells[8].Text = currencyFormat(Totalval.ToString());

                    #region TotalTestcolor
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(TotalRowIndex) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(value) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(value) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(TotalActual) < value)
                    {
                        if (Convert.ToDecimal(TotalActual) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }


                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));

                    if (Convert.ToDecimal(TotalRowIndex) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(value) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(value) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(TotalActual) < aumval1)
                    {
                        if (Convert.ToDecimal(TotalActual) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    #endregion
                }
            }

        }
        catch { }
        //  }
        //   ViewState["dtData"] = dtData;
        //  BindGrideview(dtData);
    }

    protected void cbExcludeAUM_CheckedChanged(object sender, EventArgs e)
    {
        try
        {
            lblMessage.Text = "";
            CheckBox ck1 = (CheckBox)sender;
            GridViewRow grow = (GridViewRow)ck1.NamingContainer;
            string IdNumb1 = grow.Cells[22].Text;
            string IdNumbTotalAsset = null;
            int rowindex = -1;
            string TotalVal = null;
            string AUMExtra = grow.Cells[26].Text;
            string Actual = grow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            //int a= e.Row.RowIndex;
            //DataTable dtData = (DataTable)ViewState["dtData"];
            //foreach (GridViewRow row in gvBilling.Rows) 
            //{
            //CheckBox chk = (CheckBox)row.Cells[6].FindControl("cbExcludeBilling");
            if (ck1.Checked)
            {
                TextBox tx = (TextBox)grow.Cells[6].FindControl("txtAUM");
                string A = tx.Text;
                string IdNumb = grow.Cells[22].Text;
                string AccountID = grow.Cells[10].Text;
                string SecurityID = grow.Cells[11].Text;
                string AssetClassID = grow.Cells[12].Text;
                string BillingID = grow.Cells[13].Text;
                string ACLevelBillingExceptionFlg = grow.Cells[14].Text;
                string ACLevelAUMExceptionFlg = grow.Cells[15].Text;
                string AssetLevelFlg = grow.Cells[16].Text;
                string TotalLevelFlg = grow.Cells[17].Text;
                string BillingExceptionId = grow.Cells[18].Text;
                string BillingFeeExceptionId = grow.Cells[19].Text;

                string BillingExceptionType = grow.Cells[20].Text;
                string AUMExceptionType = grow.Cells[21].Text;
                // string AUMExtra = grow.Cells[25].Text;


                //     string AUMExceptionType = row.Cells[21].Text; 
                if (AssetLevelFlg == "True")    //  for Total Asset class ID   set checkbox value = true
                {


                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;

                        if (AssetClassID == AssetClassIDNew)
                        {
                            //  IdNumbTotalAsset = gvrow.Cells[21].Text;
                            rowindex = gvrow.RowIndex;
                            TextBox Aumtxt = (TextBox)gvrow.FindControl("txtAUM");
                            Aumtxt.Text = "0";
                            Aumtxt.Enabled = false;
                            CheckBox cbAum = (CheckBox)gvrow.FindControl("cbExcludeAum");
                            // cbAum.Checked = true;
                            cbAum.Enabled = false;
                            gvrow.Cells[15].Text = "True";

                            gvrow.Cells[43].Text = "3";

                            grow.Cells[46].Text = "";

                            gvrow.Cells[46].Text = "";   // positive AUM
                            // gvrow.Cells[44].Text = "";  // positive billing
                            //   gvrow.Cells[32].Text = "";  //sub Billing 
                            gvrow.Cells[33].Text = "";   // sun AUM

                            if (AssetLevelFlgNew == "True")
                            {
                                cbAum.Enabled = true;
                                grow.Cells[46].Text = "";
                            }
                        }


                    }

                    decimal TotalBilling = 0, TotalAUM = 0;
                    int TotalRowIndex = 0;
                    string vActualValue1 = null;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;


                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TotalBilling = TotalBilling + BillingVal;

                            if (BillingVal > 0)
                            {
                                TotalAUM = TotalAUM + AUMVal;
                            }
                            else
                            {
                                TotalAUM = TotalAUM + AUMVal;
                            }
                        }

                        if (totalBilling != "False")
                        {
                            TotalRowIndex = gvrow.RowIndex;
                            //string ActualVal = gvrow.Cells[4].Text;
                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[8].Text = val.ToString();
                        }


                        #region  Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////  ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        int index1 = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            Billingtext1.BackColor = System.Drawing.Color.White;


                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        //}
                        //  int index = gvrow.RowIndex;
                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            aumtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }



                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = gvrow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }


                        #endregion

                    }


                    TextBox text = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    text.Text = TotalAUM.ToString();
                    addFormat(text, "AUM");

                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    //AUM
                    string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowIndex].Cells[8].Text = currencyFormat(Totalval.ToString());

                    //Billing
                    // TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowIndex].Cells[6].Text = currencyFormat(Totalval.ToString());

                    #region totalTextcolor
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;


                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    #endregion



                }
                else
                {

                    int roindex = grow.RowIndex;
                    int TotalRowIndex2 = 0;
                    TextBox Billingtxt = (TextBox)grow.FindControl("txtAUM");
                    Billingtxt.Text = "0";
                    Billingtxt.Enabled = false;

                    decimal newBilling = 0, TotalValue = 0;
                    grow.Cells[46].Text = "";
                    grow.Cells[33].Text = "";
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;
                        //if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True")
                        //{
                        //    // IdNumbTotalAsset = gvrow.Cells[21].Text;
                        //    rowindex = gvrow.RowIndex;
                        //    TextBox Billingtx = (TextBox)gvrow.FindControl("txtAUM");
                        //    TotalVal = Billingtx.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                        //    break;
                        //}
                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew != "True" && totalBilling != "True")
                        {

                            TextBox Billingtx = (TextBox)gvrow.FindControl("txtAUM");
                            decimal val = Convert.ToDecimal(Billingtx.Text.Replace("$", "").Replace(" ", "").Replace("(", "-").Replace(")", ""));
                            TotalValue = TotalValue + val;

                        }

                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True" && totalBilling != "True")
                        {
                            TotalRowIndex2 = gvrow.RowIndex;
                        }



                    }
                    TextBox text = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtAUM");
                    text.Text = TotalValue.ToString();
                    addFormat(text, "AUM");

                    //if (Convert.ToDecimal(BillingExtra) >= 0)
                    //{
                    //if (Convert.ToDecimal(Actual) > 0)
                    //{
                    //    newBilling = Convert.ToDecimal(TotalVal) - Convert.ToDecimal(Actual);
                    //}
                    //else
                    //{
                    //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(Actual);
                    //}
                    //}
                    //else
                    //{
                    //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(BillingExtra);
                    //}

                    //if (Convert.ToDecimal(Actual) > 0 && Convert.ToDecimal(TotalVal) > 0)
                    //{
                    //    newBilling = Convert.ToDecimal(TotalVal) - Convert.ToDecimal(Actual);
                    //}
                    //else if (Convert.ToDecimal(Actual) < 0 && Convert.ToDecimal(TotalVal) < 0)
                    //{
                    //    Actual = Actual.Replace("-", "");
                    //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(Actual);
                    //}
                    //else if (Convert.ToDecimal(Actual) > 0 && Convert.ToDecimal(TotalVal) < 0)
                    //{

                    //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(Actual);
                    //}
                    //else if (Convert.ToDecimal(Actual) < 0 && Convert.ToDecimal(TotalVal) > 0)
                    //{
                    //    Actual = Actual.Replace("-", "");
                    //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(Actual);
                    //}
                    //else if (Convert.ToDecimal(TotalVal) == 0)
                    //{

                    //    newBilling = Convert.ToDecimal(TotalVal) - Convert.ToDecimal(Actual);
                    //}





                    decimal TotalBilling = 0, TotalAUM = 0;
                    int TotalRowIndex = 0;
                    string vActualValue1 = null;

                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;


                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TotalAUM = TotalAUM + AUMVal;

                            //if (AUMVal > 0)
                            //{
                            TotalBilling = TotalBilling + BillingVal;
                            //}
                            //else
                            //{
                            //    TotalBilling = TotalBilling + AUMVal;
                            //}


                        }

                        if (totalBilling != "False")
                        {
                            TotalRowIndex = gvrow.RowIndex;
                            //string ActualVal = gvrow.Cells[4].Text;
                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[8].Text = val.ToString();
                        }
                        //string totalBilling0 = grow.Cells[15].Text;
                        //string totalBilling1 = grow.Cells[16].Text;
                        //string totalBilling2 = grow.Cells[17].Text;
                        //TextBox BilText = (TextBox)gvrow.FindControl("txtBilling");
                        //TotalBilling = TotalBilling + Convert.ToDecimal(BilText.Text);
                        //int id = gvrow.RowIndex;
                        #region  Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                        //    if (!cbBilling.Checked)
                        //    {
                        //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    }
                        //}

                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    if (roindex != index)
                        //    {
                        //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    }
                        //}

                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = gvrow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }


                        vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        int index1 = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                        {
                            Billingtext1.BackColor = System.Drawing.Color.White;
                        }

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        //}
                        //  int index = gvrow.RowIndex;
                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        #endregion


                    }
                    text = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    text.Text = TotalAUM.ToString();
                    addFormat(text, "AUM");

                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowIndex].Cells[8].Text = currencyFormat(Totalval.ToString());

                    //TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowIndex].Cells[6].Text = currencyFormat(Totalval.ToString());

                    #region totalTextcolor
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;


                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }


                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;


                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    #endregion



                }


            }
            else
            {
                string AssetClassID = grow.Cells[12].Text;
                string AssetLevelFlg = grow.Cells[16].Text;
                if (AssetLevelFlg == "True")                 //  for Total Asset class ID  set checkbox value = false
                {


                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        if (AssetClassID == AssetClassIDNew)
                        {
                            //  IdNumbTotalAsset = gvrow.Cells[21].Text;
                            rowindex = gvrow.RowIndex;
                            TextBox AUMtxt = (TextBox)gvrow.FindControl("txtAUM");
                            //if (Convert.ToDecimal(gvrow.Cells[25].Text) == 0)
                            AUMtxt.Text = gvrow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                            //else
                            //    AUMtxt.Text = gvrow.Cells[24].Text;

                            AUMtxt.Enabled = true;
                            CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeAum");
                            cbBilling.Checked = false;
                            cbBilling.Enabled = true;
                            // gvrow.Cells[42].Text = "";
                            gvrow.Cells[43].Text = ""; // added -20_2_2020
                            gvrow.Cells[26].Text = "0";
                            gvrow.Cells[46].Text = "";
                            gvrow.Cells[44].Text = "";

                            addFormat(AUMtxt, "AUM");
                        }
                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew != "True")
                        {
                            gvrow.Cells[43].Text = "";
                        }
                        else
                        {
                            //gvrow.Cells[42].Text = "";
                        }
                    }

                    decimal TotalAUM = 0, TotalBilling = 0;
                    int TotalRowIndex = 0;
                    string vActualValue1 = null;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;


                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                            TotalBilling = TotalBilling + BillingVal;

                            if (AUMVal > 0)
                                TotalAUM = TotalAUM + AUMVal;
                            else
                                TotalAUM = TotalAUM + AUMVal;
                        }

                        if (totalBilling != "False")
                        {
                            TotalRowIndex = gvrow.RowIndex;
                            //string ActualVal = gvrow.Cells[4].Text;
                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[8].Text = val.ToString();
                        }
                        //string totalBilling0 = grow.Cells[15].Text;
                        //string totalBilling1 = grow.Cells[16].Text;
                        //string totalBilling2 = grow.Cells[17].Text;
                        //TextBox BilText = (TextBox)gvrow.FindControl("txtBilling");
                        //TotalBilling = TotalBilling + Convert.ToDecimal(BilText.Text);
                        //int id = gvrow.RowIndex;

                        #region AUM Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }

                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = gvrow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }
                        vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        int index1 = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            Billingtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        //}
                        //  int index = gvrow.RowIndex;
                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            aumtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        #endregion

                    }

                    TextBox tx = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    tx.Text = TotalAUM.ToString();
                    addFormat(tx, "AUM");

                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowIndex].Cells[8].Text = currencyFormat(Totalval.ToString());

                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowIndex].Cells[6].Text = currencyFormat(Totalval.ToString());


                    #region totalTextcolor
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;


                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    #endregion


                }
                else
                {

                    int roindex = grow.RowIndex;
                    TextBox Aumtxt = (TextBox)grow.FindControl("txtAUM");
                    Aumtxt.Enabled = true;

                    Aumtxt.Text = grow.Cells[4].Text;
                    //else
                    //    Aumtxt.Text = grow.Cells[25].Text;

                    grow.Cells[26].Text = "";

                    addFormat(Aumtxt, "AUM");


                    decimal TotalValue = 0;
                    int TotalRowIndex2 = 0;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;

                        //if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True")
                        //{
                        //    // IdNumbTotalAsset = gvrow.Cells[21].Text;
                        //    rowindex = gvrow.RowIndex;
                        //    TextBox Billingtx = (TextBox)gvrow.FindControl("txtAUM");
                        //    TotalVal = Billingtx.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                        //    break;
                        //}

                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew != "True" && totalBilling != "True")
                        {
                            TextBox Billingtx = (TextBox)gvrow.FindControl("txtAUM");
                            decimal val = Convert.ToDecimal(Billingtx.Text.Replace("$", "").Replace(" ", "").Replace("(", "-").Replace(")", ""));
                            TotalValue = TotalValue + val;
                        }

                        if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True" && totalBilling != "True")
                        {
                            TotalRowIndex2 = gvrow.RowIndex;
                        }

                    }

                    TextBox tx = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtAUM");
                    tx.Text = TotalValue.ToString();
                    addFormat(tx, "AUM");

                    decimal TotalAUM = 0, TotalBilling = 0;
                    int TotalRowIndex = 0;
                    string vActualValue1 = null;
                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;


                        if (AssetLevelFlgNew == "True")
                        {
                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            TotalBilling = TotalBilling + BillingVal;

                            if (AUMVal > 0)
                            {
                                TotalAUM = TotalAUM + AUMVal;
                            }
                            else
                            {
                                TotalAUM = TotalAUM + AUMVal;
                            }
                        }

                        if (totalBilling != "False")
                        {
                            TotalRowIndex = gvrow.RowIndex;
                            //string ActualVal = gvrow.Cells[4].Text;
                            //decimal val = Convert.ToDecimal(ActualVal.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                            //gvrow.Cells[8].Text = val.ToString();
                        }
                        //string totalBilling0 = grow.Cells[15].Text;
                        //string totalBilling1 = grow.Cells[16].Text;
                        //string totalBilling2 = grow.Cells[17].Text;
                        //TextBox BilText = (TextBox)gvrow.FindControl("txtBilling");
                        //TotalBilling = TotalBilling + Convert.ToDecimal(BilText.Text);
                        //int id = gvrow.RowIndex;
                        #region AUM Colour
                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                        string txtBillPer = tx1.Text;

                        string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                        int TotalRowIndex1 = gvrow.RowIndex;
                        //Billing Negative Value Change
                        if (gvrow.Cells[32].Text != "")
                        {
                            // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            gvBilling.Rows[1].FindControl("txtBilling");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //Billing PositiveValue Change
                        if (gvrow.Cells[44].Text != "")
                        {
                            //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Negative Value change
                        if (gvrow.Cells[33].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        //AUM Positive Value Change
                        if (gvrow.Cells[46].Text != "")
                        {
                            // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        ////ddlAdminCategory
                        //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                        //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        ////ddlNonGAServiceType
                        //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                        //{
                        //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                        //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}
                        if (txtBillPer != vBillingFeePct)
                        {
                            TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                        string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");

                        //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                        //if (BillingType != "1" && BillingType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                        //    if (!cbBilling.Checked)
                        //    {
                        //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                        //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //    }
                        //}

                        //if (AUMType != "1" && AUMType != "")
                        //{
                        //    int index = gvrow.RowIndex;
                        //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                        //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        //}

                        vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                        int index1 = gvrow.RowIndex;
                        TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                        decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != val)
                            Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                            Billingtext1.BackColor = System.Drawing.Color.White;

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                        {
                            if (Convert.ToDecimal(vActualValue1) < val)
                                Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }

                        //}
                        //  int index = gvrow.RowIndex;
                        TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                        decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                        if (Convert.ToDecimal(vActualValue1) != aumval)
                            aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        else
                        {
                            aumtext1.BackColor = System.Drawing.Color.White;
                        }

                        if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }
                        else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                        {
                            if (Convert.ToDecimal(vActualValue1) < aumval)
                                aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                        }




                        if (PerValue != "1" && PerValue != "")
                        {
                            int index = gvrow.RowIndex;
                            TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                            text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                        }
                        string diffflag = gvrow.Cells[51].Text;
                        if (diffflag == "1")
                        {
                            gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                        }
                        #endregion

                    }
                    TextBox textbox = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    textbox.Text = TotalAUM.ToString();
                    addFormat(textbox, "AUM");

                    CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeBilling");
                    CbBilling.Visible = false;

                    CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex].FindControl("cbExcludeAUM");
                    CbAUM.Visible = false;

                    string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
                    gvBilling.Rows[TotalRowIndex].Cells[8].Text = currencyFormat(Totalval.ToString());

                    //   TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                    Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
                    gvBilling.Rows[TotalRowIndex].Cells[6].Text = currencyFormat(Totalval.ToString());


                    #region totalTextcolor
                    TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != value)
                        Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        Billingtex.BackColor = System.Drawing.Color.White;

                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
                    {
                        if (Convert.ToDecimal(vActualValue1) < value)
                            Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }


                    TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
                    if (Convert.ToDecimal(vActualValue1) != aumval1)
                        aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    else
                        aumtex.BackColor = System.Drawing.Color.White;


                    if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }
                    else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
                    {
                        if (Convert.ToDecimal(vActualValue1) < aumval1)
                            aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
                    }

                    #endregion



                    //if (Convert.ToDecimal(BillingExtra) >= 0)
                    //{
                    //if (Convert.ToDecimal(Actual) > 0)
                    //{
                    //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(Actual);
                    //}
                    //else
                    //{
                    //    newBilling = Convert.ToDecimal(TotalVal) - Convert.ToDecimal(Actual);
                    //}
                    //}
                    //else
                    //{
                    //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(BillingExtra);
                    //}





                }

            }





        }
        catch { }
        //  }
        //     ViewState["dtData"] = dtData;
        //  BindGrideview(dtData);
    }

    protected void txtBilling_TextChanged(object sender, EventArgs e)
    {
        try
        {
            //lblSavePopUp.Text = "changed";
            lblMessage.Text = "";
            string IdNumbTotalAsset = null, TotalVal = null;
            TextBox ck1 = (TextBox)sender;
            int rowindex = -1;
            GridViewRow grow = (GridViewRow)ck1.NamingContainer;
            string BillingNew = ck1.Text.Replace("$", "").Replace(",", "").Replace(" ", "").Replace("(", "-").Replace(")", "");

            string IdNumb1 = grow.Cells[22].Text;
            string BillingExtra = grow.Cells[25].Text;

            string AssetLevelFlgChange = grow.Cells[16].Text;

            string AssetClassID = grow.Cells[12].Text;
            string AssetLevelFlg = grow.Cells[16].Text;

            string vActual = grow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace(" ", "").Replace("(", "-").Replace(")", "");
            string vActual1 = grow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace(" ", "").Replace("(", "-").Replace(")", "");

            //TextBox tx = (TextBox)grow.FindControl("txtBillPer");
            //string txtBillPer=tx.Text;

            //string vBillingFeePct = grow.Cells[27].Text;  //30  BillingFeePct 


            if (AssetLevelFlgChange == "False")
            {
                //if (BillingNew.All(char.IsDigit))
                //{
                if (BillingNew == "")
                {
                    BillingNew = vActual;
                    ck1.Text = BillingNew;
                }
                Decimal newBilling = 0;
                if (Convert.ToDecimal(BillingExtra) != 0)
                {
                    if (Convert.ToDecimal(BillingNew) != Convert.ToDecimal(BillingExtra))
                    {
                        if (Convert.ToDecimal(BillingNew) < 0) //  -ve
                        {
                            newBilling = Convert.ToDecimal(vActual) + Decimal.Parse(BillingNew);

                            if (Convert.ToDecimal(newBilling) < 0 && Convert.ToDecimal(vActual) > 0)
                            {
                                ck1.Text = newBilling.ToString();
                            }
                            //else if ( Convert.ToDecimal(vActual) < 0)
                            //{
                            //    ck1.Text = newBilling.ToString();
                            //}
                            else
                            {
                                ck1.Text = newBilling.ToString();
                            }
                            grow.Cells[32].Text = BillingNew;
                            grow.Cells[44].Text = "";
                        }
                        else if (Convert.ToDecimal(BillingNew) > 0)  //+ve
                        {
                            grow.Cells[32].Text = "";
                            grow.Cells[44].Text = BillingNew;
                        }
                        else
                        {
                            grow.Cells[44].Text = "";
                            grow.Cells[32].Text = "";
                            ck1.Text = vActual;
                        }
                        //else
                        //{
                        //    newBilling = Convert.ToDecimal(BillingNew);
                        //}
                    }

                }
                else
                {
                    if (Convert.ToDecimal(BillingNew) != Convert.ToDecimal(vActual))
                    {
                        if (Convert.ToDecimal(BillingNew) < 0)
                        {
                            newBilling = Convert.ToDecimal(vActual) + Convert.ToDecimal(BillingNew);
                            ck1.Text = newBilling.ToString();
                            grow.Cells[32].Text = BillingNew;
                            grow.Cells[44].Text = "";
                        }
                        else if (Convert.ToDecimal(BillingNew) > 0)
                        {
                            grow.Cells[32].Text = "";
                            grow.Cells[44].Text = BillingNew;
                        }
                        else
                        {
                            grow.Cells[44].Text = "";
                            grow.Cells[32].Text = "";
                            ck1.Text = vActual;
                        }
                        //else
                        //{
                        //    newBilling = Convert.ToDecimal(BillingNew);
                        //}
                    }

                }


                addFormat(ck1, "Billing");

                //string AssetClassID = grow.Cells[11].Text;
                //string AssetLevelFlg = grow.Cells[15].Text;

                int TotalRowIndex = 0;
                decimal BillingAssetValue = 0;
                decimal TotalBilling = 0;

                foreach (GridViewRow gvrow in gvBilling.Rows)   // For updating Total asset Value
                {
                    string AssetClassIDNew = gvrow.Cells[12].Text;
                    string AssetLevelFlgNew = gvrow.Cells[16].Text;
                    string totalBilling = gvrow.Cells[17].Text;
                    if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "False")
                    {

                        TextBox tx1 = (TextBox)gvrow.FindControl("txtBilling");
                        if (tx1.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "") == "")
                        {
                            BillingAssetValue = BillingAssetValue + Convert.ToDecimal(vActual);
                        }
                        else
                        {
                            BillingAssetValue = BillingAssetValue + Convert.ToDecimal(tx1.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                        }

                        //IdNumbTotalAsset = gvrow.Cells[21].Text;
                        //rowindex = gvrow.RowIndex;
                        //TextBox tx = (TextBox)gvrow.Cells[rowindex].FindControl("txtBilling");
                        //TotalVal = tx.Text;
                        //break;
                    }
                    if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True")
                    {
                        IdNumbTotalAsset = gvrow.Cells[22].Text;
                        rowindex = gvrow.RowIndex;
                    }

                    //if (AssetLevelFlgNew == "True")
                    //{
                    //    TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                    //    TotalBilling = TotalBilling + Convert.ToDecimal(BillingText.Text);
                    //}

                    //if (totalBilling != "False")
                    //{
                    //    TotalRowIndex = gvrow.RowIndex;
                    //}

                }
                TextBox tex = (TextBox)gvBilling.Rows[rowindex].FindControl("txtBilling");
                tex.Text = BillingAssetValue.ToString();
                addFormat(tex, "Billing");



                //int TotalRowIndex1 = 0;
                //foreach (GridViewRow gvrow in gvBilling.Rows)     // For updating Total Value
                //{
                //    string AssetClassIDNew = gvrow.Cells[12].Text;
                //    string AssetLevelFlgNew = gvrow.Cells[16].Text;
                //    string totalBilling = gvrow.Cells[17].Text;

                //    if (AssetLevelFlgNew == "True")
                //    {
                //        TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                //        TotalBilling = TotalBilling + Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                //    }

                //    if (totalBilling != "False")
                //    {
                //        TotalRowIndex = gvrow.RowIndex;
                //    }

                //    #region AUM Colour
                //    TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                //    string txtBillPer = tx1.Text;

                //    string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                //    TotalRowIndex1 = gvrow.RowIndex;
                //    //Billing Negative Value Change
                //    if (gvrow.Cells[32].Text != "")
                //    {
                //        // gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    //Billing PositiveValue Change
                //    if (gvrow.Cells[44].Text != "")
                //    {
                //        //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    //AUM Negative Value change
                //    if (gvrow.Cells[33].Text != "")
                //    {
                //        // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    //AUM Positive Value Change
                //    if (gvrow.Cells[46].Text != "")
                //    {
                //        // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    if (txtBillPer != vBillingFeePct)
                //    {
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                //    string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                //    string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");

                //    //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                //    if (BillingType != "1" && BillingType != "")
                //    {
                //        int index = gvrow.RowIndex;
                //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }

                //    if (AUMType != "1" && AUMType != "")
                //    {
                //        int index = gvrow.RowIndex;
                //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }

                //    if (PerValue != "1" && PerValue != "")
                //    {
                //        int index = gvrow.RowIndex;
                //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    #endregion



                //}
                //tex = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                //tex.Text = TotalBilling.ToString();
                //addFormat(tex, "Billing");

                //string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                //decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalBilling;
                //gvBilling.Rows[TotalRowIndex].Cells[6].Text = Totalval.ToString();


                //  gvBilling.
                //         TextBox tx = (TextBox)gvBilling.Cells[5].FindControl("txtBilling");
                //if (rowindex != -1)
                //{
                //if (Convert.ToDecimal(BillingNew) >= 0)
                //{
                //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(BillingNew);
                //}
                //else
                //{
                //    newBilling = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(BillingNew);
                //}

                //  ((TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling")).Text = TotalBilling.ToString();
                //}
            }
            else
            {
                int TotalRowIndex1 = 0;
                bool Flags = false;
                if (Convert.ToDecimal(BillingNew) > 0)    // for Positive Amount
                {
                    grow.Cells[44].Text = BillingNew;
                    grow.Cells[32].Text = "";
                    ck1.Text = BillingNew;

                    //string AssetClassID = grow.Cells[11].Text;
                    //string AssetLevelFlg = grow.Cells[15].Text;

                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;

                        if (AssetLevelFlgNew != "True" && AssetClassID == AssetClassIDNew)
                        {


                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            BillingText.Enabled = false;


                            CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                            //if (!cbBilling.Enabled)
                            //   cbBilling.Enabled = false;

                            gvrow.Cells[44].Text = "";

                            gvrow.Cells[42].Text = "1";
                        }
                        if (AssetLevelFlgNew == "True" && AssetClassID == AssetClassIDNew)
                        {
                            gvrow.Cells[42].Text = "1";
                        }
                    }
                    addFormat(ck1, "Billing");

                }
                else if (Convert.ToDecimal(BillingNew) < 0)   // for -ve Amount 
                {

                    decimal total = 0;
                    int Index = 0;
                    // string newAmount = BillingNew;

                    if ((Convert.ToDecimal(vActual) + Convert.ToDecimal(BillingNew)) > 0)
                    {
                        grow.Cells[44].Text = "";
                        grow.Cells[32].Text = BillingNew;
                        ck1.Text = BillingNew;

                    }
                    else
                    {
                        grow.Cells[42].Text = "";
                        //grow.Cells[46].Text = "";
                        grow.Cells[44].Text = "";
                        grow.Cells[32].Text = "";
                        ck1.Text = vActual;
                        addFormat(ck1, "Billing");
                        Flags = true;
                    }

                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;

                        if (AssetLevelFlgNew != "True" && AssetClassID == AssetClassIDNew)
                        {
                            TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                            total = total + Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                            if (!cbBilling.Enabled)
                                BillingText.Enabled = true;

                            else if (gvrow.Cells[23].Text == "True")
                                BillingText.Enabled = false;

                            else if (cbBilling.Enabled)
                                BillingText.Enabled = true;

                            gvrow.Cells[32].Text = "";

                            gvrow.Cells[42].Text = "";
                        }
                        if (AssetLevelFlgNew == "True" && AssetClassID == AssetClassIDNew)
                        {
                            Index = gvrow.RowIndex;
                            if (!Flags)
                                gvrow.Cells[42].Text = "2";
                            else
                                gvrow.Cells[42].Text = "";
                        }

                    }

                    TextBox tex = (TextBox)gvBilling.Rows[Index].FindControl("txtBilling");
                    if ((Convert.ToDecimal(vActual) + Convert.ToDecimal(BillingNew)) > 0)
                    {
                        tex.Text = (Convert.ToDecimal(BillingNew) + total).ToString();
                    }
                    else
                    {
                        tex.Text = vActual.ToString();
                    }
                    addFormat(tex, "Billing");
                }
                else
                {
                    grow.Cells[42].Text = "";
                    grow.Cells[46].Text = "";
                    ck1.Text = vActual;
                    addFormat(ck1, "Billing");
                }

                //  addFormat(tex, "Billing");


            }


            decimal TotalBilling2 = 0, TotalAUM = 0;
            int TotalRowIndex2 = 0;
            string vActualValue1 = null;
            foreach (GridViewRow gvrow in gvBilling.Rows)
            {
                string AssetClassIDNew = gvrow.Cells[12].Text;
                string AssetLevelFlgNew = gvrow.Cells[16].Text;
                string totalBilling = gvrow.Cells[17].Text;


                if (AssetLevelFlgNew == "True")
                {
                    TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                    TotalBilling2 = TotalBilling2 + Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                    TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                    decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                    TotalAUM = TotalAUM + AUMVal;

                }

                if (totalBilling != "False")
                {
                    TotalRowIndex2 = gvrow.RowIndex;
                }
                //string totalBilling0 = grow.Cells[15].Text;
                //string totalBilling1 = grow.Cells[16].Text;
                //string totalBilling2 = grow.Cells[17].Text;
                //TextBox BilText = (TextBox)gvrow.FindControl("txtBilling");
                //TotalBilling = TotalBilling + Convert.ToDecimal(BilText.Text);
                //int id = gvrow.RowIndex;


                #region AUM Colour Asset Level

                TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                string txtBillPer = tx1.Text;

                string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                int TotalRowIndex = gvrow.RowIndex;
                if (gvrow.Cells[32].Text != "")
                {
                    //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                //Billing PositiveValue Change
                if (gvrow.Cells[44].Text != "")
                {
                    //   gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                //AUM Negative Value change
                if (gvrow.Cells[33].Text != "")
                {
                    // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                //AUM Positive Value Change
                if (gvrow.Cells[46].Text != "")
                {
                    //   gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                ////Admin Category
                //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex].FindControl("ddlAdminCategory");
                //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                ////ddlNonGAServiceType
                //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex].FindControl("ddlNonGAServiceType");
                //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                if (txtBillPer != vBillingFeePct)
                {
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "").Replace(",", "");
                string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "").Replace(",", "");
                string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "").Replace(",", "");

                //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                //if (BillingType != "1" && BillingType != "")
                //{
                //    int index = gvrow.RowIndex;
                //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}

                //if (AUMType != "1" && AUMType != "")
                //{
                //    int index = gvrow.RowIndex;
                //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}

                if (PerValue != "1" && PerValue != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string diffflag = gvrow.Cells[51].Text;
                if (diffflag == "1")
                {
                    gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                }

                vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                int index1 = gvrow.RowIndex;
                TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != val)
                    Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");

                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }



                //}
                //  int index = gvrow.RowIndex;
                TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != aumval)
                    aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }

                #endregion


            }
            TextBox text = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtBilling");
            text.Text = TotalBilling2.ToString();
            addFormat(text, "Billing");

            string TotalActual1 = gvBilling.Rows[TotalRowIndex2].Cells[4].Text;
            decimal Totalval1 = Convert.ToDecimal(TotalActual1.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling2;
            gvBilling.Rows[TotalRowIndex2].Cells[6].Text = currencyFormat(Totalval1.ToString());


            TotalActual1 = gvBilling.Rows[TotalRowIndex2].Cells[4].Text;
            Totalval1 = Convert.ToDecimal(TotalActual1.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
            gvBilling.Rows[TotalRowIndex2].Cells[8].Text = currencyFormat(Totalval1.ToString());


            #region totalTextcolor
            TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtBilling");
            decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));

            if (Convert.ToDecimal(vActualValue1) != value)
                Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                Billingtex.BackColor = System.Drawing.Color.White;

            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtAUM");
            decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
            if (Convert.ToDecimal(vActualValue1) != aumval1)
                aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                aumtex.BackColor = System.Drawing.Color.White;

            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            #endregion


            grow.FindControl("cbExcludeBilling").Focus();



            // }
            //  else
            //{
            //    ((TextBox)grow.FindControl("txtBilling")).Text = BillingExtra;
            //      //grow.Cells[4].Text="*";

            //}
        }
        catch { }
    }

    protected void txtAUM_TextChanged(object sender, EventArgs e)
    {
        try
        {
            //lblSavePopUp.Text = "changed";
            lblMessage.Text = "";
            string IdNumbTotalAsset = null, TotalVal = null;
            int rowindex = -1;
            TextBox txtAum = (TextBox)sender;
            GridViewRow grow = (GridViewRow)txtAum.NamingContainer;
            string AUMNew = txtAum.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            string IdNumb1 = grow.Cells[22].Text;
            string AUMExtra = grow.Cells[26].Text;
            decimal AUMAssetValue = 0;
            decimal TotalAUM = 0;
            int ChildRowindex = grow.RowIndex;


            string Bill = grow.Cells[10].Text;
            // string AssetLevelFlg = grow.Cells[15].Text;

            string AssetClassID = grow.Cells[12].Text;
            string AssetLevelFlg = grow.Cells[16].Text;
            string vActual = grow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            //TextBox tx1 = (TextBox)grow.FindControl("txtBillPer");
            //string txtBillPer = tx1.Text;

            //string vBillingFeePct = grow.Cells[27].Text;  //30  BillingFeePct 


            if (AssetLevelFlg == "False")
            {
                //if (AUMNew.All(char.IsDigit))
                //{
                if (AUMNew == "")
                {
                    AUMNew = vActual;
                    txtAum.Text = AUMNew;
                }
                Decimal AUM = 0;

                //if (AUMExtra != "")
                //{
                //    if (Convert.ToDecimal(AUMNew) != Convert.ToDecimal(AUMExtra))
                //    {
                //        //  if (Convert.ToDecimal(AUMNew) > 0)
                //        //   decimal a = Convert.ToDecimal(AUMNew);
                //        if (Convert.ToDecimal(AUMNew) < 0)
                //        {
                //            AUM = Convert.ToDecimal(AUMExtra) + Convert.ToDecimal(AUMNew);      // change value in textbox
                //            txtAum.Text = AUM.ToString();
                //            grow.Cells[32].Text = AUMNew;
                //        }
                //        //else
                //        //{
                //        //    AUM = Convert.ToDecimal(AUMExtra);

                //        //}
                //        //else
                //        //    AUM = Convert.ToDecimal(AUMExtra) - Convert.ToDecimal(AUMNew);
                //    }
                //}
                //else
                //{
                if (Convert.ToDecimal(AUMNew) != Convert.ToDecimal(vActual))
                {
                    //  if (Convert.ToDecimal(AUMNew) > 0)
                    //   decimal a = Convert.ToDecimal(AUMNew);
                    if (Convert.ToDecimal(AUMNew) < 0)
                    {
                        AUM = Convert.ToDecimal(vActual) + Convert.ToDecimal(AUMNew);      // change value in textbox

                        if (Convert.ToDecimal(AUM) < 0 && Convert.ToDecimal(vActual) > 0)
                        {
                            txtAum.Text = AUM.ToString();
                        }
                        else
                        {
                            txtAum.Text = AUM.ToString();
                        }
                        grow.Cells[33].Text = AUMNew;
                        grow.Cells[46].Text = "";
                    }
                    else if (Convert.ToDecimal(AUMNew) > 0)
                    {
                        grow.Cells[33].Text = "";
                        grow.Cells[46].Text = AUMNew;
                    }
                    else
                    {
                        grow.Cells[33].Text = "";
                        grow.Cells[46].Text = "";
                        txtAum.Text = vActual;
                    }

                    //else
                    //{
                    //    AUM = Convert.ToDecimal(AUMExtra);

                    //}
                    //else
                    //    AUM = Convert.ToDecimal(AUMExtra) - Convert.ToDecimal(AUMNew);
                }

                // }


                //TextBox txt = (TextBox)gvBilling.Rows[rowindex].FindControl("txtAUM");
                //txt.Text = AUMAssetValue.ToString();
                addFormat(txtAum, "AUM");

                //string AssetClassID = grow.Cells[11].Text;
                //string AssetLevelFlg = grow.Cells[15].Text;
                int TotalRowIndex = 0;
                foreach (GridViewRow gvrow in gvBilling.Rows)   // For updating Total asset Value 
                {
                    string AssetClassIDNew = gvrow.Cells[12].Text;
                    string AssetLevelFlgNew = gvrow.Cells[16].Text;
                    string totalAUM = gvrow.Cells[17].Text;
                    if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "False")
                    {
                        TextBox tx = (TextBox)gvrow.FindControl("txtAum");
                        string val = tx.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                        //if(Convert.ToDecimal(val)>0)
                        AUMAssetValue = AUMAssetValue + Convert.ToDecimal(val);
                        //else
                        //    AUMAssetValue = AUMAssetValue - Convert.ToDecimal(val);
                        //IdNumbTotalAsset = gvrow.Cells[21].Text;
                        //rowindex = gvrow.RowIndex;
                        //TextBox tx = (TextBox)gvrow.Cells[rowindex].FindControl("txtAUM");
                        //TotalVal = tx.Text;
                        //break;
                    }
                    if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True")
                    {
                        rowindex = gvrow.RowIndex;
                    }

                }
                TextBox txt = (TextBox)gvBilling.Rows[rowindex].FindControl("txtAUM");
                txt.Text = AUMAssetValue.ToString();
                addFormat(txt, "AUM");

                //if (rowindex != -1)
                //{
                //if (Convert.ToDecimal(AUMNew) >= 0)
                //{
                //    AUM = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(AUMNew);
                //}
                //else
                //{
                //    AUM = Convert.ToDecimal(TotalVal) + Convert.ToDecimal(AUMNew);
                //}
                //int TotalRowIndex1 = 0;
                //foreach (GridViewRow gvrow in gvBilling.Rows)     // For updating Total Value
                //{
                //    string AssetClassIDNew = gvrow.Cells[12].Text;
                //    string AssetLevelFlgNew = gvrow.Cells[16].Text;
                //    string totalBilling = gvrow.Cells[17].Text;

                //    if (AssetLevelFlgNew == "True")
                //    {
                //        TextBox TotalAUMText = (TextBox)gvrow.FindControl("txtAUM");
                //        TotalAUM = TotalAUM + Convert.ToDecimal(TotalAUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                //    }

                //    if (totalBilling != "False")
                //    {
                //        TotalRowIndex = gvrow.RowIndex;
                //    }

                //    #region AUM Colour
                //    TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                //    string txtBillPer = tx1.Text;

                //    string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 
                //    TotalRowIndex1 = gvrow.RowIndex;
                //    //Billing Negative Value Change
                //    if (gvrow.Cells[32].Text != "")
                //    {
                //        //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    //Billing PositiveValue Change
                //    if (gvrow.Cells[44].Text != "")
                //    {
                //        //   gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    //AUM Negative Value change
                //    if (gvrow.Cells[33].Text != "")
                //    {
                //        //  gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    //AUM Positive Value Change
                //    if (gvrow.Cells[46].Text != "")
                //    {
                //        //  gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    if (txtBillPer != vBillingFeePct)
                //    {
                //        TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                //    string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                //    string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");

                //    //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                //    if (BillingType != "1" && BillingType != "")
                //    {
                //        int index = gvrow.RowIndex;
                //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }

                //    if (AUMType != "1" && AUMType != "")
                //    {
                //        int index = gvrow.RowIndex;
                //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }

                //    if (PerValue != "1" && PerValue != "")
                //    {
                //        int index = gvrow.RowIndex;
                //        TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                //        text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //    }
                //    #endregion



                //}
                //txt = (TextBox)gvBilling.Rows[TotalRowIndex].FindControl("txtAUM");
                //txt.Text = TotalAUM.ToString();
                //addFormat(txt, "AUM");

                //string TotalActual = gvBilling.Rows[TotalRowIndex].Cells[4].Text;
                //decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "")) - TotalAUM;
                //gvBilling.Rows[TotalRowIndex].Cells[8].Text = Totalval.ToString();
                //}
            }
            else
            {
                int TotalRowIndex1 = 0;
                bool flags = false;
                if (Convert.ToDecimal(AUMNew) > 0)
                {
                    grow.Cells[33].Text = "";
                    grow.Cells[46].Text = AUMNew;
                    //string AssetClassID = grow.Cells[11].Text;
                    //string AssetLevelFlg = grow.Cells[15].Text;

                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;

                        if (AssetLevelFlgNew != "True" && AssetClassID == AssetClassIDNew)
                        {
                            TextBox BillingText = (TextBox)gvrow.FindControl("txtAUM");
                            BillingText.Enabled = false;
                            gvrow.Cells[43].Text = "1";

                            gvrow.Cells[46].Text = "";

                        }
                        if (AssetLevelFlgNew == "True" && AssetClassID == AssetClassIDNew)
                            gvrow.Cells[43].Text = "1";
                    }
                    addFormat(txtAum, "AUM");
                }
                else if (Convert.ToDecimal(AUMNew) < 0)
                {
                    decimal total = 0;
                    int Index = 0;
                    if ((Convert.ToDecimal(vActual) + Convert.ToDecimal(AUMNew)) > 0)
                    {
                        grow.Cells[33].Text = AUMNew;
                        grow.Cells[46].Text = "";
                        txtAum.Text = AUMNew;
                    }
                    else
                    {
                        grow.Cells[33].Text = "";
                        grow.Cells[46].Text = "";
                        grow.Cells[43].Text = "";
                        txtAum.Text = vActual;
                        addFormat(txtAum, "Billing");
                        flags = true;

                    }

                    // string newAmount = BillingNew;

                    foreach (GridViewRow gvrow in gvBilling.Rows)
                    {
                        string AssetClassIDNew = gvrow.Cells[12].Text;
                        string AssetLevelFlgNew = gvrow.Cells[16].Text;
                        string totalBilling = gvrow.Cells[17].Text;

                        if (AssetLevelFlgNew != "True" && AssetClassID == AssetClassIDNew)
                        {
                            TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                            total = total + Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                            //if(!flags)
                            // gvrow.Cells[43].Text = "2";
                            //else
                            // gvrow.Cells[43].Text = "";

                            CheckBox cbAUM = (CheckBox)gvrow.FindControl("cbExcludeAum");
                            if (!cbAUM.Enabled)
                                AUMText.Enabled = true;

                            else if (gvrow.Cells[23].Text == "True")
                                AUMText.Enabled = false;

                            if (cbAUM.Enabled)
                                AUMText.Enabled = true;
                            //CheckBox cbBilling = (CheckBox)gvrow.FindControl("cbExcludeBilling");
                            //if (!cbBilling.Enabled)
                            //    BillingText.Enabled = true;

                            gvrow.Cells[33].Text = "";

                        }
                        if (AssetLevelFlgNew == "True" && AssetClassID == AssetClassIDNew)
                        {
                            Index = gvrow.RowIndex;
                            if (!flags)
                                gvrow.Cells[43].Text = "2";
                            else
                                gvrow.Cells[43].Text = "";
                        }
                    }

                    TextBox tx = (TextBox)gvBilling.Rows[Index].FindControl("txtAUM");

                    if (!flags)
                    {
                        tx.Text = (Convert.ToDecimal(AUMNew) + total).ToString();
                    }
                    else
                    {
                        tx.Text = vActual;
                    }
                    addFormat(tx, "AUM");
                }
                else
                {
                    grow.Cells[33].Text = "";
                    grow.Cells[46].Text = "";
                    txtAum.Text = vActual;
                    addFormat(txtAum, "AUM");
                }



            }




            int TotalRowIndex2 = 0;
            decimal TotalBilling = 0;
            string vActualValue1 = null;
            foreach (GridViewRow gvrow in gvBilling.Rows)
            {
                string AssetClassIDNew = gvrow.Cells[12].Text;
                string AssetLevelFlgNew = gvrow.Cells[16].Text;
                string totalAUM = gvrow.Cells[17].Text;


                if (AssetLevelFlgNew == "True")
                {
                    TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                    TotalAUM = TotalAUM + Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                    TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                    TotalBilling = TotalBilling + Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                }

                if (totalAUM != "False")
                {
                    TotalRowIndex2 = gvrow.RowIndex;
                }


                #region AUM Colour Asset Level
                TextBox tx1 = (TextBox)gvrow.FindControl("txtBillPer");
                string txtBillPer = tx1.Text;

                string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 
                int TotalRowIndex1 = gvrow.RowIndex;
                if (gvrow.Cells[32].Text != "")
                {
                    //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                //Billing PositiveValue Change
                if (gvrow.Cells[44].Text != "")
                {
                    //   gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                //AUM Negative Value change
                if (gvrow.Cells[33].Text != "")
                {
                    // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                //AUM Positive Value Change
                if (gvrow.Cells[46].Text != "")
                {
                    //   gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                ////ddlAdminCategory
                //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                //{

                //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                ////ddlNonGAServiceType
                //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                if (txtBillPer != vBillingFeePct)
                {
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");
                string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace(",", "");

                //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                //if (BillingType != "1" && BillingType != "")
                //{
                //    int index = gvrow.RowIndex;
                //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}

                //if (AUMType != "1" && AUMType != "")
                //{
                //    int index = gvrow.RowIndex;
                //    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                //    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}

                if (PerValue != "1" && PerValue != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string diffflag = gvrow.Cells[51].Text;
                if (diffflag == "1")
                {
                    gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                }

                vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                int index1 = gvrow.RowIndex;
                TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != val)
                    Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                    Billingtext1.BackColor = System.Drawing.Color.White;

                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }

                //}
                //  int index = gvrow.RowIndex;
                TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != aumval)
                    aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                    aumtext1.BackColor = System.Drawing.Color.White;

                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }

                #endregion

            }
            // txtAum

            TextBox tex = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtAUM");
            tex.Text = TotalAUM.ToString();
            addFormat(tex, "AUM");
            addFormat(txtAum, "AUM");

            CheckBox CbBilling = (CheckBox)gvBilling.Rows[TotalRowIndex2].FindControl("cbExcludeBilling");
            CbBilling.Visible = false;

            CheckBox CbAUM = (CheckBox)gvBilling.Rows[TotalRowIndex2].FindControl("cbExcludeAUM");
            CbAUM.Visible = false;

            string TotalActual = gvBilling.Rows[TotalRowIndex2].Cells[4].Text;
            decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
            gvBilling.Rows[TotalRowIndex2].Cells[8].Text = currencyFormat(Totalval.ToString());

            TotalActual = gvBilling.Rows[TotalRowIndex2].Cells[4].Text;
            Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
            gvBilling.Rows[TotalRowIndex2].Cells[6].Text = currencyFormat(Totalval.ToString());


            #region totalTextcolor
            TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtBilling");
            decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
            if (Convert.ToDecimal(vActualValue1) != value)
                Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                Billingtex.BackColor = System.Drawing.Color.White;

            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex2].FindControl("txtAUM");
            decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
            if (Convert.ToDecimal(vActualValue1) != aumval1)
                aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                aumtex.BackColor = System.Drawing.Color.White;


            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            #endregion

            grow.FindControl("cbExcludeAum").Focus();



            //}
            //else
            //{
            //    ((TextBox)grow.FindControl("txtAUM")).Text = AUMExtra;

            //    //grow.Cells[6].Text = "*";

            //}
        }
        catch
        { }
    }

    protected void txtBillPer_TextChanged(object sender, EventArgs e)
    {
        try
        {
            lblMessage.Text = "";

            TextBox txtPer = (TextBox)sender;
            GridViewRow gvrow1 = (GridViewRow)txtPer.NamingContainer;

            //TextBox tx = (TextBox)gvrow1.FindControl("txtBillPer");
            //string txtBillPer = tx.Text;

            //



            #region AUM Colour Asset Level
            int TotalRowIndex1 = 0;
            string vActualValue1 = null;
            foreach (GridViewRow gvrow in gvBilling.Rows)   // For updating Total asset Value 
            {
                TextBox tx = (TextBox)gvrow.FindControl("txtBillPer");
                string txtBillPer = tx.Text;
                string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                TotalRowIndex1 = gvrow.RowIndex;
                if (gvrow.Cells[32].Text != "")
                {
                    //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (gvrow.Cells[44].Text != "")
                {
                    //   gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (gvrow.Cells[33].Text != "")
                {
                    // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                //AUM Positive Value Change
                if (gvrow.Cells[46].Text != "")
                {
                    //   gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                ////ddlAdminCategory
                //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                //{

                //    DropDownList ddlAdminCategory = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                ////ddlNonGAServiceType
                //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                if (txtBillPer != vBillingFeePct)
                {
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");

                //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                if (BillingType != "1" && BillingType != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (AUMType != "1" && AUMType != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (PerValue != "1" && PerValue != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string diffflag = gvrow.Cells[51].Text;
                if (diffflag == "1")
                {
                    gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                }


                vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                int index1 = gvrow.RowIndex;
                TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != val)
                    Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                {
                    Billingtext1.BackColor = System.Drawing.Color.White;
                }

                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }

                //}
                //  int index = gvrow.RowIndex;


                TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != aumval)
                    aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                    aumtext1.BackColor = System.Drawing.Color.White;


                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }




            }
            #endregion

            #region totalTextcolor
            TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
            decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));


            if (Convert.ToDecimal(vActualValue1) != value)
                Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                Billingtex.BackColor = System.Drawing.Color.White;

            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
            decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
            if (Convert.ToDecimal(vActualValue1) != aumval1)
                aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                aumtex.BackColor = System.Drawing.Color.White;


            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            #endregion




            //        string IdNumbTotalAsset = null, TotalVal = null;
            //        int rowindex = -1;
            //        TextBox txtPer = (TextBox)sender;
            //        GridViewRow grow = (GridViewRow)txtPer.NamingContainer;
            //        rowindex=grow.RowIndex;
            //        string BillPerNew = txtPer.Text;
            //        string IdNumb1 = grow.Cells[21].Text;
            //        string BillPerExtra = grow.Cells[27].Text;
            //        string AssetClassID = grow.Cells[11].Text;
            //        string AssetLevelFlg = grow.Cells[15].Text;

            //      //  string Billing = grow.Cells[27].Text;

            //        TextBox Billingtxt = (TextBox)grow.Cells[rowindex].FindControl("txtBilling");
            //        string BillingOld = Billingtxt.Text;
            //        decimal Bill = 0;

            //        if (BillPerNew == "")
            //            BillPerNew = "0";

            //        if (BillingOld != "0")
            //        {
            //            if (BillPerNew != "0")
            //            {

            //                if (Convert.ToDecimal(BillPerNew) != Convert.ToDecimal(BillPerExtra))
            //                {
            //                    decimal per = (Convert.ToDecimal(BillPerNew) / 100);
            //                    decimal val = per * Convert.ToDecimal(BillingOld);
            //                    Bill = Convert.ToDecimal(BillingOld) - val;
            //                    Math.Round(Bill, 2);
            //                    Billingtxt.Text = Math.Round(Bill, 2).ToString();

            //                }

            //            }
            //        }


            //        decimal TotalBillingNew = -1;
            //        int TotalAssetRowIndex = -1;

            //        foreach (GridViewRow gvrow in gvBilling.Rows)
            //        {
            //            string AssetClassIDNew = gvrow.Cells[11].Text;
            //            string AssetLevelFlgNew = gvrow.Cells[15].Text;
            //            int temp = gvrow.RowIndex;

            //            if (AssetClassID == AssetClassIDNew && AssetLevelFlgNew == "True")
            //            {
            //                IdNumbTotalAsset = gvrow.Cells[21].Text;
            //                TotalAssetRowIndex = gvrow.RowIndex;
            //                TextBox tx = (TextBox)gvrow.Cells[TotalAssetRowIndex].FindControl("txtBilling");
            //                TotalVal = tx.Text;
            //                break;
            //            }
            //        }

            //        if (TotalAssetRowIndex != -1)
            //        {
            //            if (Convert.ToDecimal(Bill) >= 0)
            //            {
            //                TotalBillingNew = Convert.ToDecimal(TotalVal) - Convert.ToDecimal(BillingOld);
            //            }

            //            ((TextBox)gvBilling.Rows[TotalAssetRowIndex].FindControl("txtBilling")).Text = TotalBillingNew.ToString();
            //        }

            ////        ((TextBox)grow.Cells[rowindex + 1].FindControl("txtBilling")).Focus();

            //        ((TextBox)gvBilling.Rows[rowindex + 1].FindControl("txtBilling")).Focus();
        }
        catch { }
    }

    #endregion

    protected void btnSave_Click(object sender, EventArgs e)
    {
        SaveNew();

    }

    public void addFormat(TextBox text, string Type)
    {
        string value = text.Text.Replace(",", "").Replace("$", "").Replace("%", "").Replace("(", "-").Replace(")", "");
        decimal ul = 0;
        if (value == "")
            ul = 0;//text.Text = "";
        else
            ul = Convert.ToDecimal(value);


        if (Type == "Billing")
        {
            text.TextChanged -= txtBilling_TextChanged;
            text.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
            text.TextChanged += txtBilling_TextChanged;
        }
        else
        {
            text.TextChanged -= txtAUM_TextChanged;
            text.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
            text.TextChanged += txtAUM_TextChanged;
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


    #region Not in used
    //else if (SubBilling != "" || vBilling != "" || vBillingPer != "")
    //{
    //    updateBill(gvRow, false, Convert.ToDecimal(SubBilling), Convert.ToDecimal(vBilling), Convert.ToDecimal(vBillingPer), "Billing");
    //}
    //else
    //{

    //}


    //else
    //{
    //    if (BillingFlag != "")
    //        insertRecord(gvRow);

    //    else if (gvRow.Cells[19].Text == "1" && gvRow.Cells[19].Text == "2")
    //        UpdatingDelete(gvRow);

    //    //insertRecord(gvRow);
    //    else
    //    {
    //        CheckBox ExcludeBillingcb = (CheckBox)gvRow.Cells[5].FindControl("cbExcludeBilling");   //5
    //        bool ExcludeBilling = ExcludeBillingcb.Checked;

    //        string vBillingExcludeFlg = gvRow.Cells[22].Text;  //22   BillingExcludeFlg

    //        if (Convert.ToBoolean(vBillingExcludeFlg) != ExcludeBilling)
    //        {
    //            UpdatingDelete(gvRow);
    //        }

    //    }


    //}


    //    if (AUMFlag == "1")    // AUM Removed all exceptions   Assetlevel _+veAmount 
    //    {
    //        if (AumExceptionID != "" && AssetLevelFlg != "True")   //delete
    //        {
    //            UpdatingDelete(gvRow);

    //        }
    //        else if (AumExceptionID != "" && AssetLevelFlg == "True")   //delete
    //        {
    //            if (startDate < dAsOFDate)
    //            {
    //                UpdatingDelete(gvRow);
    //                insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vAUM), 0);
    //            }
    //            else if(startDate == dAsOFDate)
    //            {
    //                updateBill(gvRow, false, 0, Convert.ToDecimal(vBilling), Convert.ToDecimal(vBillingPer),"AUM");
    //            }

    //    }

    //}

    #endregion

    #region not in used
    public void save1()
    {
        //DataTable dtTempData = (DataTable)ViewState["dtTempData"];
        foreach (GridViewRow gvRow in gvBilling.Rows)
        {

            string BillingExceptionID = gvRow.Cells[17].Text;
            string AssetLevelFlg = gvRow.Cells[15].Text;
            string TotalAssetLevelFlg = gvRow.Cells[16].Text;
            string AumExceptionID = gvRow.Cells[17].Text;


            string ACLevelBillingExceptionFlg = gvRow.Cells[13].Text;
            string ACLevelAUMExceptionFlg = gvRow.Cells[14].Text;
            //DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            BillingExceptionID = BillingExceptionID.Replace("&nbsp;", "");
            AumExceptionID = AumExceptionID.Replace("&nbsp;", "");
            string BillingFlag = gvRow.Cells[41].Text;
            string AUMFlag = gvRow.Cells[42].Text;

            TextBox BillingPer = (TextBox)gvRow.Cells[4].FindControl("txtBillPer");   //4
            string vBillingPer = BillingPer.Text;

            TextBox Billingtxt = (TextBox)gvRow.Cells[4].FindControl("txtBilling");   //4
            string vBilling = Billingtxt.Text;


            TextBox AUMtxt = (TextBox)gvRow.Cells[4].FindControl("txtAUM");   //4
            string vAUM = AUMtxt.Text;

            string SubBilling = gvRow.Cells[31].Text;
            SubBilling = SubBilling.Replace("&nbsp", "");

            string SubAum = gvRow.Cells[32].Text;
            SubAum = SubAum.Replace("&nbsp", "");

            DateTime startDate;
            string ssi_Startdate = gvRow.Cells[33].Text;
            //int ind = ssi_Startdate.IndexOf(" ");
            //ssi_Startdate = ssi_Startdate.Substring(0, ind);
            if (ssi_Startdate == "1/1/1900")
                startDate = DateTime.ParseExact(ssi_Startdate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            else
            {
                startDate = DateTime.ParseExact(ssi_Startdate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            }

            string ssi_EndDate = gvRow.Cells[34].Text;
            //ind = ssi_EndDate.IndexOf(" ");
            //ssi_EndDate = ssi_EndDate.Substring(0, ind);
            DateTime EndDate = DateTime.ParseExact(ssi_EndDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            string MinDate = gvRow.Cells[40].Text;
            MinDate = MinDate.Replace("&nbsp;", "");
            if (MinDate != "")
            {
                //ind = MinDate.IndexOf(" ");
                //MinDate = MinDate.Substring(0, ind);
                DateTime dMinDate = DateTime.ParseExact(MinDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            }

            CheckBox ExcludeBillingcb = (CheckBox)gvRow.Cells[5].FindControl("cbExcludeBilling");   //5
            bool ExcludeBilling = ExcludeBillingcb.Checked;

            string vBillingExcludeFlg = gvRow.Cells[22].Text;  //22   BillingExcludeFlg

            string AsOFDate = txtAUMDate.Text;
            if (AsOFDate == "")
                AsOFDate = "03/31/2016";

            DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            if (BillingFlag == "1")   //Billing Removed all exceptions   Assetlevel _+veAmount 
            {
                if (BillingExceptionID != "" && AssetLevelFlg != "True")   //delete
                {

                    UpdatingDelete(gvRow);

                }
                else if (BillingExceptionID != "" && AssetLevelFlg == "True")   //delete
                {
                    if (startDate < dAsOFDate)
                    {
                        UpdatingDelete(gvRow);
                        insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vBilling), 0);
                    }
                    else if (startDate == dAsOFDate)
                    {
                        if (vBillingExcludeFlg == "True")
                        {
                            if (!ExcludeBilling)
                            {
                                UpdatingDelete(gvRow);
                            }
                            else
                            {
                                updateBill(gvRow, ExcludeBilling, 0, Convert.ToDecimal(vBilling), Convert.ToDecimal(vBillingPer), "Billing");
                            }
                        }
                        else
                        {
                            updateBill(gvRow, ExcludeBilling, 0, Convert.ToDecimal(vBilling), Convert.ToDecimal(vBillingPer), "Billing");
                        }
                    }

                }
                else if (BillingExceptionID == "" && AssetLevelFlg == "True")
                {
                    insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vBilling), 0);

                }
                gvRow.Cells[41].Text = "";

            }
            else if (BillingFlag == "2")      //Assetlevel _-veAmount
            {
                if (BillingExceptionID != "" && AssetLevelFlg == "True")   //delete
                {
                    if (startDate < dAsOFDate)
                    {
                        UpdatingDelete(gvRow);
                        insertCRM(gvRow, false, false, Convert.ToDecimal(SubBilling), 0, 0, 0);
                    }
                    else if (startDate == dAsOFDate)
                    {
                        if (vBillingExcludeFlg == "True")
                        {
                            if (!ExcludeBilling)
                            {
                                UpdatingDelete(gvRow);
                            }
                            else
                            {
                                updateBill(gvRow, ExcludeBilling, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                        }
                        else
                        {
                            updateBill(gvRow, ExcludeBilling, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");

                        }
                    }
                    else if (BillingExceptionID == "" && AssetLevelFlg == "True")
                    {
                        insertCRM(gvRow, false, false, Convert.ToDecimal(SubBilling), 0, 0, 0);
                    }
                    gvRow.Cells[41].Text = "";
                }
            }
            else if (BillingFlag == "3")  // Asset level Exculde
            {
                if (BillingExceptionID != "" && AssetLevelFlg != "True")   //delete
                {
                    UpdatingDelete(gvRow);

                }
                else if (BillingExceptionID != "" && AssetLevelFlg == "True")   //delete
                {
                    if (startDate < dAsOFDate)
                    {
                        UpdatingDelete(gvRow);

                        insertCRM(gvRow, true, false, 0, 0, 0, 0);
                    }
                    else if (startDate == dAsOFDate)
                    {
                        if (vBillingExcludeFlg == "True")
                        {
                            if (!ExcludeBilling)
                            {
                                UpdatingDelete(gvRow);
                            }
                            else
                            {

                                updateBill(gvRow, true, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                        }
                        else
                        {
                            updateBill(gvRow, true, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                        }
                    }
                }
                else if (BillingExceptionID == "" && AssetLevelFlg == "True")
                {
                    insertCRM(gvRow, true, false, 0, 0, 0, 0);

                }

                gvRow.Cells[41].Text = "";
            }
            else if (BillingFlag == "" && BillingExceptionID != "" && AssetLevelFlg == "True")
            {
                if (gvRow.Cells[19].Text == "1")
                {
                    CheckBox ExcludeBillingcb1 = (CheckBox)gvRow.Cells[5].FindControl("cbExcludeBilling");   //5
                    bool ExcludeBilling1 = ExcludeBillingcb1.Checked;

                    // string vBillingExcludeFlg = gvRow.Cells[22].Text;  //22   BillingExcludeFlg
                    if (Convert.ToBoolean(vBillingExcludeFlg) != ExcludeBilling1)
                    {
                        UpdatingDelete(gvRow);
                    }
                }
            }

            else
            {
                if (checkChangesForInsert(gvRow))
                {
                    string vBillingMarketValue = gvRow.Cells[3].Text;
                    //  string vBillingExcludeFlg = gvRow.Cells[22].Text;  //22   BillingExcludeFlg
                    if (AssetLevelFlg != "True" && TotalAssetLevelFlg != "True")
                    {
                        if (BillingExceptionID == "")
                            insertRecord(gvRow);

                        else if (TotalAssetLevelFlg != "True" && BillingExceptionID == "")
                        { insertRecord(gvRow); }

                        else if ((SubBilling != "" || vBilling != vBillingMarketValue || Convert.ToDecimal(vBillingPer) != 0))  //&& ExcludeBilling !=Convert.ToBoolean( vBillingExcludeFlg)
                        {
                            if (!ExcludeBilling)
                                if (SubBilling != "")
                                {
                                    updateBill(gvRow, ExcludeBilling, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                                }
                                else if (vBilling != vBillingMarketValue)
                                {
                                    updateBill(gvRow, ExcludeBilling, 0, Convert.ToDecimal(vBilling), Convert.ToDecimal(vBillingPer), "Billing");
                                }
                                else
                                {
                                    updateBill(gvRow, ExcludeBilling, Convert.ToDecimal(SubBilling), Convert.ToDecimal(vBilling), Convert.ToDecimal(vBillingPer), "Billing");
                                }

                            if (vBillingExcludeFlg == "True")
                            {
                                if (!ExcludeBilling)
                                {
                                    UpdatingDelete(gvRow);
                                }

                            }
                        }
                        else
                        {


                            if (vBillingExcludeFlg == "True")
                            {
                                if (!ExcludeBilling)
                                {
                                    UpdatingDelete(gvRow);
                                }

                            }
                        }
                    }
                    else if (TotalAssetLevelFlg != "True")
                    {
                        if (BillingExceptionID == "")
                        {
                            if (BillingFlag != "")
                                insertRecord(gvRow);
                        }
                        //-- BillingFlag
                    }
                }
            }
        }
    }
    public void Save()
    {
        foreach (GridViewRow gvRow in gvBilling.Rows)
        {
            string BillingExceptionID = gvRow.Cells[17].Text;
            string AssetLevelFlg = gvRow.Cells[15].Text;
            string TotalAssetLevelFlg = gvRow.Cells[16].Text;
            string AumExceptionID = gvRow.Cells[17].Text;

            string ssi_Startdate = gvRow.Cells[33].Text;
            string ssi_EndDate = gvRow.Cells[34].Text;

            string ACLevelBillingExceptionFlg = gvRow.Cells[13].Text;
            string ACLevelAUMExceptionFlg = gvRow.Cells[14].Text;

            //int ind = ssi_Startdate.IndexOf(" ");
            //ssi_Startdate = ssi_Startdate.Substring(0, ind);

            //DateTime startDate = DateTime.ParseExact(ssi_Startdate, "M/dd/yyyy", CultureInfo.InvariantCulture);

            //ind = ssi_EndDate.IndexOf(" ");
            //ssi_EndDate = ssi_EndDate.Substring(0, ind);
            //DateTime EndDate = DateTime.ParseExact(ssi_EndDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            string AsOFDate = txtAUMDate.Text;
            if (AsOFDate == "")
                AsOFDate = "03/31/2016";

            DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            BillingExceptionID = BillingExceptionID.Replace("&nbsp;", "");
            AumExceptionID = AumExceptionID.Replace("&nbsp;", "");
            string Flag = gvRow.Cells[41].Text;

            if (BillingExceptionID == "" && AumExceptionID == "")
            {
                if (ACLevelBillingExceptionFlg != "True" && ACLevelAUMExceptionFlg != "True" && AssetLevelFlg != "True" && TotalAssetLevelFlg != "True")
                {
                    if (checkChangesForInsert(gvRow))
                    {
                        insertRecord(gvRow);
                    }
                    // insertRecord(gvRow);
                }
                else if ((ACLevelBillingExceptionFlg == "True" || ACLevelAUMExceptionFlg == "True") && AssetLevelFlg == "True" && TotalAssetLevelFlg != "True")
                {
                    insertRecord(gvRow);
                }
            }

            else if (BillingExceptionID != "" && AssetLevelFlg != "True" && TotalAssetLevelFlg != "True")
            {
                if (checkUpdatesForBilling(gvRow) == "Delete")
                {
                    UpdateRecordBilling(gvRow, BillingExceptionID, true);
                }

            }

            //else if(  )
            //{

            //}

        }
    }

    #endregion

    public void UpdateRecordBilling(GridViewRow row, string BillingExceptionID, bool bIsDelete)
    {

        bool bISBilling = false;
        bool bProceed = true;
        TextBox Billingtxt = (TextBox)row.FindControl("txtBilling");   //4
        string vBilling = Billingtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        CheckBox ExcludeBillingcb = (CheckBox)row.FindControl("cbExcludeBilling");   //5
        bool ExcludeBilling = ExcludeBillingcb.Checked;

        TextBox Aumtxt = (TextBox)row.FindControl("txtAUM");   //6
        string vAum = Aumtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        CheckBox ExcludeAUMcb = (CheckBox)row.FindControl("cbExcludeAum");   //7
        bool ExcludeAum = ExcludeAUMcb.Checked;

        TextBox BillPertxt = (TextBox)row.FindControl("txtBillPer");   //8
        string vBillPer = BillPertxt.Text;

        string vActual = row.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
        string vAccountID = row.Cells[10].Text;   //9  AccountID
        string vSecurityID = row.Cells[11].Text;   //10 SecurityID
        string vAssetClassID = row.Cells[12].Text;  //11 "Asset Class ID"
        string vBillingID = row.Cells[13].Text;   //12   Billing ID"
        string vAccountBillingExceptionFlg = row.Cells[14].Text;     //13   ACLevelBillingExceptionFlg
        string vAccountAumExceptionFlg = row.Cells[15].Text;  //14  ACLevelAUMExceptionFlg 
        string vAssetExceptionFlg = row.Cells[16].Text;  //15   AssetLevelFlg

        string vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 
        string vBillingFeeExceptionID = row.Cells[19].Text;  //18   BillingFeeExceptionId
        string vBillingExceptionType = row.Cells[20].Text;  //19   BillingExceptionType

        string vAUMExceptionType = row.Cells[21].Text;  //20  AUMExceptionType
        string vIdNmb = row.Cells[22].Text;  //21  IdNmb
        string vBillingExcludeFlg = row.Cells[23].Text;  //22   BillingExcludeFlg


        string vBillingFeePct = row.Cells[27].Text;  //26   BillingFeePct
        string AsofDate = txtAUMDate.Text;// 03/31/2016
        if (AsofDate == "")
            AsofDate = "03/31/2016";

        string vBillingMarketValue = row.Cells[29].Text;  //28  FinalBillingMarketValue 


        string SubtractBilling = row.Cells[32].Text;
        SubtractBilling = SubtractBilling.Replace("&nbsp", "");

        string ssi_Startdate = row.Cells[34].Text;
        string ssi_EndDate = row.Cells[35].Text;
        string MinDate = row.Cells[41].Text;

        //int ind = ssi_Startdate.IndexOf(" ");
        //ssi_Startdate = ssi_Startdate.Substring(0, ind);
        DateTime startDate = DateTime.ParseExact(ssi_Startdate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

        //ind = ssi_EndDate.IndexOf(" ");
        //ssi_EndDate = ssi_EndDate.Substring(0, ind);
        DateTime EndDate = DateTime.ParseExact(ssi_EndDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

        //ind = MinDate.IndexOf(" ");
        //MinDate = MinDate.Substring(0, ind);
        DateTime dMinDate = DateTime.ParseExact(MinDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

        string AsOFDate = txtAUMDate.Text;
        if (AsOFDate == "")
            AsOFDate = "03/31/2016";
        DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

        #region CRM Connection
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        //  CrmService service = null;
        IOrganizationService service = null;


        try
        {
            string UserId = GetcurrentUser();

            service = GM.GetCrmService();// GetCrmService(crmServerUrl, orgName, UserId);
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        #endregion

        if (bProceed)
        {
            if (dAsOFDate == startDate)
            {
                // update  only
                // ssi_billingexception objBillingException = new ssi_billingexception();

                Entity objBillingException = new Entity("ssi_billingexception");

                //objBillingException.ssi_billingexceptionid = new Key();
                //objBillingException.ssi_billingexceptionid.Value = new Guid(Convert.ToString(vBillingExceptionID));
                objBillingException["ssi_billingexceptionid"] = new Guid(Convert.ToString(vBillingExceptionID));

                if (bIsDelete != true)
                {
                    //objBillingException.ssi_excludeasset = new CrmBoolean();
                    //objBillingException.ssi_excludeasset.Value = ExcludeBilling;     // excude billing
                    objBillingException["ssi_excludeasset"] = Convert.ToBoolean(ExcludeBilling);
                    if (SubtractBilling != "0")
                    {
                        //objBillingException.ssi_subtractfromasset = new CrmMoney();
                        //objBillingException.ssi_subtractfromasset.Value = Convert.ToDecimal(SubtractBilling);
                        objBillingException["ssi_subtractfromasset"] = new Money(Convert.ToDecimal(SubtractBilling));
                    }
                    else if (vBilling != "0")
                    {
                        //objBillingException.ssi_billingassetamount = new CrmMoney();
                        //objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vBilling);
                        objBillingException["ssi_billingassetamount"] = new Money(Convert.ToDecimal(vBilling));
                    }

                    if (vBillPer != "0")
                    {
                        //objBillingException.ssi_feepercent = new CrmDecimal();
                        //objBillingException.ssi_feepercent.Value = Convert.ToDecimal(vBillPer);
                        objBillingException["ssi_feepercent"] = Convert.ToDecimal(vBillPer);
                    }
                }
                else
                {
                    //objBillingException.ssi_enddate = new CrmDateTime();
                    //objBillingException.ssi_enddate.Value = dAsOFDate.AddDays(-1).ToString();
                    objBillingException["ssi_enddate"] = dAsOFDate.AddDays(-1);

                }

                //  Service.Update(objBillingException);

                service.Update(objBillingException);



            }

            else if (startDate > dAsOFDate)
            {
                // update with -1 
                // insert new record



            }
            else if (dAsOFDate > startDate)
            {
                // sd=asofdate
                // enddate =-1
            }

        }
    }

    public void updateBill(GridViewRow row, bool bExcludeBilling, decimal dSubtractBilling, decimal dPositiveBilling, decimal dBillPer, String Type, bool IsBillPerOnly = false)
    {
        bool bISBilling = false;

        bool bProceed = true;
        string vBillingExceptionID = "";
        bool Exclude = false;
        string ssiAum = string.Empty;
        string vddlAdminCategory = string.Empty;
        string vddlNonGAServiceType = string.Empty;

        //if (Type != "AUM")
        //    vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 
        //else
        //    vBillingExceptionID = row.Cells[36].Text;
        if (Type == "Billing")
        {
            vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 
            CheckBox cbBilling = (CheckBox)row.FindControl("cbExcludeBilling");   //7
            Exclude = cbBilling.Checked;
            ssiAum = "100000000";
        }
        else if (Type == "AUM")
        {
            vBillingExceptionID = row.Cells[36].Text;
            CheckBox cbAUM = (CheckBox)row.FindControl("cbExcludeAUM");   //7
            Exclude = cbAUM.Checked;
            ssiAum = "100000001";
        }
        else if (Type == "NONGA")
        {
            vBillingExceptionID = row.Cells[60].Text;
            DropDownList ddlAdminCategory = (DropDownList)row.FindControl("ddlAdminCategory");
            vddlAdminCategory = ddlAdminCategory.SelectedValue;

            DropDownList ddlNonGAServiceType = (DropDownList)row.FindControl("ddlNonGAServiceType");
            vddlNonGAServiceType = ddlNonGAServiceType.SelectedValue;

        }
        string AsOFDate = txtAUMDate.Text;
        if (AsOFDate == "")
            AsOFDate = "03/31/2016";

        DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

        #region CRM Connection
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;

        try
        {
            string UserId = GetcurrentUser();

            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        #endregion

        if (bProceed)
        {
            // update  only
            //   ssi_billingexception objBillingException = new ssi_billingexception();
            Entity objBillingException = new Entity("ssi_billingexception");
            //objBillingException.ssi_billingexceptionid = new Key();
            //objBillingException.ssi_billingexceptionid.Value = new Guid(Convert.ToString(vBillingExceptionID));
            objBillingException["ssi_billingexceptionid"] = new Guid(Convert.ToString(vBillingExceptionID));

            if (!IsBillPerOnly)
            {

                if (Type == "Billing" || Type == "AUM")
                {
                    //objBillingException.ssi_excludeasset = new CrmBoolean();
                    //objBillingException.ssi_excludeasset.Value = bExcludeBilling;     // excude billing
                    objBillingException["ssi_excludeasset"] = Exclude;

                    objBillingException["ssi_aum"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ssiAum));



                }
                if (Type == "NONGA")
                {
                    if (vddlNonGAServiceType != "")
                    {
                        objBillingException["ssi_nongaservicetype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vddlNonGAServiceType));
                    }
                    else
                    {
                        objBillingException["ssi_nongaservicetype"] = null;
                    }
                    if (vddlAdminCategory != "")
                    {
                        objBillingException["ssi_admincategory"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vddlAdminCategory));
                    }
                    else
                    {
                        objBillingException["ssi_admincategory"] = null;
                    }
                    if (vddlNonGAServiceType == "" && vddlAdminCategory == "")
                    {
                        UpdatingDelete(row, "NONGA");
                    }
                }
                //if (Type == "Billing")
                //{
                //    //objBillingException.ssi_excludeasset = new CrmBoolean();
                //    //objBillingException.ssi_excludeasset.Value = bExcludeBilling;     // excude billing
                //    objBillingException["ssi_excludeasset"] = bExcludeBilling;
                //}
                //else
                //{
                //    if (dPositiveBilling != 0 || dSubtractBilling != 0)
                //    {
                //        //objBillingException.ssi_aum = new CrmBoolean();
                //        //objBillingException.ssi_aum.Value = true;
                //        objBillingException["ssi_aum"] = true;
                //    }
                //    else
                //    {
                //        //objBillingException.ssi_aum = new CrmBoolean();
                //        //objBillingException.ssi_aum.Value = bExcludeBilling;
                //        objBillingException["ssi_aum"] = bExcludeBilling;
                //    }
                //}

                //if (Type == "AUM" && bExcludeBilling == true)
                //{
                //    //objBillingException.ssi_excludeasset = new CrmBoolean();
                //    //objBillingException.ssi_excludeasset.Value = bExcludeBilling;
                //    objBillingException["ssi_excludeasset"] = bExcludeBilling;
                //}

                //if ((dPositiveBilling != 0 && Type == "AUM" || dSubtractBilling != 0) && Type != "Billing")
                //{
                //    //objBillingException.ssi_excludeasset = new CrmBoolean();
                //    //objBillingException.ssi_excludeasset.Value = bExcludeBilling;
                //    objBillingException["ssi_excludeasset"] = bExcludeBilling;
                //}

                //if (dSubtractBilling != 0)
                //{

                //objBillingException.ssi_subtractfromasset = new CrmMoney();
                //objBillingException.ssi_subtractfromasset.Value = Math.Abs(dSubtractBilling);
                objBillingException["ssi_subtractfromasset"] = new Money(Math.Abs(dSubtractBilling));
                //}
                //if (dPositiveBilling != 0)
                //{
                //objBillingException.ssi_billingassetamount = new CrmMoney();
                //objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(dPositiveBilling);
                objBillingException["ssi_billingassetamount"] = new Money(Convert.ToDecimal(dPositiveBilling));
                //}
            }


            //objBillingException.ssi_feepercent = new CrmDecimal();
            //objBillingException.ssi_feepercent.Value = Convert.ToDecimal(dBillPer);
            objBillingException["ssi_feepercent"] = Convert.ToDecimal(dBillPer);

            service.Update(objBillingException);


        }

    }

    public void insertRecord(GridViewRow row)
    {

        bool bISBilling = false;

        TextBox Billingtxt = (TextBox)row.FindControl("txtBilling");   //4
        string vBilling = Billingtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        CheckBox ExcludeBillingcb = (CheckBox)row.FindControl("cbExcludeBilling");   //5
        bool ExcludeBilling = ExcludeBillingcb.Checked;

        TextBox Aumtxt = (TextBox)row.FindControl("txtAUM");   //6
        string vAum = Aumtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        CheckBox ExcludeAUMcb = (CheckBox)row.FindControl("cbExcludeAum");   //7
        bool ExcludeAum = ExcludeAUMcb.Checked;

        TextBox BillPertxt = (TextBox)row.FindControl("txtBillPer");   //8
        string vBillPer = BillPertxt.Text;

        string vActual = row.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
        string vAccountID = row.Cells[10].Text;   //9  AccountID
        string vSecurityID = row.Cells[11].Text;   //10 SecurityID
        string vAssetClassID = row.Cells[12].Text;  //11 "Asset Class ID"
        string vBillingID = row.Cells[13].Text;   //12   Billing ID"
        string vAccountBillingExceptionFlg = row.Cells[14].Text;     //13   ACLevelBillingExceptionFlg
        string vAccountAumExceptionFlg = row.Cells[15].Text;  //14  ACLevelAUMExceptionFlg 
        string vAssetExceptionFlg = row.Cells[16].Text;  //15   AssetLevelFlg

        string vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 
        string vBillingFeeExceptionID = row.Cells[19].Text;  //18   BillingFeeExceptionId
        string vBillingExceptionType = row.Cells[20].Text;  //19   BillingExceptionType

        string vAUMExceptionType = row.Cells[21].Text;  //20  AUMExceptionType
        string vIdNmb = row.Cells[22].Text;  //21  IdNmb
        string vBillingExcludeFlg = row.Cells[23].Text;  //22   BillingExcludeFlg
        string vAUMExcludeFlg = row.Cells[24].Text;  //23  AUMExcludeFlg   

        string vBillingFeePct = row.Cells[25].Text;  //26   BillingFeePct
        string AsofDate = txtAUMDate.Text;// 03/31/2016
        if (AsofDate == "")
            AsofDate = "03/31/2016";

        string vBillingMarketValue = row.Cells[29].Text;  //28  FinalBillingMarketValue 
        string vAUMMarketValue = row.Cells[30].Text;  //29  FinalAUMMarketValue 

        string SubtractBilling = row.Cells[32].Text;
        SubtractBilling = SubtractBilling.Replace("&nbsp", "");

        string SubtractAum = row.Cells[33].Text;
        SubtractAum = SubtractAum.Replace("&nbsp", "");

        string BillingExceptionID = row.Cells[18].Text;

        if (ExcludeBilling && !ExcludeAum)
        {
            insertCRM(row, true, false, 0, 0, 0, 0);

            if (SubtractAum != "")
                insertCRM(row, false, true, 0, Convert.ToDecimal(SubtractAum), 0, 0);
            else if (vActual != vAum)
                insertCRM(row, false, true, 0, 0, 0, Convert.ToDecimal(vAum));
        }
        else if (!ExcludeBilling && ExcludeAum)
        {
            insertCRM(row, false, true, 0, 0, 0, 0);

            if (SubtractBilling != "")
                insertCRM(row, false, false, 0, Convert.ToDecimal(SubtractBilling), 0, 0);
            else if (vBilling != vActual)
                insertCRM(row, false, false, 0, 0, Convert.ToDecimal(vBilling), 0);

        }
        else if (ExcludeBilling && ExcludeAum)
        {
            insertCRM(row, false, true, 0, 0, 0, 0);
            insertCRM(row, true, false, 0, 0, 0, 0);
        }
        else if (!ExcludeBilling && !ExcludeAum)
        {

            if (SubtractBilling != "")
                insertCRM(row, false, false, 0, Convert.ToDecimal(SubtractBilling), 0, 0);
            else if (vBilling != vActual)
                insertCRM(row, false, false, 0, 0, Convert.ToDecimal(vBilling), 0);

            if (SubtractAum != "")
                insertCRM(row, false, true, 0, Convert.ToDecimal(SubtractAum), 0, 0);
            else if (vActual != vAum)
                insertCRM(row, false, true, 0, 0, 0, Convert.ToDecimal(vAum));
        }
        //else
        //{
        //    insertCRM(row, ExcludeBilling, ExcludeAum);
        //}
        ////  ssi_name
        //  objBillingException.ssi_name= new Lookup();

        //  //ownerid
        //  objBillingException.ownerid= new Owner();

    }

    public void insertCRM(GridViewRow row, bool bExcludeBilling, bool bExcludeAum, decimal SubtractBilling, decimal SubtractAum, decimal vBilling, decimal vAum, decimal Fees = 0, string type = null, string CloseDate = null)
    {

        #region CRM Connection
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;

        try
        {
            string UserId = GetcurrentUser();

            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            //  bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            //bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        #endregion

        bool bProceed = true;
        bool bISBilling = false;
        string Actual = row.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        TextBox Billingtxt = (TextBox)row.FindControl("txtBilling");   //4
        string vBilltext = Billingtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        CheckBox ExcludeBillingcb = (CheckBox)row.FindControl("cbExcludeBilling");   //5
        bool ExcludeBilling = ExcludeBillingcb.Checked;

        TextBox Aumtxt = (TextBox)row.FindControl("txtAUM");   //6
        string vAumtext = Aumtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        CheckBox ExcludeAUMcb = (CheckBox)row.FindControl("cbExcludeAum");   //7
        bool ExcludeAum = ExcludeAUMcb.Checked;

        TextBox BillPertxt = (TextBox)row.FindControl("txtBillPer");   //8
        string vBillPer = BillPertxt.Text;
        if (vBillPer == "")
        { vBillPer = "0"; }

        string vAccountID = row.Cells[10].Text;   //9  AccountID
        string vSecurityID = row.Cells[11].Text;   //10 SecurityID
        string vAssetClassID = row.Cells[12].Text;  //11 "Asset Class ID"
        string vBillingID = row.Cells[13].Text;   //12   Billing ID"
        string vAccountBillingExceptionFlg = row.Cells[14].Text;     //13   ACLevelBillingExceptionFlg
        string vAccountAumExceptionFlg = row.Cells[15].Text;  //14  ACLevelAUMExceptionFlg 
        string vAssetExceptionFlg = row.Cells[16].Text;  //15   AssetLevelFlg

        string vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 
        string vBillingFeeExceptionID = row.Cells[19].Text;  //18   BillingFeeExceptionId
        string vBillingExceptionType = row.Cells[20].Text;  //19   BillingExceptionType

        string vAUMExceptionType = row.Cells[21].Text;  //20  AUMExceptionType
        string vIdNmb = row.Cells[22].Text;  //21  IdNmb
        string vBillingExcludeFlg = row.Cells[23].Text;  //22   BillingExcludeFlg
        string vAUMExcludeFlg = row.Cells[24].Text;  //23  AUMExcludeFlg   

        string vBillingFeePct = row.Cells[27].Text;  //26   BillingFeePct
        string AsofDate = txtAUMDate.Text;// 03/31/2016


        //string MinAumDate = row.Cells[40].Text;
        //string MinBillingDate = row.Cells[41].Text;


        //if (AsofDate == "")
        //    AsofDate = "03/31/2016";

        string vBillingMarketValue = row.Cells[29].Text;  //28  FinalBillingMarketValue 
        string vAUMMarketValue = row.Cells[30].Text;  //29  FinalAUMMarketValue 

        string SubBilling = row.Cells[32].Text;
        SubBilling = SubBilling.Replace("&nbsp", "");

        string SubAum = row.Cells[33].Text;
        SubAum = SubAum.Replace("&nbsp", "");

        vSecurityID = vSecurityID.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
        vAssetClassID = vAssetClassID.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");

        string BillingExceptionID = row.Cells[18].Text;
        string ACLevelBillingExceptionFlg = row.Cells[14].Text;
        string ACLevelAUMExceptionFlg = row.Cells[15].Text;
        string ACLevelFlgChng = row.Cells[44].Text;
        //ssi_billingexception objBillingException = new ssi_billingexception();
        Entity objBillingException = new Entity("ssi_billingexception");
        //objBillingException.ssi_billingexceptionid = new Key();
        //objBillingException.ssi_billingexceptionid.Value = Guid.NewGuid();

        //if ((ACLevelBillingExceptionFlg != "True" || ACLevelAUMExceptionFlg != "True") && vAssetExceptionFlg != "True")
        //{

        //    //objBillingException.ssi_account = new Lookup();
        //    //objBillingException.ssi_account.type = EntityName.ssi_account.ToString();
        //    //objBillingException.ssi_account.Value = new Guid(vAccountID);
        //    objBillingException["ssi_account"] = new EntityReference("ssi_account", new Guid(vAccountID));

        //    //ssi_security

        //    if (vSecurityID != "")
        //    {
        //        //objBillingException.ssi_security = new Lookup();
        //        //objBillingException.ssi_security.type = EntityName.ssi_security.ToString();
        //        //objBillingException.ssi_security.Value = new Guid(vSecurityID);
        //        objBillingException["ssi_security"] = new EntityReference("ssi_security", new Guid(vSecurityID));
        //    }
        //    //else
        //    //{
        //    //    objBillingException.ssi_security = new Lookup();
        //    //    objBillingException.ssi_security.IsNull = true;
        //    //    objBillingException.ssi_security.IsNullSpecified = true;
        //    //    objBillingException["ssi_security"] = null;
        //    //}
        //}
        if (type == "Billing") //--  changed 19_2_2020 
        {
            if ((ACLevelBillingExceptionFlg != "True" || ACLevelFlgChng == "" ) && vAssetExceptionFlg != "True")
            {
                if (vAccountID != "")
                {
                    objBillingException["ssi_account"] = new EntityReference("ssi_account", new Guid(vAccountID));
                }
                if (vSecurityID != "")
                {
                    //objBillingException.ssi_security = new Lookup();
                    //objBillingException.ssi_security.type = EntityName.ssi_security.ToString();
                    //objBillingException.ssi_security.Value = new Guid(vSecurityID);
                    objBillingException["ssi_security"] = new EntityReference("ssi_security", new Guid(vSecurityID));
                }
            }

        }
        else if (type == "AUM")
        {
            if ((ACLevelAUMExceptionFlg != "True" || ACLevelFlgChng == "")&& vAssetExceptionFlg != "True")
            {
                if (vAccountID != "")
                {
                    objBillingException["ssi_account"] = new EntityReference("ssi_account", new Guid(vAccountID));
                }
                if (vSecurityID != "")
                {
                    //objBillingException.ssi_security = new Lookup();
                    //objBillingException.ssi_security.type = EntityName.ssi_security.ToString();
                    //objBillingException.ssi_security.Value = new Guid(vSecurityID);
                    objBillingException["ssi_security"] = new EntityReference("ssi_security", new Guid(vSecurityID));
                }
            }

        }
        else if (type == "NONGA")
        {
            if (vAccountID != "")
            {
                objBillingException["ssi_account"] = new EntityReference("ssi_account", new Guid(vAccountID));
            }
            if (vSecurityID != "")
            {
                //objBillingException.ssi_security = new Lookup();
                //objBillingException.ssi_security.type = EntityName.ssi_security.ToString();
                //objBillingException.ssi_security.Value = new Guid(vSecurityID);
                objBillingException["ssi_security"] = new EntityReference("ssi_security", new Guid(vSecurityID));
            }
        }
        //else
        //{

        //    //ssi_assetclass
        //    objBillingException.ssi_assetclass = new Lookup();
        //    objBillingException.ssi_assetclass.type = EntityName.sas_assetclass.ToString();
        //    objBillingException.ssi_assetclass.Value = new Guid(vAssetClassID);

        //}

        if (vAssetClassID != "")
        {
            //objBillingException.ssi_assetclass = new Lookup();
            //objBillingException.ssi_assetclass.type = EntityName.sas_assetclass.ToString();
            //objBillingException.ssi_assetclass.Value = new Guid(vAssetClassID);
            objBillingException["ssi_assetclass"] = new EntityReference("sas_assetclass", new Guid(vAssetClassID));
        }
        //else
        //{
        //    objBillingException.ssi_assetclass = new Lookup();
        //    objBillingException.ssi_assetclass.IsNull = true;
        //    objBillingException.ssi_assetclass.IsNullSpecified = true;
        //}


        //vBillingID
        //objBillingException.ssi_billingfor = new Lookup();
        //objBillingException.ssi_billingfor.type = EntityName.ssi_billing.ToString();
        //objBillingException.ssi_billingfor.Value = new Guid(vBillingID);
        objBillingException["ssi_billingfor"] = new EntityReference("ssi_billing", new Guid(vBillingID));


        // ssi_aum
        //objBillingException.ssi_aum = new CrmBoolean();    // excude AUM 
        //objBillingException.ssi_aum.Value = false;

        ////  ssi_excludeasset
        //if (bExcludeBilling)
        //{
        //    //objBillingException.ssi_excludeasset = new CrmBoolean();
        //    //objBillingException.ssi_excludeasset.Value = true;     // excude billing
        //    objBillingException["ssi_excludeasset"] = true;

        //    //objBillingException.ssi_aum = new CrmBoolean();
        //    //objBillingException.ssi_aum.Value = false;      //    excude AUM
        //}
        //if (bExcludeAum)
        //{
        //    //objBillingException.ssi_aum = new CrmBoolean();
        //    //objBillingException.ssi_aum.Value = true;      //    excude AUM
        //    objBillingException["ssi_aum"] = true;
        //}

        if (type == "Billing")
        {
            CheckBox ExcludeBillingcb1 = (CheckBox)row.FindControl("cbExcludeBilling");   //5
            bool ExcludeBilling1 = ExcludeBillingcb.Checked;
            objBillingException["ssi_excludeasset"] = ExcludeBilling1;

            objBillingException["ssi_aum"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000000);
        }
        else if (type == "AUM")
        {

            CheckBox ExcludeAUMcb1 = (CheckBox)row.FindControl("cbExcludeAum");   //7
            bool ExcludeAum1 = ExcludeAUMcb.Checked;
            objBillingException["ssi_excludeasset"] = ExcludeAum1;

            objBillingException["ssi_aum"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000001);
        }
        else if (type == "NONGA")
        {
            DropDownList ddlAdminCategory = (DropDownList)row.FindControl("ddlAdminCategory");
            string vddlAdminCategory = ddlAdminCategory.SelectedValue;
            if (vddlAdminCategory != "")
            {
                objBillingException["ssi_admincategory"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vddlAdminCategory));
            }


            DropDownList ddlNonGAServiceType = (DropDownList)row.FindControl("ddlNonGAServiceType");
            string vddlNonGAServiceType = ddlNonGAServiceType.SelectedValue;
            if (vddlNonGAServiceType != "")
            {
                objBillingException["ssi_nongaservicetype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vddlNonGAServiceType));
            }

            if (vddlNonGAServiceType == "" && vddlAdminCategory == "")
            {
                bProceed = false;
            }
            objBillingException["ssi_aum"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000002);

        }






        if (SubtractBilling != 0)
        {
            //objBillingException.ssi_subtractfromasset = new CrmMoney();
            //objBillingException.ssi_subtractfromasset.Value = Math.Abs(SubtractBilling);
            objBillingException["ssi_subtractfromasset"] = new Money(Math.Abs(SubtractBilling));
        }
        if (vBilling != 0)
        {
            //objBillingException.ssi_billingassetamount = new CrmMoney();
            //objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vBilling);
            objBillingException["ssi_billingassetamount"] = new Money(Convert.ToDecimal(vBilling));
        }

        if (SubtractAum != 0)
        {
            //objBillingException.ssi_subtractfromasset = new CrmMoney();
            //objBillingException.ssi_subtractfromasset.Value = Math.Abs(SubtractAum);

            objBillingException["ssi_subtractfromasset"] = new Money(Math.Abs(SubtractAum));

        }
        if (vAum != 0)
        {
            //objBillingException.ssi_billingassetamount = new CrmMoney();
            //objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vAum);
            objBillingException["ssi_billingassetamount"] = new Money(Convert.ToDecimal(vAum));
        }

        #region Notinused
        // if (bExcludeAum && !bExcludeBilling)
        // {
        // if (SubtractAum != 0)
        // {
        // objBillingException.ssi_subtractfromasset = new CrmMoney();
        // objBillingException.ssi_subtractfromasset.Value = Convert.ToDecimal(SubtractAum);
        // }
        // else
        // {
        // if(vAum)
        // objBillingException.ssi_billingassetamount = new CrmMoney();
        // objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vAum);
        // }
        // }


        // if (bExcludeAum && !bExcludeBilling)
        // {
        // if (SubtractBilling != 0)
        // {
        // objBillingException.ssi_subtractfromasset = new CrmMoney();
        // objBillingException.ssi_subtractfromasset.Value = Convert.ToDecimal(SubtractBilling);
        // }

        // else
        // {
        // objBillingException.ssi_billingassetamount = new CrmMoney();
        // objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vBilling);
        // }
        // }

        // if (!bExcludeBilling)
        // {
        // if (SubtractBilling != 0)
        // {
        // objbillingexception.ssi_subtractfromasset = new crmmoney();
        // objBillingException.ssi_subtractfromasset.Value = Convert.ToDecimal(SubtractBilling);
        // }
        // else
        // {
        // objBillingException.ssi_billingassetamount = new CrmMoney();
        // objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vBilling);
        // }
        // }
        // if (!bExcludeAum)
        // {
        // if (SubtractAum != 0)
        // {
        // objbillingexception.ssi_subtractfromasset = new crmmoney();
        // objBillingException.ssi_subtractfromasset.Value = Convert.ToDecimal(SubtractBilling);
        // }
        // else
        // {
        // objBillingException.ssi_billingassetamount = new CrmMoney();
        // objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vAum);
        // }
        // }

        // if (!bExcludeAum && !bExcludeBilling)
        // {
        // if (SubtractBilling != 0 && SubtractAum == 0 )
        // {
        // objBillingException.ssi_subtractfromasset = new CrmMoney();
        // objBillingException.ssi_subtractfromasset.Value = Convert.ToDecimal(SubtractBilling);
        // }
        // else  //if (vBilling != Actual && SubtractAum == 0 && SubtractBilling != 0)
        // {
        // objBillingException.ssi_billingassetamount = new CrmMoney();
        // objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vBilling);

        // }

        // if (SubtractBilling == 0 && SubtractAum != 0)
        // {
        // objBillingException.ssi_subtractfromasset = new CrmMoney();
        // objBillingException.ssi_subtractfromasset.Value = Convert.ToDecimal(SubtractAum);
        // }
        // else //if(vAum != Actual && SubtractAum == 0 && SubtractBilling == 0)
        // {
        // objBillingException.ssi_billingassetamount = new CrmMoney();
        // objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vAum);
        // }

        // }
        #endregion

        //  ssi_startdate
        //objBillingException.ssi_startdate = new CrmDateTime();
        //objBillingException.ssi_startdate.Value = AsofDate;// "20160331"; // DateTime.Now.ToString();
        objBillingException["ssi_startdate"] = Convert.ToDateTime(AsofDate);

        //ssi_enddate
        string MinAumDate = row.Cells[40].Text;
        string MinBillingDate = row.Cells[41].Text;

        MinAumDate = MinAumDate.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
        MinBillingDate = MinBillingDate.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");

        string MinFutureDate = row.Cells[47].Text;
        string MinAumFutureDate = row.Cells[48].Text;
        string MinNGAAFutureDate = row.Cells[63].Text;

        MinFutureDate = MinFutureDate.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
        MinAumFutureDate = MinAumFutureDate.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
        MinNGAAFutureDate = MinNGAAFutureDate.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
        if (type == "Billing")
        {
            if (MinFutureDate != "")
            {
                DateTime date = DateTime.ParseExact(MinFutureDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                //objBillingException.ssi_enddate = new CrmDateTime();
                //objBillingException.ssi_enddate.Value = date.AddDays(-1).ToString();
                objBillingException["ssi_enddate"] = date.AddDays(-1);

            }
        }
        else if (type == "AUM")
        {
            if (MinAumFutureDate != "")
            {
                DateTime date = DateTime.ParseExact(MinAumFutureDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                //objBillingException.ssi_enddate = new CrmDateTime();
                //objBillingException.ssi_enddate.Value = date.AddDays(-1).ToString();
                objBillingException["ssi_enddate"] = date.AddDays(-1);
            }
        }
        else if (type == "NONGA")
        {
            if (MinNGAAFutureDate != "")
            {
                DateTime date = DateTime.ParseExact(MinNGAAFutureDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                //objBillingException.ssi_enddate = new CrmDateTime();
                //objBillingException.ssi_enddate.Value = date.AddDays(-1).ToString();
                objBillingException["ssi_enddate"] = date.AddDays(-1);
            }
        }

        if (CloseDate != null)
        {
            DateTime date = DateTime.ParseExact(CloseDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            //objBillingException.ssi_enddate = new CrmDateTime();
            //objBillingException.ssi_enddate.Value = date.AddDays(-1).ToString();
            objBillingException["ssi_enddate"] = date.AddDays(-1);
        }


        //objBillingException.ssi_enddate = new CrmDateTime();
        //objBillingException.ssi_enddate.IsNull = true;
        //objBillingException.ssi_enddate.IsNullSpecified = true;



        // // ssi_billingassetamount
        //objBillingException.ssi_billingassetamount = new CrmMoney();
        //objBillingException.ssi_billingassetamount.Value = Convert.ToDecimal(vBilling);




        //ssi_feepercent
        if (Fees != 0)
        {
            //objBillingException.ssi_feepercent = new CrmDecimal();
            //objBillingException.ssi_feepercent.Value = Convert.ToDecimal(Fees);
            objBillingException["ssi_feepercent"] = Convert.ToDecimal(Fees);
        }
        if (bProceed)
        {
            service.Create(objBillingException);
        }

    }

    public bool checkChangesForInsert(GridViewRow row)
    {



        bool bIsResult = false;

        string TotalAssetLevelFlg = row.Cells[17].Text;

        if (TotalAssetLevelFlg != "True")
        {
            TextBox Billingtxt = (TextBox)row.FindControl("txtBilling");   //4
            string vBilling = Billingtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");


            CheckBox ExcludeBillingcb = (CheckBox)row.FindControl("cbExcludeBilling");   //5
            bool ExcludeBilling = ExcludeBillingcb.Checked;

            TextBox Aumtxt = (TextBox)row.FindControl("txtAUM");   //6
            string vAum = Aumtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            CheckBox ExcludeAUMcb = (CheckBox)row.FindControl("cbExcludeAum");   //7
            bool ExcludeAum = ExcludeAUMcb.Checked;

            TextBox BillPertxt = (TextBox)row.FindControl("txtBillPer");   //8
            string vBillPer = BillPertxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            if (vBillPer == "")
                vBillPer = "0";

            string vBillingExcludeFlg = row.Cells[23].Text;  //22   BillingExcludeFlg
            string vAUMExcludeFlg = row.Cells[24].Text;  //23  AUMExcludeFlg 

            string vBillingMarketValue = row.Cells[29].Text;  //28  FinalBillingMarketValue 
            string vAUMMarketValue = row.Cells[30].Text;  //29  FinalAUMMarketValue 
            string vBillingFeePct = row.Cells[27].Text;  //30  BillingFeePct 

            string positiveValue = row.Cells[44].Text;
            string negetiveValue = row.Cells[32].Text;

            positiveValue = positiveValue.Replace("&nbsp;", "");
            positiveValue = positiveValue.Replace("&nbsp", "");
            positiveValue = positiveValue.Replace(";", "");

            negetiveValue = negetiveValue.Replace("&nbsp;", "");
            negetiveValue = negetiveValue.Replace("&nbsp", "");
            negetiveValue = negetiveValue.Replace(";", "");
            //if (Convert.ToBoolean(vAUMExcludeFlg) != ExcludeAum)
            //    bIsResult = true;

            if (Convert.ToBoolean(vBillingExcludeFlg) != ExcludeBilling)
                bIsResult = true;

            //if (Convert.ToBoolean(vBillingExcludeFlg) != ExcludeBilling)
            //    bIsResult = true;


            if (positiveValue != "")
                bIsResult = true;

            if (negetiveValue != "")
                bIsResult = true;
            //if (Convert.ToDecimal(vBillingMarketValue) != 0)
            //{
            //    if (vBilling != vBillingMarketValue)
            //        bIsResult = true;
            //}

            //if (vAum != vAUMMarketValue)
            //    bIsResult = true;

            if (vBillPer != vBillingFeePct)
                bIsResult = true;

        }
        return bIsResult;

    }
    public bool checkChangesForInsertAum(GridViewRow row)
    {
        bool bIsResult = false;
        string TotalAssetLevelFlg = row.Cells[17].Text;

        if (TotalAssetLevelFlg != "True")
        {
            TextBox Billingtxt = (TextBox)row.FindControl("txtBilling");   //4
            string vBilling = Billingtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            CheckBox ExcludeBillingcb = (CheckBox)row.FindControl("cbExcludeBilling");   //5
            bool ExcludeBilling = ExcludeBillingcb.Checked;

            TextBox Aumtxt = (TextBox)row.FindControl("txtAUM");   //6
            string vAum = Aumtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            CheckBox ExcludeAUMcb = (CheckBox)row.FindControl("cbExcludeAum");   //7
            bool ExcludeAum = ExcludeAUMcb.Checked;

            TextBox BillPertxt = (TextBox)row.FindControl("txtBillPer");   //8
            string vBillPer = BillPertxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            string vBillingExcludeFlg = row.Cells[23].Text;  //22   BillingExcludeFlg
            string vAUMExcludeFlg = row.Cells[24].Text;  //23  AUMExcludeFlg 

            string vBillingMarketValue = row.Cells[29].Text;  //28  FinalBillingMarketValue 
            string vAUMMarketValue = row.Cells[30].Text;  //29  FinalAUMMarketValue 
            string vBillingFeePct = row.Cells[27].Text;  //30  BillingFeePct 

            string positiveValue = row.Cells[46].Text;
            string negetiveValue = row.Cells[33].Text;

            positiveValue = positiveValue.Replace("&nbsp;", "");
            positiveValue = positiveValue.Replace("&nbsp", "");
            positiveValue = positiveValue.Replace(";", "");

            negetiveValue = negetiveValue.Replace("&nbsp;", "");
            negetiveValue = negetiveValue.Replace("&nbsp", "");
            negetiveValue = negetiveValue.Replace(";", "");

            if (Convert.ToBoolean(vAUMExcludeFlg) != ExcludeAum)
                bIsResult = true;

            if (positiveValue != "")
                bIsResult = true;

            if (negetiveValue != "")
                bIsResult = true;
            //if (Convert.ToBoolean(vBillingExcludeFlg) != ExcludeBilling)
            //    bIsResult = true;

            //if (Convert.ToBoolean(vBillingExcludeFlg) != ExcludeBilling)
            //    bIsResult = true;

            //if (Convert.ToDecimal(vAUMMarketValue) != 0)
            //{
            //    if (vAum != vAUMMarketValue)
            //        bIsResult = true;
            //}

            //if (vAum != vAUMMarketValue)
            //    bIsResult = true;

            //if (vBillPer != vBillingFeePct)
            //    bIsResult = true;
        }

        return bIsResult;

    }

    public bool checkChangesForInsertNGAA(GridViewRow row)
    {
        bool bIsResult = false;
        string TotalAssetLevelFlg = row.Cells[17].Text;

        if (TotalAssetLevelFlg != "True")
        {

            DropDownList ddlAdminCategory = (DropDownList)row.FindControl("ddlAdminCategory");
            string vddlAdminCategory = ddlAdminCategory.SelectedValue;
            string gridddlAdminCategory = row.Cells[55].Text.Replace("&nbsp;", ""); ;

            DropDownList ddlNonGAServiceType = (DropDownList)row.FindControl("ddlNonGAServiceType");
            string vddlNonGAServiceType = ddlNonGAServiceType.SelectedValue;
            string gridddlNonGAServiceType = row.Cells[56].Text.Replace("&nbsp;", ""); ;

            if (vddlNonGAServiceType != gridddlNonGAServiceType)
                bIsResult = true;
            if (gridddlAdminCategory != vddlAdminCategory)
                bIsResult = true;

        }

        return bIsResult;

    }
    public string checkUpdatesForBilling(GridViewRow row)
    {
        bool bIsResult = false;
        string status = null;

        TextBox Billingtxt = (TextBox)row.FindControl("txtBilling");   //4
        string vBilling = Billingtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        CheckBox ExcludeBillingcb = (CheckBox)row.FindControl("cbExcludeBilling");   //5
        bool ExcludeBilling = ExcludeBillingcb.Checked;

        TextBox BillPertxt = (TextBox)row.FindControl("txtBillPer");   //8
        string vBillPer = BillPertxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

        string vBillingExcludeFlg = row.Cells[23].Text;  //22   BillingExcludeFlg
        string vBillingMarketValue = row.Cells[25].Text;  //28  FinalBillingMarketValue 

        string vBillingFeePct = row.Cells[27].Text;  //30  BillingFeePct 

        if (Convert.ToBoolean(vBillingExcludeFlg) != ExcludeBilling)
        {
            bIsResult = true;
            status = "Delete";

        }


        vBillingMarketValue = row.Cells[4].Text;

        if (vBilling != vBillingMarketValue)
        {
            bIsResult = true;
            status = "Update";
        }

        if (vBillPer != vBillingFeePct)
        {
            bIsResult = true;
            status = "Update";
        }



        return status;
    }

    #region CRM Connection
    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;

        ////Changed Windows to - ADFS Claims Login 8_9_2019
        //IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        //string strName = claimsIdentity.Name;

        //Changed Windows to - ADFS Claims Login 8_9_2019
        string strName = string.Empty;
        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            strName = "corp\\gbhagia";
        }
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            strName = claimsIdentity.Name;
        }


        //////////
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
            //return UserID = "DFCE21B1-B81E-E211-A2B7-0002A5443D86";
        }
        else
        {
            return UserID = "";
        }
    }

    //public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    // string username = WindowsIdentity.GetCurrent().Name;

    //    if (CallerId != "")
    //        token.CallerId = new Guid(CallerId);

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

    #endregion

    public void UpdatingDelete(GridViewRow row, string Type = null, string Isclose = null)
    {

        bool bProceed = false;

        #region CRM Connection
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;

        try
        {
            string UserId = GetcurrentUser();

            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            bProceed = true;
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        #endregion

        #region Notinused

        //string AsOFDate = txtAUMDate.Text;
        //if (AsOFDate == "")
        //    AsOFDate = "03/31/2016";
        //DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);


        //string ssi_Startdate = row.Cells[33].Text;
        //int ind = ssi_Startdate.IndexOf(" ");
        //ssi_Startdate = ssi_Startdate.Substring(0, ind);
        //DateTime startDate = DateTime.ParseExact(ssi_Startdate, "M/dd/yyyy", CultureInfo.InvariantCulture);

        //string ssi_EndDate = row.Cells[34].Text;
        //ind = ssi_EndDate.IndexOf(" ");
        //ssi_EndDate = ssi_EndDate.Substring(0, ind);
        //DateTime EndDate = DateTime.ParseExact(ssi_EndDate, "M/d/yyyy", CultureInfo.InvariantCulture);

        //string MinDate = row.Cells[40].Text;
        //ind = MinDate.IndexOf(" ");
        //MinDate = MinDate.Substring(0, ind);
        //DateTime dMinDate = DateTime.ParseExact(MinDate, "M/d/yyyy", CultureInfo.InvariantCulture);
        #endregion
        string ssi_Startdate = null;
        string vBillingExceptionID = null;

        if (Type == null || Type == "Billing")
        {
            vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 
            ssi_Startdate = row.Cells[34].Text;
        }
        else if (Type == "AUM")
        {
            vBillingExceptionID = row.Cells[36].Text;
            ssi_Startdate = row.Cells[38].Text;
        }
        else if (Type == "NONGA")
        {
            vBillingExceptionID = row.Cells[60].Text;
            ssi_Startdate = row.Cells[61].Text;
        }

        string AsofDate = txtAUMDate.Text;// 03/31/2016

        //if (AsofDate == "")
        //    AsofDate = "03/31/2016";

        if (Isclose != null)
            ssi_Startdate = AsofDate;


        //int ind = ssi_Startdate.IndexOf(" ");
        //ssi_Startdate = ssi_Startdate.Substring(0, ind);
        DateTime startDate = DateTime.ParseExact(ssi_Startdate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
        if (bProceed)
        {
            //if (dAsOFDate == startDate)
            //{
            // update  only
            //  ssi_billingexception objBillingException = new ssi_billingexception();
            Entity objBillingException = new Entity("ssi_billingexception");

            //objBillingException.ssi_billingexceptionid = new Key();
            //objBillingException.ssi_billingexceptionid.Value = new Guid(Convert.ToString(vBillingExceptionID));
            objBillingException["ssi_billingexceptionid"] = new Guid(Convert.ToString(vBillingExceptionID));

            //objBillingException.ssi_enddate = new CrmDateTime();
            //objBillingException.ssi_enddate.Value = startDate.AddDays(-1).ToString();
            objBillingException["ssi_enddate"] = startDate.AddDays(-1);

            service.Update(objBillingException);



            if (Isclose == null)
            {

                DeleteException(row, Type);
            }



            #region Notinused
            //}

            //else if (startDate > dAsOFDate)
            //{
            //    // update with asofdate -1 
            //    // insert new record

            //    ssi_billingexception objBillingException = new ssi_billingexception();
            //    objBillingException.ssi_billingexceptionid = new Key();
            //    objBillingException.ssi_billingexceptionid.Value = new Guid(Convert.ToString(vBillingExceptionID));

            //    objBillingException.ssi_enddate = new CrmDateTime();
            //    objBillingException.ssi_enddate.Value = dAsOFDate.AddDays(-1).ToString();

            //    service.Update(objBillingException);



            //}
            //else if (dAsOFDate < startDate)
            //{
            //    // sd=asofdate
            //    // enddate =-1
            //}
            #endregion
        }

    }

    //public void DeleteException(GridViewRow row, string Type = null)
    //{

    //    string vBillingExceptionID = null;

    //    if (Type == null || Type == "Billing")
    //    {
    //        vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 

    //    }
    //    else if (Type == "AUM")
    //    {
    //        vBillingExceptionID = row.Cells[36].Text;

    //    }

    //   // Guid gUId = new Guid(vBillingExceptionID);

    //    //ClientCredentials Credentials = new ClientCredentials();
    //    //Credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;
    //    //  string crmUrl =AppLogic.GetParam( AppLogic.ConfigParam.CRMUrl).ToString();

    //  //  Uri OrganizationUri = new Uri("http://crm-test3/GreshamPartners/XRMServices/2011/Organization.svc");
    //  //  Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
    //    // Uri OrganizationUri = new Uri(crmUrl);
    //    //Uri HomeRealmUri = null;
    //    //using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null))
    //    //{
    //    //    IOrganizationService iservice = (IOrganizationService)serviceProxy;
    //    //    SetStateRequest objSetStateRequest = new SetStateRequest()
    //    //    {
    //    //        EntityMoniker = new EntityReference("ssi_billingexception", gUId),
    //    //        State = new OptionSetValue(1),
    //    //        Status = new OptionSetValue(2),
    //    //    };

    //    //    iservice.Execute(objSetStateRequest);
    //    //}


    // //   #region CRM Connection
    // ////   string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
    // //   //string crmServerURL = "http://server:5555/";
    // // //  string orgName = "GreshamPartners";
    // //   //string orgName = "Webdev";
    // //   IOrganizationService service = null;

    // //   try
    // //   {
    // //       string UserId = GetcurrentUser();

    // //       service = GM.GetCrmService();
    // //       strDescription = "Crm Service starts successfully";
    // //   }
    // //   catch (System.Web.Services.Protocols.SoapException exc)
    // //   {
    // //       //  bProceed = false;
    // //       strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
    // //       lblMessage.Text = strDescription;
    // //   }
    // //   catch (Exception exc)
    // //   {
    // //       //bProceed = false;
    // //       strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
    // //       lblMessage.Text = strDescription;
    // //   }

    // //   #endregion

    // //   SetStateRequest objSetStateRequest1 = new SetStateRequest()
    // //   {
    // //       EntityMoniker = new EntityReference("ssi_billingexception", gUId),
    // //       State = new OptionSetValue(1),
    // //       Status = new OptionSetValue(2),
    // //   };

    // //   service.Execute(objSetStateRequest1);


    //    //DeactivateRecord(gUId,(OrganizationServiceProxy) service);


    //    //ssi_billingexceptionStateInfo sa = new ssi_billingexceptionStateInfo();
    //    //string a = sa.formattedvalue;

    //    //  service.Delete(EntityName.ssi_billingexception.ToString(), gUId);


    //    //   DeactivateRecord(gUId);


    //    //GeneralMethods GM = new GeneralMethods();
    //    //Guid gUId = new Guid(vBillingExceptionID);

    //    //ClientCredentials Credentials = new ClientCredentials();
    //    //Credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;


    //    //Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
    //    //// Uri OrganizationUri = new Uri(crmUrl);
    //    //Uri HomeRealmUri = null;
    //    //using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null))
    //    //{
    //    //    IOrganizationService iservice = (IOrganizationService)serviceProxy;

    //    //    //SetStateRequest objSetStateRequest = new SetStateRequest()
    //    //    //{
    //    //    //    EntityMoniker = new EntityReference("ssi_billingexception", gUId),
    //    //    //    State = new OptionSetValue(1),
    //    //    //    Status = new OptionSetValue(2),
    //    //    //};
    //    //    SetStateRequest objSetStateRequest = new SetStateRequest()
    //    //    {
    //    //        EntityMoniker = new EntityReference("ssi_billingexception", gUId),
    //    //        State = new OptionSetValue(1),
    //    //        Status = new OptionSetValue(2),
    //    //    };


    //    //    iservice.Execute(objSetStateRequest);
    //    //}



    //    //var EntityName = organizationService.Retrieve("ssi_billingexception", recordId, new ColumnSet(new[] { "statecode", "statuscode" }));
    //    //try
    //    //{
    //    //    if (EntityName != null)
    //    //    {
    //    //        SetStateRequest objSetStateRequest = new SetStateRequest()
    //    //        {
    //    //            EntityMoniker = new EntityReference("EntityName", recordId),
    //    //            State = new OptionSetValue(1),
    //    //            Status = new OptionSetValue(2),
    //    //        };
    //    //        organizationService.Execute(objSetStateRequest);
    //    //    }
    //    //}
    //    //catch (TimeoutException ex)
    //    //{
    //    //    //Exception Code
    //    //}


    //}


    public static void DeleteException(GridViewRow row, string Type = null)
    {

        string vBillingExceptionID = null;

        if (Type == null || Type == "Billing")
        {
            vBillingExceptionID = row.Cells[18].Text;  //17   BillingExceptionId 

        }
        else if (Type == "AUM")
        {
            vBillingExceptionID = row.Cells[36].Text;

        }
        else if (Type == "NONGA")
        {
            vBillingExceptionID = row.Cells[60].Text;

        }

        Guid gUId = new Guid(vBillingExceptionID);
        #region commented (11_21_2018) - IFD Changes
        //ClientCredentials Credentials = new ClientCredentials();
        //Credentials.Windows.ClientCredential = System.Net.CredentialCache.DefaultNetworkCredentials;


        //Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
        //// Uri OrganizationUri = new Uri(crmUrl);
        //Uri HomeRealmUri = null;
        //using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null))
        //{
        //    IOrganizationService iservice = (IOrganizationService)serviceProxy;

        //    Entity party = new Entity("ssi_billingexception", gUId);
        //    party["statecode"] = new OptionSetValue(1); //Status
        //    party["statuscode"] = new OptionSetValue(2); //Status reason
        //    iservice.Update(party);

        //}

        IOrganizationService service = null;
        GeneralMethods GM = new GeneralMethods();
        //try
        //{   
        service = GM.GetCrmService();// GetCrmService(crmServerUrl, orgName, UserId);

        Entity party = new Entity("ssi_billingexception", gUId);
        party["statecode"] = new OptionSetValue(1); //Status
        party["statuscode"] = new OptionSetValue(2); //Status reason
        service.Update(party);
        //}

        //catch (Exception exc)
        //{

        //}

        #endregion

    }


    public void SaveNew()
    {

        lblMessage.Text = "";

        //DataTable dtTempData = (DataTable)ViewState["dtTempData"];
        foreach (GridViewRow gvRow in gvBilling.Rows)
        {

            string BillingExceptionID = gvRow.Cells[18].Text;
            string AssetLevelFlg = gvRow.Cells[16].Text;
            string TotalAssetLevelFlg = gvRow.Cells[17].Text;

            string AumExceptionID = gvRow.Cells[36].Text;
            string NONGAExceptionID = gvRow.Cells[60].Text;

            string ACLevelBillingExceptionFlg = gvRow.Cells[14].Text;
            string ACLevelAUMExceptionFlg = gvRow.Cells[15].Text;
            //DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            BillingExceptionID = BillingExceptionID.Replace("&nbsp;", "");
            AumExceptionID = AumExceptionID.Replace("&nbsp;", "");
            NONGAExceptionID = NONGAExceptionID.Replace("&nbsp;", "");


            TextBox BillingPer = (TextBox)gvRow.FindControl("txtBillPer");   //4
            string vBillingPer = BillingPer.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            if (vBillingPer == "")
            {
                vBillingPer = "0";
            }
            TextBox Billingtxt = (TextBox)gvRow.FindControl("txtBilling");   //4
            string vBilling = Billingtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            TextBox NGAAtxt = (TextBox)gvRow.FindControl("txtNGAA");   //4
            string vNGAA = NGAAtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");


            TextBox AUMtxt = (TextBox)gvRow.FindControl("txtAUM");   //4
            string vAUM = AUMtxt.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            string SubBilling = gvRow.Cells[32].Text;
            SubBilling = SubBilling.Replace("&nbsp", "");

            string SubAum = gvRow.Cells[33].Text;
            SubAum = SubAum.Replace("&nbsp", "");

            string PositiveAum = gvRow.Cells[46].Text;
            PositiveAum = PositiveAum.Replace("&nbsp", "").Replace(";", "");


            string PositiveBilling = gvRow.Cells[44].Text;
            PositiveBilling = PositiveBilling.Replace("&nbsp", "").Replace(";", "");


            DateTime startDate;
            string ssi_Startdate = gvRow.Cells[34].Text;
            //int ind = ssi_Startdate.IndexOf(" ");
            //ssi_Startdate = ssi_Startdate.Substring(0, ind);
            if (ssi_Startdate == "1/1/1900")
                startDate = DateTime.ParseExact(ssi_Startdate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            else
            {
                startDate = DateTime.ParseExact(ssi_Startdate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            }

            string ssi_EndDate = gvRow.Cells[35].Text;
            //ind = ssi_EndDate.IndexOf(" ");
            //ssi_EndDate = ssi_EndDate.Substring(0, ind);
            //  DateTime EndDate = DateTime.ParseExact(ssi_EndDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            string MinDate = gvRow.Cells[41].Text;
            MinDate = MinDate.Replace("&nbsp;", "");
            if (MinDate != "")
            {
                //int ind = MinDate.IndexOf(" ");
                // MinDate = MinDate.Substring(0, ind);
                DateTime dMinDate = DateTime.ParseExact(MinDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            }
            CheckBox ExcludeBillingcb = null;
            CheckBox ExcludeAUMcb = null;
            bool ExcludeAUM = false, ExcludeBilling = false;
            if (TotalAssetLevelFlg == "False")
            {
                ExcludeBillingcb = (CheckBox)gvRow.FindControl("cbExcludeBilling");   //5
                ExcludeBilling = ExcludeBillingcb.Checked;

                ExcludeAUMcb = (CheckBox)gvRow.FindControl("cbExcludeAum");   //5
                ExcludeAUM = ExcludeAUMcb.Checked;
            }



            string vBillingExcludeFlg = gvRow.Cells[23].Text;  //22   BillingExcludeFlg

            string AsOFDate = txtAUMDate.Text;
            string[] date = AsOFDate.Split(new char[] { '/' });
            if (AsOFDate == "")
                AsOFDate = "03/31/2016";

            DateTime dAsOFDate = DateTime.ParseExact(AsOFDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            string BillingFlag = gvRow.Cells[42].Text;
            string AUMFlag = gvRow.Cells[43].Text;
            string vBillingMarketValue = gvRow.Cells[4].Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            //if(Convert.ToDecimal( vBillingMarketValue)==0) FinalBillingMarketValue
            //    vBillingMarketValue=gvRow.Cells[3].Text;

            string BillingExceptionType = gvRow.Cells[20].Text;    //   BillingExceptionType
            string AUMExceptionType = gvRow.Cells[21].Text;

            AUMExceptionType = AUMExceptionType.Replace("&nbsp;", "");
            BillingExceptionType = BillingExceptionType.Replace("&nbsp;", "");

            string MinFutureDate = gvRow.Cells[47].Text;
            string MinAumFutureDate = gvRow.Cells[48].Text;
            string MinNGAAFutureDate = gvRow.Cells[63].Text;

            MinFutureDate = MinFutureDate.Replace("&nbsp;", "");
            MinAumFutureDate = MinAumFutureDate.Replace("&nbsp;", "");
            MinNGAAFutureDate = MinNGAAFutureDate.Replace("&nbsp;", "");

            #region Billing
            if (BillingFlag == "" && BillingExceptionType == "")
            {
                if (BillingExceptionID == "")
                {
                    if (checkChangesForInsert(gvRow))
                    {
                        if (AssetLevelFlg == "False" && TotalAssetLevelFlg == "False")
                        {
                            if (ExcludeBilling)
                                insertCRM(gvRow, true, false, 0, 0, 0, 0, 0, "Billing");

                            else if (SubBilling != "")
                            {
                                insertCRM(gvRow, false, false, 0, Convert.ToDecimal(SubBilling), 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                            else if (PositiveBilling != "")
                            {
                                insertCRM(gvRow, false, false, 0, 0, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                            }
                            else if (Convert.ToDecimal(vBillingPer) != 0)
                            {
                                insertCRM(gvRow, false, false, 0, 0, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                        }

                        //insertRecord(gvRow);

                    }

                }
                else if (BillingExceptionID != "")
                {
                    if (!ExcludeBilling)
                    {
                        if ((SubBilling != "" || PositiveBilling != "" || Convert.ToDecimal(vBillingPer) != 0))  //&& ExcludeBilling !=Convert.ToBoolean( vBillingExcludeFlg)
                        {
                            if (!ExcludeBilling)
                                if (SubBilling != "")
                                {
                                    if (startDate < dAsOFDate)
                                    {
                                        if (MinFutureDate == "")
                                        {
                                            UpdatingDelete(gvRow, "Billing", "Close");
                                            // DeleteException(gvRow, "Billing");
                                            insertCRM(gvRow, false, false, 0, Convert.ToDecimal(SubBilling), 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                                        }
                                        else
                                        {
                                            UpdatingDelete(gvRow, "Billing", "Close");

                                        }
                                    }
                                    else if (startDate == dAsOFDate)
                                    {
                                        updateBill(gvRow, false, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                                        insertCRM(gvRow, false, false, 0, Convert.ToDecimal(SubBilling), 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                                    }

                                }
                                else if (PositiveBilling != "")
                                {
                                    if (startDate < dAsOFDate)
                                    {
                                        UpdatingDelete(gvRow, "Billing", "Close");
                                        //   DeleteException(gvRow, "Billing");
                                        insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                                    }

                                    else if (startDate == dAsOFDate)
                                    {
                                        updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                                    }
                                }
                                else
                                {
                                    if (Convert.ToDecimal(vBillingPer) != 0)
                                    {
                                        if (checkChangesForInsert(gvRow))
                                        {
                                            if (startDate < dAsOFDate)
                                            {
                                                UpdatingDelete(gvRow, "Billing", "Close");
                                                // DeleteException(gvRow, "Billing");
                                                insertCRM(gvRow, false, false, 0, 0, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                                            }
                                            else if (startDate == dAsOFDate)
                                            {
                                                updateBill(gvRow, false, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                                            }
                                        }
                                    }
                                    //updateBill(gvRow, ExcludeBilling, Convert.ToDecimal(SubBilling), Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                                }

                            //    if (vBillingExcludeFlg == "True")
                            //    {
                            //        if (!ExcludeBilling)
                            //        {
                            //            UpdatingDelete(gvRow);
                            //        }

                            //    }
                        }
                        else
                        {
                            UpdatingDelete(gvRow, "Billing");
                            // DeleteException(gvRow, "Billing");
                        }
                    }
                    else
                    {
                        if (checkChangesForInsert(gvRow))
                            updateBill(gvRow, ExcludeBilling, 0, 0, 0, "Billing");
                        // UpdatingDelete(gvRow);

                    }

                }

            }
            else if (BillingFlag == "3")     // excude billing
            {
                if (BillingExceptionID != "" && AssetLevelFlg != "True")   //delete
                {
                    if (startDate < dAsOFDate)
                    {
                        UpdatingDelete(gvRow, "Billing", "Close");
                    }
                    else if (startDate == dAsOFDate)
                    {
                        UpdatingDelete(gvRow, "Billing");
                        //  DeleteException(gvRow, "Billing");
                    }
                }
                else if (BillingExceptionID != "" && AssetLevelFlg == "True")
                {
                    if (checkChangesForInsert(gvRow))
                    {
                        if (startDate < dAsOFDate)
                        {
                            if (MinFutureDate == "")
                            {

                                UpdatingDelete(gvRow, "Billing", "Close");
                                insertCRM(gvRow, true, false, 0, 0, 0, 0, 0, "Billing");


                            }
                            else
                            {
                                UpdatingDelete(gvRow, "Billing", "Close");
                                insertCRM(gvRow, true, false, 0, 0, 0, 0, 0, "Billing", MinFutureDate);
                            }
                        }
                        else if (startDate == dAsOFDate)
                        {
                            if (ExcludeBilling)
                            {
                                updateBill(gvRow, ExcludeBilling, 0, 0, 0, "Billing");
                                // insertCRM(gvRow, true, false, 0, 0, 0, 0);

                            }
                            else if (SubBilling != "")
                            {
                                updateBill(gvRow, false, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                            else if (PositiveBilling != "")
                            {
                                updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                            }
                            else if (Convert.ToDecimal(vBillingPer) != 0)
                            {
                                updateBill(gvRow, false, 0, 0, 0, "Billing");
                            }

                            else
                            {
                                UpdatingDelete(gvRow, "Billing");
                                // DeleteException(gvRow, "Billing");
                            }


                        }
                    }
                }
                else if (BillingExceptionID == "" && TotalAssetLevelFlg == "False" && AssetLevelFlg == "True")
                {
                    if (checkChangesForInsert(gvRow))
                    {

                        insertCRM(gvRow, true, false, 0, 0, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                    }
                    //insertRecord(gvRow);
                    // updateBill(gvRow, true, 0, 0, 0, "Billing");
                }
            }
            else if (BillingFlag == "2")    // -ve  value
            {

                if (BillingExceptionID != "" && AssetLevelFlg == "True")   //delete
                {
                    if (checkChangesForInsert(gvRow))
                    {
                        if (startDate < dAsOFDate)
                        {
                            if (MinFutureDate == "")
                            {
                                UpdatingDelete(gvRow, "Billing", "Close");
                                insertCRM(gvRow, false, false, 0, Convert.ToDecimal(SubBilling), 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                            else
                            {
                                UpdatingDelete(gvRow, "Billing", "Close");
                                insertCRM(gvRow, false, false, 0, Convert.ToDecimal(SubBilling), 0, 0, Convert.ToDecimal(vBillingPer), "Billing", MinFutureDate);
                            }
                        }
                        else if (startDate == dAsOFDate)
                        {
                            updateBill(gvRow, false, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");

                        }
                    }
                }
                else
                {
                    if (checkChangesForInsert(gvRow))
                    {
                        if (BillingExceptionID == "")
                        {
                            if (SubBilling != "")
                            {
                                insertCRM(gvRow, false, false, Convert.ToDecimal(SubBilling), 0, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                            //insertRecord(gvRow);
                        }
                        else
                        {
                            updateBill(gvRow, false, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                        }
                    }
                }


            }
            else if (BillingFlag == "1")
            {
                if (BillingExceptionID != "" && AssetLevelFlg != "True")   //delete
                {

                    if (startDate < dAsOFDate)
                    {

                        UpdatingDelete(gvRow, "Billing", "Close");
                    }

                    else if (startDate == dAsOFDate)
                    {
                        UpdatingDelete(gvRow, "Billing");
                        //   DeleteException(gvRow, "Billing");
                    }
                }
                else if (BillingExceptionID != "" && AssetLevelFlg == "True")   //delete
                {
                    if (checkChangesForInsert(gvRow))
                    {
                        if (PositiveBilling != "")
                        {
                            if (startDate < dAsOFDate)
                            {
                                if (MinFutureDate == "")
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");
                                    insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                                }
                                else
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");
                                    insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vBilling), 0, Convert.ToDecimal(vBillingPer), "Billing", MinFutureDate);
                                }

                            }
                            else if (startDate == dAsOFDate)
                            {
                                updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                            }
                        }
                    }
                }
                else
                {
                    if (checkChangesForInsert(gvRow))
                    {
                        if (BillingExceptionID == "")
                        {
                            if (PositiveBilling != "")
                            {
                                insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(PositiveBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                            }
                        }
                        //insertRecord(gvRow);
                        else
                        {

                            updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                        }
                    }
                }


            }



            else if (BillingExceptionID != "")
            {
                if (!ExcludeBilling)
                {
                    if ((SubBilling != "" || PositiveBilling != "" || Convert.ToDecimal(vBillingPer) != 0))  //&& ExcludeBilling !=Convert.ToBoolean( vBillingExcludeFlg)
                    {

                        if (SubBilling != "")
                        {
                            if (startDate < dAsOFDate)
                            {
                                if (MinFutureDate == "")
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");
                                    insertCRM(gvRow, false, false, 0, Convert.ToDecimal(SubBilling), 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                                }
                                else
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");
                                    insertCRM(gvRow, false, false, 0, Convert.ToDecimal(SubBilling), 0, 0, Convert.ToDecimal(vBillingPer), "Billing", MinFutureDate);
                                }
                            }
                            else if (startDate == dAsOFDate)
                            {
                                updateBill(gvRow, false, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");

                            }

                        }
                        else if (PositiveBilling != "")
                        {
                            if (startDate < dAsOFDate)
                            {
                                if (MinFutureDate == "")
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");

                                    insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                                }
                                else
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");
                                    insertCRM(gvRow, false, false, 0, 0, Convert.ToDecimal(vBilling), 0, Convert.ToDecimal(vBillingPer), "Billing", MinFutureDate);
                                }
                            }

                            else if (startDate == dAsOFDate)
                            {
                                updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                            }
                        }
                        else if (Convert.ToDecimal(vBillingPer) != 0)
                        {
                            if (checkChangesForInsert(gvRow))
                            {
                                if (startDate < dAsOFDate)
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");

                                    insertCRM(gvRow, false, false, 0, 0, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                                }
                                else if (startDate == dAsOFDate)
                                {
                                    updateBill(gvRow, false, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                                }
                            }

                        }
                        //else
                        //{
                        //    updateBill(gvRow, ExcludeBilling, Convert.ToDecimal(SubBilling), Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                        //}

                        //    if (vBillingExcludeFlg == "True")
                        //    {
                        //        if (!ExcludeBilling)
                        //        {
                        //            UpdatingDelete(gvRow);
                        //        }

                        //    }
                    }
                    else
                    {
                        // if (vBillingExcludeFlg == "True") -- changed - 19_2_2020
                        if (ExcludeBilling)
                        {
                            if (startDate < dAsOFDate)
                            {
                                if (MinFutureDate == "")
                                {
                                    UpdatingDelete(gvRow, "Billing", "Close");
                                }
                                else
                                {
                                    insertCRM(gvRow, true, false, 0, 0, 0, 0, 0, "Billing", MinFutureDate);
                                }
                            }
                            else if (startDate == dAsOFDate)
                            {
                                UpdatingDelete(gvRow, "Billing");
                                // DeleteException(gvRow, "Billing");
                            }
                        }
                        else  //--added - 19_2_2020
                        {
                            
                             if (startDate == dAsOFDate)
                            {
                                UpdatingDelete(gvRow, "Billing");
                            }
                            else
                            {
                                UpdatingDelete(gvRow, "Billing", "Close");

                            }
                                
                        }
                    }
                }
                else
                {
                    if (checkChangesForInsert(gvRow))
                        updateBill(gvRow, ExcludeBilling, 0, 0, 0, "Billing");
                    // UpdatingDelete(gvRow);

                }


                //if (!ExcludeBilling)
                //{
                //    if (SubBilling != "")
                //    {
                //        updateBill(gvRow, ExcludeBilling, Convert.ToDecimal(SubBilling), 0, Convert.ToDecimal(vBillingPer), "Billing");
                //    }
                //    else if (PositiveBilling != "")
                //    {
                //        updateBill(gvRow, ExcludeBilling, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                //    }
                //    else
                //    {
                //        if (checkChangesForInsert(gvRow))
                //            UpdatingDelete(gvRow);
                //    }
                //}
                //else
                //{
                //    if (checkChangesForInsert(gvRow))
                //        updateBill(gvRow, ExcludeBilling, 0, 0, Convert.ToDecimal(vBillingPer), "Billing");
                //}

            }
            else // added -- 19_2_2020
            {
                if (checkChangesForInsert(gvRow))
                {
                    if (BillingExceptionID == "")
                    {

                        insertCRM(gvRow, false, false, 0, 0, 0, 0, 0, "Billing");

                    }
                    //insertRecord(gvRow);
                    else
                    {

                        updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveBilling), Convert.ToDecimal(vBillingPer), "Billing");
                    }
                }
            }
            #endregion

            #region AUM
            DateTime dAUMstartDate;
            string AUMStartDate = gvRow.Cells[38].Text;
            dAUMstartDate = DateTime.ParseExact(AUMStartDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            if (AUMFlag == "" && AUMExceptionType == "")
            {
                if (AumExceptionID == "")
                {
                    if (checkChangesForInsertAum(gvRow))
                    {
                        if (AssetLevelFlg == "False" && TotalAssetLevelFlg == "False")
                        {
                            // insertRecord(gvRow);
                            if (ExcludeAUM)
                                insertCRM(gvRow, true, true, 0, 0, 0, 0, 0, "AUM");

                            else if (SubAum != "")
                            {
                                insertCRM(gvRow, false, true, 0, Convert.ToDecimal(SubAum), 0, 0, 0, "AUM");
                            }
                            else if (PositiveAum != "")
                            {
                                insertCRM(gvRow, false, true, 0, 0, 0, Convert.ToDecimal(PositiveAum), 0, "AUM");
                            }
                            //else if (Convert.ToDecimal(vBillingPer) != 0)
                            //{
                            //    insertCRM(gvRow, false, false, 0, 0, 0, 0);
                            //}
                        }
                    }
                }
            }
            else if (AUMFlag == "3")   // Exclude 
            {
                if (AumExceptionID != "" && AssetLevelFlg != "True")   //delete
                {
                    if (dAUMstartDate < dAsOFDate)
                    {
                        UpdatingDelete(gvRow, "AUM", "Close");
                    }
                    else if (dAUMstartDate == dAsOFDate)
                    {
                        UpdatingDelete(gvRow, "AUM");
                        // DeleteException(gvRow, "AUM");
                    }
                }
                else if (AumExceptionID != "" && AssetLevelFlg == "True")   //delete
                {
                    if (checkChangesForInsertAum(gvRow))
                    {
                        if (dAUMstartDate < dAsOFDate)
                        {
                            if (MinAumFutureDate == "")
                            {
                                UpdatingDelete(gvRow, "AUM", "Close");
                                insertCRM(gvRow, true, true, 0, 0, 0, 0, 0, "AUM");
                            }
                            else
                            {
                                UpdatingDelete(gvRow, "AUM", "Close");
                                insertCRM(gvRow, true, true, 0, 0, 0, 0, 0, "AUM", MinAumFutureDate);
                            }
                        }
                        else if (dAUMstartDate == dAsOFDate)
                        {
                            if (ExcludeAUM)
                            {
                                updateBill(gvRow, ExcludeAUM, 0, 0, 0, "AUM");
                                // insertCRM(gvRow, true, false, 0, 0, 0, 0);

                            }
                            else if (PositiveAum != "")
                            {
                                updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveAum), 0, "AUM");
                            }
                            else if (SubAum != "")
                            {
                                updateBill(gvRow, false, Convert.ToDecimal(SubAum), 0, 0, "AUM");
                            }
                            else
                            {
                                UpdatingDelete(gvRow, "AUM");
                                //  DeleteException(gvRow, "AUM");
                            }


                        }
                    }
                }
                else if (AumExceptionID == "" && TotalAssetLevelFlg == "False" && AssetLevelFlg == "True")
                {
                    //  insertRecord(gvRow);
                    // updateBill(gvRow, true, 0, 0, 0, "Billing");
                    insertCRM(gvRow, true, true, 0, 0, 0, 0, 0, "AUM");
                }



            }
            else if (AUMFlag == "2")    // -ve 
            {
                if (checkChangesForInsertAum(gvRow))
                {
                    if (AumExceptionID != "" && AssetLevelFlg == "True")   //delete
                    {

                        if (dAUMstartDate < dAsOFDate)
                        {
                            if (MinAumFutureDate == "")
                            {
                                UpdatingDelete(gvRow, "AUM", "Close");
                                insertCRM(gvRow, false, true, 0, Convert.ToDecimal(SubAum), 0, 0, 0, "AUM");
                            }
                            else
                            {
                                UpdatingDelete(gvRow, "AUM", "Close");
                                insertCRM(gvRow, false, true, 0, Convert.ToDecimal(SubAum), 0, 0, 0, "AUM", MinAumFutureDate);
                            }
                        }
                        else if (dAUMstartDate == dAsOFDate)
                        {
                            updateBill(gvRow, false, Convert.ToDecimal(SubAum), 0, 0, "AUM");

                        }
                    }
                    else
                    {

                        if (checkChangesForInsertAum(gvRow))
                        {
                            if (AumExceptionID == "")
                            {
                                insertCRM(gvRow, false, true, Convert.ToDecimal(SubAum), 0, 0, 0, 0, "AUM");
                            }
                            else
                            {
                                updateBill(gvRow, false, Convert.ToDecimal(SubAum), 0, 0, "AUM");
                            }
                            //insertRecord(gvRow);
                        }
                    }

                }
            }
            else if (AUMFlag == "1")
            {
                if (AumExceptionID != "" && AssetLevelFlg != "True")   //delete
                {
                    if (dAUMstartDate < dAsOFDate)
                    {
                        UpdatingDelete(gvRow, "AUM", "Close");
                    }
                    else if (dAUMstartDate == dAsOFDate)
                    {
                        UpdatingDelete(gvRow, "AUM");
                        // DeleteException(gvRow, "AUM");
                    }

                }
                else if (AumExceptionID != "" && AssetLevelFlg == "True")   //delete
                {
                    if (checkChangesForInsertAum(gvRow))
                    {
                        if (dAUMstartDate < dAsOFDate)
                        {
                            if (MinAumFutureDate == "")
                            {
                                UpdatingDelete(gvRow, "AUM", "Close");
                                insertCRM(gvRow, false, true, 0, 0, Convert.ToDecimal(vAUM), 0, 0, "AUM");
                            }
                            else
                            {
                                UpdatingDelete(gvRow, "AUM", "Close");
                                insertCRM(gvRow, false, true, 0, 0, Convert.ToDecimal(vAUM), 0, 0, "AUM", MinAumFutureDate);
                            }

                        }
                        else if (dAUMstartDate == dAsOFDate)
                        {
                            updateBill(gvRow, false, 0, Convert.ToDecimal(vAUM), 0, "AUM");
                        }
                    }
                }
                else
                {
                    if (checkChangesForInsertAum(gvRow))
                    {
                        if (AumExceptionID == "")
                            if (PositiveAum != "")
                            {
                                insertCRM(gvRow, false, true, 0, 0, 0, Convert.ToDecimal(vAUM), 0, "AUM");
                            }
                            //insertRecord(gvRow);
                            else
                            {

                                updateBill(gvRow, false, 0, Convert.ToDecimal(vBilling), 0, "AUM");
                            }
                    }
                }
            }


            else if (AumExceptionID != "")
            {
                if (!ExcludeAUM)
                {

                    if ((SubAum != "" || PositiveAum != ""))  //&& ExcludeBilling !=Convert.ToBoolean( vBillingExcludeFlg)
                    {
                        if (SubAum != "")
                        {
                            if (dAUMstartDate < dAsOFDate)
                            {
                                if (MinAumFutureDate == "")
                                {
                                    UpdatingDelete(gvRow, "AUM", "Close");
                                    insertCRM(gvRow, false, true, 0, Convert.ToDecimal(SubAum), 0, 0, 0, "AUM");
                                }
                                else
                                {
                                    UpdatingDelete(gvRow, "AUM", "Close");
                                    insertCRM(gvRow, false, true, 0, Convert.ToDecimal(SubAum), 0, 0, 0, "AUM", MinAumFutureDate);
                                }
                            }
                            else if (dAUMstartDate == dAsOFDate)
                            {
                                updateBill(gvRow, false, Convert.ToDecimal(SubAum), 0, 0, "AUM");
                            }
                        }
                        else if (PositiveAum != "")
                        {
                            if (dAUMstartDate < dAsOFDate)
                            {
                                if (MinAumFutureDate == "")
                                {

                                    UpdatingDelete(gvRow, "AUM", "Close");
                                    insertCRM(gvRow, false, true, 0, 0, Convert.ToDecimal(vAUM), 0, 0, "AUM");
                                }
                                else
                                {
                                    UpdatingDelete(gvRow, "AUM", "Close");
                                    insertCRM(gvRow, false, true, 0, 0, Convert.ToDecimal(vAUM), 0, 0, "AUM", MinAumFutureDate);
                                }
                            }
                            else if (dAUMstartDate == dAsOFDate)
                            {
                                updateBill(gvRow, false, 0, Convert.ToDecimal(vAUM), 0, "AUM");
                            }
                        }
                        //else if (Convert.ToDecimal(vBillingPer) != 0)
                        //{
                        //    updateBill(gvRow, false, Convert.ToDecimal(SubAum), Convert.ToDecimal(vAUM), 0, "AUM", true);
                        //}
                    }
                    else
                    {
                        if (checkChangesForInsertAum(gvRow))
                            if (dAUMstartDate < dAsOFDate)
                            {
                                if (MinAumFutureDate == "")
                                {

                                    UpdatingDelete(gvRow, "AUM", "Close");
                                }
                                else
                                {
                                    UpdatingDelete(gvRow, "AUM", "Close");
                                    insertCRM(gvRow, true, true, 0, 0, 0, 0, 0, "AUM", MinAumFutureDate);
                                }
                            }
                            else if (dAUMstartDate == dAsOFDate)
                            {

                                UpdatingDelete(gvRow, "AUM");
                                // DeleteException(gvRow, "AUM");
                            }

                    }

                }
                else
                {
                    if (checkChangesForInsertAum(gvRow))
                        updateBill(gvRow, ExcludeAUM, 0, 0, 0, "AUM");

                }

                //    if (vBillingExcludeFlg == "True")
                //    {
                //        if (!ExcludeBilling)
                //        {
                //            UpdatingDelete(gvRow);
                //        }

                //    }


            }
            else // added -- 19_2_2020
            {
                if (checkChangesForInsertAum(gvRow))
                {
                    if (AumExceptionID == "")
                    {

                        insertCRM(gvRow, false, false, 0, 0, 0, 0, 0, "AUM");

                    }
                    else
                    {

                        updateBill(gvRow, false, 0, Convert.ToDecimal(PositiveAum), 0, "AUM");
                    }
                }
            }
            #endregion


            #region NONGA
            DateTime dNGAAstartDate;
            string NGAAstartDate = gvRow.Cells[61].Text;
            dNGAAstartDate = DateTime.ParseExact(NGAAstartDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            if (NONGAExceptionID == "")
            {
                if (checkChangesForInsertNGAA(gvRow))
                {
                    insertCRM(gvRow, false, false, 0, 0, 0, 0, 0, "NONGA");
                }
            }
            else
            {
                if (checkChangesForInsertNGAA(gvRow))
                {
                    if (dNGAAstartDate < dAsOFDate)
                    {
                        if (MinNGAAFutureDate == "")
                        {
                            UpdatingDelete(gvRow, "NONGA", "Close");

                            insertCRM(gvRow, false, false, 0, 0, 0, 0, 0, "NONGA");
                        }
                        else
                        {
                            UpdatingDelete(gvRow, "NONGA", "Close");
                            insertCRM(gvRow, false, false, 0, 0, 0, 0, 0, "NONGA", MinNGAAFutureDate);
                        }
                    }

                    else if (dNGAAstartDate == dAsOFDate)
                    {
                        updateBill(gvRow, false, 0, 0, 0, "NONGA");
                    }
                }

            }

            #endregion
        }
        BindGrideview();
        lblMessage.Text = "Records Saved Successfully";

    }

    protected void btnSaveGenrate_Click(object sender, EventArgs e)
    {
        SaveNew();
        GenratePortfolioOld();

    }

    public void GenratePortfolioOld1()
    {
        try
        {
            // int liPageSize = 26;  //--> Original Value
            int liPageSize = 22;

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

            string lsSQL = "EXEC  SP_R_BILLING @ReportFlg = 1,  @BillingForUUID = '" + billingVal + "',@AsOfDate = '" + dDate + "'";

            newdataset = clsDB.getDataSet(lsSQL);

            int DSCount = newdataset.Tables[0].Rows.Count;
            DataTable table = newdataset.Tables[0].Copy();
            table.Columns["PosAssetClassName"].SetOrdinal(0);
            table.Columns["SecurityName"].SetOrdinal(1);
            //table.Columns["AccountLegalEntityName"].SetOrdinal(2);
            table.Columns["PdfAccountName"].SetOrdinal(2);
            table.Columns["ssi_BillingMarketValue"].SetOrdinal(3);
            table.Columns["FinalBillingMarketValue"].SetOrdinal(4);
            table.Columns["FinalAUMMarketValue"].SetOrdinal(5);

            Random rand = new Random();
            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();


            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            //  iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
            String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingrpt.pdf";
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strGUID + ".xls";

            string HHName = ddlBillFor.SelectedItem.Text;
            //string strTitle = Convert.ToString(newdataset.Tables[2].Rows[0][0]);
            //if (strTitle != "")
            //    HHName = strTitle;

            string strheader = "Billing Report";
            string Title = "How Have My Gresham Advised Assets Performed vs. Their Benchmarks?";

            //   DateTime asofDT = Convert.ToDateTime("12/31/2015");
            //   string _AsOfDate = Convert.ToString(asofDT.ToString("MMMM")) + " " + Convert.ToString(asofDT.Day) + ", " + Convert.ToString(asofDT.Year);


            iTextSharp.text.Table loTable = new iTextSharp.text.Table(6, table.Rows.Count);   // 2 rows, 2 columns           
            // lsTotalNumberofColumns = "9";
            iTextSharp.text.Cell loCell = new Cell();


            #region Table Style
            int[] headerwidths9 = { 15, 23, 24, 10, 10, 10 };
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
                liPageSize = 22;
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

                    lochunk = new Chunk("\n " + txtAUMDate.Text + " \n", setFontsAllFrutiger(10, 0, 1));
                    loCell.Add(lochunk);
                    loCell.Border = 0;
                    loTable.AddCell(loCell);



                    // Add table columns
                    for (int k = 0; k < colsize; k++)
                    {
                        string ColHeader = Convert.ToString(table.Columns[k].ColumnName);
                        if (ColHeader == "PosAssetClassName")
                            ColHeader = "Asset Class (Position)";
                        if (ColHeader.Contains("SecurityName"))
                            ColHeader = "Security Name";
                        if (ColHeader.Contains("PdfAccountName"))
                            ColHeader = "Legal Entity (Account)";
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
                        loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                        //if (k != 0)
                        //    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                        //else
                        loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                        loTable.AddCell(loCell);
                    }

                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                    png.SetAbsolutePosition(45, 557);//540
                    png.ScalePercent(10);
                    document.Add(png);
                }

                for (int j = 0; j < colsize; j++)
                {
                    // string cellBackgroundColor = Convert.ToString(table.Rows[i]["ColourCode"]);
                    string ColValue = Convert.ToString(table.Rows[i][j]);

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

                    if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True") // if it is total asset
                    {
                        //if (j == 0 && ColValue != "")
                        //    ColValue = ColValue + " Total";

                        if ((j == 3 || j == 4 || j == 5) && ColValue != "")
                        {
                            if (ColValue.Contains("-"))
                            {
                                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                ColValue = ColValue.Replace("(", "($");
                            }
                            else
                                ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                        }

                        //lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        # region Color on PDF


                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            lochunk.SetBackground(Color.YELLOW);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            lochunk.SetBackground(Color.YELLOW);
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
                                ColValue = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                                ColValue = ColValue.Replace("(", "($");
                            }
                            else
                                ColValue = String.Format("${0:#,###0;(#,###0)}", Convert.ToDecimal(ColValue));
                        }

                        if (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True")
                        {
                            if (j == 0)
                                ColValue = "Total";

                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                        }
                        else
                            lochunk = new Chunk(ColValue, setFontsAll(7, 0, 0));

                        # region Color on PDF


                        if (j == 4 && ((BillExcepTyp == "2" || BillExcepTyp == "3") || BillFeeTyp != ""))// || Convert.ToString(table.Rows[i]["BillingExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            lochunk.SetBackground(Color.YELLOW);
                        }

                        else if (j == 5 && (AUMExcepTyp == "2" || AUMExcepTyp == "3"))// || Convert.ToString(table.Rows[i]["AUMExceptionType"]) == "3"))
                        {
                            lochunk = new Chunk(ColValue, setFontsAll(7, 1, 0));
                            lochunk.SetBackground(Color.YELLOW);
                        }
                        else
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

                    loCell = new iTextSharp.text.Cell();
                    loCell.Add(lochunk);
                    // if ((i == 0 || i == 1) && j == 0)
                    //   loCell.Add(lochunknew);
                    loCell.Border = 0;

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

                    if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True" || (Convert.ToString(table.Rows[i]["TotalLevelFlg"]) == "True"))
                    {
                        loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
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

                    FileInfo loFile = new FileInfo(ls);
                    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                    //Response.Write("<script>");
                    //string lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + strGUID + ".pdf";
                    //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    ////  Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                    //Response.Write("</script>");

                    //Changed on 22_8_2019 --> ADFS LOGOUT ISSUE
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");
                    sb.Append("\nwindow.open('ViewReport.aspx?" + strGUID + ".pdf" + "', 'mywindow');");
                    sb.Append("</script>");
                    ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
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

            strGUID = strGUID.Replace("/", "");
            //  DestinationPath = ddlBillFor.SelectedItem.Text + " " + AUMdate.ToString("yyyy-MMdd") + ".pdf";

            //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4);
            //  iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
            String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingrpt.pdf";

            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strGUID + ".PDF";
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
                        #region Color on PDF


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

                        #region Color on PDF
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
                    //  fsFinalLocation = fsFinalLocation.Replace("'", "");
                    if (File.Exists(fsFinalLocation))
                    {
                        File.Delete(fsFinalLocation);
                    }

                    FileInfo loFile = new FileInfo(ls);

                    loFile.MoveTo(fsFinalLocation);
                    //Response.Write("<script>");
                    //strGUID = strGUID.Replace("'", "%27");
                    //string lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + strGUID + ".pdf";
                    //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    ////  Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                    //Response.Write("</script>");

                    //Changed on 22_8_2019 --> ADFS LOGOUT ISSUE
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");
                    sb.Append("\nwindow.open('ViewReport.aspx?" + strGUID + ".pdf" + "', 'mywindow');");
                    sb.Append("</script>");
                    ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
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

    public void GenratePortfolioOldwithNonGA()
    {
        DataSet dsBillingDropDowns = (DataSet)ViewState["BillingDropDowns"];
        DataTable dtAdminCategory = dsBillingDropDowns.Tables[0];
        DataTable dtNonGAServiceType = dsBillingDropDowns.Tables[1];

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
            int liPageSize = 19;

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
            table.Columns["FinalNGAAMarketValue"].SetOrdinal(6);
            table.Columns["ssi_AdminCategory"].SetOrdinal(7);
            table.Columns["ssi_NonGAServiceType"].SetOrdinal(8);
            Random rand = new Random();
            //   string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + rand.Next().ToString();
            string strGUID = ddlBillFor.SelectedItem.Text + " " + AUMdate.ToString("yyyy-MMdd");

            strGUID = strGUID.Replace("/", "");
            //  DestinationPath = ddlBillFor.SelectedItem.Text + " " + AUMdate.ToString("yyyy-MMdd") + ".pdf";

            //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate());
            //  iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
            String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingrpt.pdf";

            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strGUID + ".PDF";
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

            iTextSharp.text.Table loTable = new iTextSharp.text.Table(9, table.Rows.Count);   // 2 rows, 2 columns      

            // lsTotalNumberofColumns = "9";
            iTextSharp.text.Cell loCell = new Cell();


            #region Table Style
            //  int[] headerwidths9 = { 17, 31, 10, 13, 13, 13 };
            int[] headerwidths9 = { 13, 20, 7, 10, 10, 10, 10, 10, 10 };
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
                // liPageSize = 26; //--> Original Value
                liPageSize = 19;
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
                    loTable = new iTextSharp.text.Table(9, table.Rows.Count);
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
                    loCell.Colspan = 9;
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
                        if (ColHeader.Contains("FinalNGAAMarketValue"))
                            ColHeader = "Non-GA Admin";
                        if (ColHeader.Contains("ssi_AdminCategory"))
                            ColHeader = "Admin Category";
                        if (ColHeader.Contains("ssi_NonGAServiceType"))
                            ColHeader = "Non-GA Service Type";


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
                    if (j == 7)
                    {
                        if (ColValue != "")
                        {
                            string DisplayName = string.Empty;
                            for (int l = 0; l < dtAdminCategory.Rows.Count; l++)
                            {
                                string ID = dtAdminCategory.Rows[l]["ID"].ToString();
                                if (ID == ColValue)
                                {
                                    ColValue = dtAdminCategory.Rows[l]["DisplayTxt"].ToString();
                                    break;
                                }
                            }
                        }
                    }
                    else if (j == 8)
                    {
                        if (ColValue != "")
                        {
                            string DisplayName = string.Empty;
                            for (int l = 0; l < dtNonGAServiceType.Rows.Count; l++)
                            {
                                string ID = dtNonGAServiceType.Rows[l]["ID"].ToString();
                                if (ID == ColValue)
                                {
                                    ColValue = dtNonGAServiceType.Rows[l]["DisplayTxt"].ToString();
                                    break;
                                }
                            }
                        }
                    }
                    if (Convert.ToString(table.Rows[i]["AssetLevelFlg"]) == "True") // if it is total asset
                    {
                        //if (j == 0 && ColValue != "")
                        //    ColValue = ColValue + " Total";

                        if ((j == 3 || j == 4 || j == 5 || j == 6) && ColValue != "")
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
                        #region Color on PDF


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
                        if ((j == 3 || j == 4 || j == 5 || j == 6) && ColValue != "")
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

                        #region Color on PDF
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
                    //  fsFinalLocation = fsFinalLocation.Replace("'", "");
                    if (File.Exists(fsFinalLocation))
                    {
                        File.Delete(fsFinalLocation);
                    }

                    FileInfo loFile = new FileInfo(ls);

                    loFile.MoveTo(fsFinalLocation);
                    //Response.Write("<script>");
                    //strGUID = strGUID.Replace("'", "%27");
                    //string lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + strGUID + ".pdf";
                    //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    ////  Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                    //Response.Write("</script>");

                    //Changed on 22_8_2019 --> ADFS LOGOUT ISSUE
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");
                    sb.Append("\nwindow.open('ViewReport.aspx?" + strGUID + ".pdf" + "', 'mywindow');");
                    sb.Append("</script>");
                    ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
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
        //  SignatureJpg.SetAbsolutePosition(43, 800); // potrait logo
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

    protected void ddlBillFor_SelectedIndexChanged(object sender, EventArgs e)
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

        if (dsBilling.Tables[0].Rows.Count > 2)
        {
            ddlHH.SelectedIndex = 0;

        }
        else
        {
            ddlHH.SelectedValue = Convert.ToString(dsBilling.Tables[0].Rows[1]["HouseHoldUUID"]);
        }

        lblMessageShow.Text = "";
        if (ddlBillFor.SelectedValue != "ALL")
        {

            //lblAUMDate.Visible = true;
            //lblBillingFor.Visible = true;
            //lblBillingFor.Text = ddlBillFor.SelectedItem.Text;
            //lblAUMDate.Text = txtAUMDate.Text;
            BindGrideview();
        }


    }

    protected void txtAUMDate_TextChanged(object sender, EventArgs e)
    {
        lblMessageShow.Text = "";
        lblMessage.Text = "";


        lblAUMDate.Visible = false;
        lblBillingFor.Visible = false;
        gvBilling.Visible = false;
        btnSave.Visible = false;
        btnSave0.Visible = false;
        btnSaveGenrate.Visible = false;
        btnSaveGenrate0.Visible = false;
        btnSaveGenratewithNonGA.Visible = false;
        btnSaveGenratewithNonGA0.Visible = false;
        if (txtAUMDate.Text != "")
        {
            //lblAUMDate.Visible = true;
            //lblBillingFor.Visible = true;
            //lblBillingFor.Text = ddlBillFor.SelectedItem.Text;
            //lblAUMDate.Text = txtAUMDate.Text;
            BindGrideview();
        }
    }





    protected void txtNGAA_TextChanged(object sender, EventArgs e)
    {

    }

    protected void ddlAdminCategory_SelectedIndexChanged(object sender, EventArgs e)
    {
        try
        {
            lblMessage.Text = "";


            DropDownList ddlAdminCategory = (DropDownList)sender;
            GridViewRow gvrow1 = (GridViewRow)ddlAdminCategory.NamingContainer;
            string gridddlAdminCategory = gvrow1.Cells[55].Text.Replace("&nbsp;", "");
            string vddlAdminCategory = ddlAdminCategory.SelectedValue;

            //if (gridddlAdminCategory != vddlAdminCategory)
            //{
            //    ddlAdminCategory.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            //    gvrow1.Cells[58].Text = "1";
            //}
            //else
            //{
            //    gvrow1.Cells[58].Text = "";
            //}
            #region  Colour Asset Level
            int TotalRowIndex1 = 0;
            string vActualValue1 = null;
            foreach (GridViewRow gvrow in gvBilling.Rows)   // For updating Total asset Value 
            {
                TextBox tx = (TextBox)gvrow.FindControl("txtBillPer");
                string txtBillPer = tx.Text;
                string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                TotalRowIndex1 = gvrow.RowIndex;
                if (gvrow.Cells[32].Text != "")
                {
                    //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (gvrow.Cells[44].Text != "")
                {
                    //   gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (gvrow.Cells[33].Text != "")
                {
                    // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                //AUM Positive Value Change
                if (gvrow.Cells[46].Text != "")
                {
                    //   gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }


                ////ddlAACategory
                //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddl = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                //    ddl.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                ////ddlNonGAServiceType
                //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddlNonGAServiceType = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}

                if (txtBillPer != vBillingFeePct)
                {
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");


                //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                if (BillingType != "1" && BillingType != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (AUMType != "1" && AUMType != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }


                if (PerValue != "1" && PerValue != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string diffflag = gvrow.Cells[51].Text;
                if (diffflag == "1")
                {
                    gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                }


                vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                int index1 = gvrow.RowIndex;
                TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != val)
                    Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                {
                    Billingtext1.BackColor = System.Drawing.Color.White;
                }

                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }

                //}
                //  int index = gvrow.RowIndex;


                TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != aumval)
                    aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                    aumtext1.BackColor = System.Drawing.Color.White;


                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }







            }
            #endregion

            #region totalTextcolor
            TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
            decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));


            if (Convert.ToDecimal(vActualValue1) != value)
                Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                Billingtex.BackColor = System.Drawing.Color.White;

            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
            decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
            if (Convert.ToDecimal(vActualValue1) != aumval1)
                aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                aumtex.BackColor = System.Drawing.Color.White;


            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }





            #endregion


            decimal TotalBilling = 0, TotalAUM = 0;
            int TotalRowIndex = 0;
            // string vActualValue1 = null;
            foreach (GridViewRow gvrow in gvBilling.Rows)
            {
                string AssetLevelFlgNew = gvrow.Cells[16].Text;

                if (AssetLevelFlgNew == "True")
                {
                    TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                    decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                    TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                    decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                    TotalAUM = TotalAUM + AUMVal;


                    TotalBilling = TotalBilling + BillingVal;


                }
            }

            string TotalActual = gvBilling.Rows[TotalRowIndex1].Cells[4].Text;
            decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
            gvBilling.Rows[TotalRowIndex1].Cells[8].Text = currencyFormat(Totalval.ToString());
            Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
            gvBilling.Rows[TotalRowIndex1].Cells[6].Text = currencyFormat(Totalval.ToString());
        }
        catch (Exception ex)
        {
            Response.Write("Error Occurred: " + ex.Message.ToString());
        }
    }

    protected void ddlNonGAServiceType_SelectedIndexChanged(object sender, EventArgs e)
    {
        try
        {
            lblMessage.Text = "";


            DropDownList ddlNonGAServiceType = (DropDownList)sender;
            GridViewRow gvrow1 = (GridViewRow)ddlNonGAServiceType.NamingContainer;
            string gridddlNonGAServiceType = gvrow1.Cells[56].Text.Replace("&nbsp;", "");
            string vddlNonGAServiceType = ddlNonGAServiceType.SelectedValue;

            //if (gridddlNonGAServiceType != vddlNonGAServiceType)
            //{
            //    ddlNonGAServiceType.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            //    gvrow1.Cells[59].Text = "1";
            //}
            //else
            //{
            //    gvrow1.Cells[59].Text = "";
            //}
            #region  Colour Asset Level
            int TotalRowIndex1 = 0;
            string vActualValue1 = null;
            foreach (GridViewRow gvrow in gvBilling.Rows)   // For updating Total asset Value 
            {
                TextBox tx = (TextBox)gvrow.FindControl("txtBillPer");
                string txtBillPer = tx.Text;
                string vBillingFeePct = gvrow.Cells[27].Text;  //30  BillingFeePct 

                TotalRowIndex1 = gvrow.RowIndex;
                if (gvrow.Cells[32].Text != "")
                {
                    //  gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (gvrow.Cells[44].Text != "")
                {
                    //   gvrow.Cells[5].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (gvrow.Cells[33].Text != "")
                {
                    // gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                //AUM Positive Value Change
                if (gvrow.Cells[46].Text != "")
                {
                    //   gvrow.Cells[7].BackColor = System.Drawing.Color.FromName("#FCFB9C");
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }


                ////ddlAACategory
                //if (gvrow.Cells[58].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddlAdminCategory1 = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlAdminCategory");
                //    ddlAdminCategory1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}
                ////ddlNonGAServiceType
                //if (gvrow.Cells[59].Text.Replace("&nbsp;", "") != "")
                //{
                //    DropDownList ddlNonGAServiceType1 = (DropDownList)gvBilling.Rows[TotalRowIndex1].FindControl("ddlNonGAServiceType");
                //    ddlNonGAServiceType1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                //}

                if (txtBillPer != vBillingFeePct)
                {
                    TextBox text1 = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string PerValue = gvrow.Cells[19].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                string BillingType = gvrow.Cells[49].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");
                string AUMType = gvrow.Cells[50].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "");


                //     PositiveBilling = PositiveBilling.Replace("&nbsp", "");
                if (BillingType != "1" && BillingType != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBilling");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }

                if (AUMType != "1" && AUMType != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtAUM");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }


                if (PerValue != "1" && PerValue != "")
                {
                    int index = gvrow.RowIndex;
                    TextBox text1 = (TextBox)gvBilling.Rows[index].FindControl("txtBillPer");
                    text1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                }
                string diffflag = gvrow.Cells[51].Text;
                if (diffflag == "1")
                {
                    gvrow.BackColor = System.Drawing.Color.FromName("#C6EFCE");
                }


                vActualValue1 = gvrow.Cells[4].Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", "");
                int index1 = gvrow.RowIndex;
                TextBox Billingtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtBilling");
                decimal val = Convert.ToDecimal(Billingtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != val)
                    Billingtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                {
                    Billingtext1.BackColor = System.Drawing.Color.White;
                }

                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (val > 0 && Convert.ToDecimal(vActualValue1) < val)
                {
                    if (Convert.ToDecimal(vActualValue1) < val)
                        Billingtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }

                //}
                //  int index = gvrow.RowIndex;


                TextBox aumtext1 = (TextBox)gvBilling.Rows[index1].FindControl("txtAUM");
                decimal aumval = Convert.ToDecimal(aumtext1.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", "").Replace(",", ""));
                if (Convert.ToDecimal(vActualValue1) != aumval)
                    aumtext1.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                else
                    aumtext1.BackColor = System.Drawing.Color.White;


                if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }
                else if (aumval > 0 && Convert.ToDecimal(vActualValue1) < aumval)
                {
                    if (Convert.ToDecimal(vActualValue1) < aumval)
                        aumtext1.BackColor = System.Drawing.Color.FromName("#FF5353");
                }







            }
            #endregion

            #region totalTextcolor
            TextBox Billingtex = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtBilling");
            decimal value = Convert.ToDecimal(Billingtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));


            if (Convert.ToDecimal(vActualValue1) != value)
                Billingtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                Billingtex.BackColor = System.Drawing.Color.White;

            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (value > 0 && Convert.ToDecimal(vActualValue1) < value)
            {
                if (Convert.ToDecimal(vActualValue1) < value)
                    Billingtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }

            TextBox aumtex = (TextBox)gvBilling.Rows[TotalRowIndex1].FindControl("txtAUM");
            decimal aumval1 = Convert.ToDecimal(aumtex.Text.Replace("&nbsp", "").Replace(";", "").Replace("(", "-").Replace(")", "").Replace("$", ""));
            if (Convert.ToDecimal(vActualValue1) != aumval1)
                aumtex.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            else
                aumtex.BackColor = System.Drawing.Color.White;


            if (Convert.ToDecimal(vActualValue1) > 0)  // for red color
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }
            else if (aumval1 > 0 && Convert.ToDecimal(vActualValue1) < aumval1)
            {
                if (Convert.ToDecimal(vActualValue1) < aumval1)
                    aumtex.BackColor = System.Drawing.Color.FromName("#FF5353");
            }





            #endregion

            decimal TotalBilling = 0, TotalAUM = 0;
            int TotalRowIndex = 0;
            // string vActualValue1 = null;
            foreach (GridViewRow gvrow in gvBilling.Rows)
            {
                string AssetLevelFlgNew = gvrow.Cells[16].Text;

                if (AssetLevelFlgNew == "True")
                {
                    TextBox AUMText = (TextBox)gvrow.FindControl("txtAUM");
                    decimal AUMVal = Convert.ToDecimal(AUMText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                    TextBox BillingText = (TextBox)gvrow.FindControl("txtBilling");
                    decimal BillingVal = Convert.ToDecimal(BillingText.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

                    TotalAUM = TotalAUM + AUMVal;


                    TotalBilling = TotalBilling + BillingVal;


                }
            }

            string TotalActual = gvBilling.Rows[TotalRowIndex1].Cells[4].Text;
            decimal Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalAUM;
            gvBilling.Rows[TotalRowIndex1].Cells[8].Text = currencyFormat(Totalval.ToString());
            Totalval = Convert.ToDecimal(TotalActual.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace(",", "")) - TotalBilling;
            gvBilling.Rows[TotalRowIndex1].Cells[6].Text = currencyFormat(Totalval.ToString());
        }
        catch (Exception ex)
        {
            Response.Write("Error Occurred: " + ex.Message.ToString());
        }
    }

    protected void gvBilling_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            DataSet dsBillingDropDowns = (DataSet)ViewState["BillingDropDowns"];
            string AssetLevelFlgNew = e.Row.Cells[16].Text;
            string TotalAssetLevelFlgNew = e.Row.Cells[17].Text;
            string BillingExcludeFlg = e.Row.Cells[23].Text;
            string FinalNGAAMarketValue = e.Row.Cells[57].Text;

            TextBox txtNGAA = (TextBox)e.Row.FindControl("txtNGAA");
            txtNGAA.Enabled = false;

            DropDownList ddlAdminCategory = (DropDownList)e.Row.FindControl("ddlAdminCategory");
            DropDownList ddlNonGAServiceType = (DropDownList)e.Row.FindControl("ddlNonGAServiceType");


            string AdminCategoryId = e.Row.Cells[55].Text.Replace("&nbsp;", ""); ;
            string NonGAServiceTypeId = e.Row.Cells[56].Text.Replace("&nbsp;", ""); ;


            if (AssetLevelFlgNew != "True" && TotalAssetLevelFlgNew != "True")
            {
                ddlAdminCategory.Items.Clear();
                ddlNonGAServiceType.Items.Clear();
                // DS = clsdb.getDataSet(sqlstr);
                #region AdminCategory
                if (dsBillingDropDowns.Tables[0].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    dt = dsBillingDropDowns.Tables[0];
                    ddlAdminCategory.DataSource = dt;
                    ddlAdminCategory.DataTextField = "DisplayTxt";
                    ddlAdminCategory.DataValueField = "ID";
                    ddlAdminCategory.DataBind();
                    ddlAdminCategory.Items.Insert(0, "");
                    ddlAdminCategory.Items[0].Value = "";

                    ddlAdminCategory.SelectedValue = AdminCategoryId;
                }
                else
                {
                    ddlAdminCategory.SelectedIndex = ddlAdminCategory.SelectedIndex - 1;
                    ddlAdminCategory.Items.Insert(0, "No Record Found");
                    ddlAdminCategory.Items[0].Value = "";
                    ddlAdminCategory.SelectedIndex = 0;
                }
                #endregion
                #region ddlNonGAServiceType

                if (dsBillingDropDowns.Tables[1].Rows.Count > 0)
                {
                    DataTable dt = new DataTable();
                    dt = dsBillingDropDowns.Tables[1];
                    ddlNonGAServiceType.DataSource = dt;
                    ddlNonGAServiceType.DataTextField = "DisplayTxt";
                    ddlNonGAServiceType.DataValueField = "ID";
                    ddlNonGAServiceType.DataBind();
                    ddlNonGAServiceType.Items.Insert(0, "");
                    ddlNonGAServiceType.Items[0].Value = "";

                    ddlNonGAServiceType.SelectedValue = NonGAServiceTypeId;
                }
                else
                {
                    ddlNonGAServiceType.SelectedIndex = ddlNonGAServiceType.SelectedIndex - 1;
                    ddlNonGAServiceType.Items.Insert(0, "No Record Found");
                    ddlNonGAServiceType.Items[0].Value = "";
                    ddlNonGAServiceType.SelectedIndex = 0;
                }

                #endregion
            }
            else
            {
                ddlNonGAServiceType.Visible = false;
                ddlNonGAServiceType.SelectedValue = "";

                ddlAdminCategory.Visible = false;
                ddlAdminCategory.SelectedValue = "";
            }
            //if (BillingExcludeFlg == "True")
            //{
            TextBox txtNGAA1 = (TextBox)e.Row.FindControl("txtNGAA");
            txtNGAA1.Text = currencyFormat(FinalNGAAMarketValue);


            //}

        }
    }


    protected void btnSaveGenratewithNonGA_Click(object sender, EventArgs e)
    {
        SaveNew();
        GenratePortfolioOldwithNonGA();
    }
}