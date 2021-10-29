/*using System;TESTIN_G
using System.Collections.Generic;
*/using System.Data;TESTIN_G
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
//using CrmSdk;
using System.IO;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using iTextSharp.text;
using System.Globalization;
using Microsoft.SharePoint.Client;
using System.Security;
using Microsoft.Xrm.Sdk;
using OfficeOpenXml.Style;
using Spire.Xls;

using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class Billing_Invoice_Entry : System.Web.UI.Page
{
    string BillingUUID = "";
    string[] SourceFileArray;
    string FeeAmount = "0";
    List<string> ls = new List<string>();
    GeneralMethods GM = new GeneralMethods();

    string SharepointPath = "https://greshampartners.sharepoint.com/sites/DataBackup/";
    string Server = AppLogic.GetParam(AppLogic.ConfigParam.Server);
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            System.Web.UI.WebControls.ListItem itemToRemove = ddlClientType.Items.FindByValue("0");
            if (itemToRemove != null)
            {
                ddlClientType.Items.Remove(itemToRemove);
            }

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
            BindClientType();


            //txtBox1.Style.Add("display", "none");
            //lblbox1.Style.Add("display", "none");
            //txtBox2.Style.Add("display", "none");
            //lblbox2.Style.Add("display", "none");
            //txtBox3.Style.Add("display", "none");
            //lblbox3.Style.Add("display", "none");

            txtTotalAUM.Enabled = false;
            txtBillingAUM.Enabled = false;
            txtBillingAUM.Visible = false;
            //  Response.Write(Request.Url.AbsoluteUri);
        }

    }
    /** To Find The Nearest Quarter Date Union(QuatersInYear Column) & OrderBy LastDate **/
    public DateTime NearestQuarterEnd(DateTime date)
    {
        IEnumerable<DateTime> candidates =//Enumarateor used To show candiadtes by LastDate
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

    #region Populating the Dropdownlists & Textbox changed events
    protected void bindDropDowns()
    {
        try
        {
            /* Populating Year DropDown */
            string sqlstr = "[sp_s_Get_HouseHoldName] ";
            BindDropdown(ddlHH, sqlstr, "name", "accountid");
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
        }
    }

    protected void BindListBox(DataTable dtData)
    {

        lbGroup.DataSource = dtData;
        lbGroup.DataTextField = "BillingName";
        //lbGroup.DataValueField = "ProductID";
        lbGroup.DataBind();

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
            checkbillfortype();
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
        }
    }

    protected void BindClientType()//Function TO Bind ClientType DropDown
    {
        try
        {
            //Populate clientType
            string sqlstr = "[SP_S_BILLING_FEESCHEDULETYPE] ";//Fetch Record From [SP_S_BILLING_FEESCHEDULETYPE] -> Procedure
            BindDropdown(ddlClientType, sqlstr, "FeeScheduleTypeTxt", "FeeScheduleTypeId");

        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
        }
    }

    public void checkbillfortype()//Function TO Bind BillFor DropDown
    {
        /*check condition for selected value*/
        if (Convert.ToString(ddlBillFor.SelectedValue) != "ALL" && Convert.ToString(ddlBillFor.SelectedValue) != "")
        {
            string strFeeType = ddlBillFor.SelectedValue;
            char[] delimiterChars = { '|' };//spliting feetype and billinguid
            string[] words = strFeeType.ToString().Split(delimiterChars);
            int len = words.Length;
            if (len > 0)
                strFeeType = words[1];

            if (strFeeType.ToUpper() == "Flat".ToUpper())//feetype=flat
            {
                /*where FeeType=Flat
                 * Hidding AnnualFeeCalc and 
                 * RequiredFieldValidator7 */
                //  trStdAnnualFeeCalc.Style.Add("display", "none");
                lblStdAnnFeeCalc.Visible = false;
                txtStdAnnualFeeCalc.Visible = false;
                txtStdAnnualFeeCalc.Text = "";
                //    RequiredFieldValidator7.Visible = false;
            }
            else
            {
                //   RequiredFieldValidator7.Visible = true;
                // trStdAnnualFeeCalc.Style.Add("display", "");
                lblStdAnnFeeCalc.Visible = true;
                txtStdAnnualFeeCalc.Visible = true;
            }

            trRelationshipFee.Style.Add("display", "none");
            //  RequiredFieldValidator11.Visible = false;

            if (ddlClientType.SelectedValue == "100000007") //Custom
            {
                txtStdAnnualFeeCalc.Text = "";
                //  trStdAnnualFeeCalc.Style.Add("display", "none");
                lblStdAnnFeeCalc.Visible = false;
                txtStdAnnualFeeCalc.Visible = false;
                // txtCustFeeAmount.Focus();

                /** where FeeType=custom
                 * Display RequiredValidator8
                 *  Hidding StdAnnualFeecalc and 
                 *  RequiredFieldValidator7 **/
                RequiredFieldValidator7.Visible = false;
                RequiredFieldValidator8.Visible = true;
                txtDiscount.Visible = false;
                lbldiscount.Visible = false;
                lbldecimal.Visible = false;
            }

            else if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
            {
                /* Display StdAnnualFeeCalc
                 *  RequiredValidator7 
                 *  RequiredValidator8 
                 *  for FeeType=Standard with Relationship Fees */
                // trStdAnnualFeeCalc.Style.Add("display", "");
                lblStdAnnFeeCalc.Visible = true;
                txtStdAnnualFeeCalc.Visible = true;
                //   RequiredFieldValidator7.Visible = true;
                //   RequiredFieldValidator11.Visible = true;
                trRelationshipFee.Style.Add("display", "");

                txtDiscount.Visible = true;
                lbldiscount.Visible = true;
                lbldecimal.Visible = true;

            }
            else
            {
                /* for FeeType=other than Custom and Standard With Relationship Fees
                 * Display StdAnnualFeeCalc and 
                 * RequiredValidator7 
                 * Hide RelationshipFee and
                 *  RequiredValidator8 
                       */
                trRelationshipFee.Style.Add("display", "none");
                //  trStdAnnualFeeCalc.Style.Add("display", "");

                lblStdAnnFeeCalc.Visible = true;
                txtStdAnnualFeeCalc.Visible = true;
                //   RequiredFieldValidator7.Visible = true;
                RequiredFieldValidator8.Visible = false;
                if (strFeeType.ToUpper() == "Flat".ToUpper())//Check FeeType="Flat" if true Hide StdAnnualFeeCalc
                {
                    // trStdAnnualFeeCalc.Style.Add("display", "none");
                    lblStdAnnFeeCalc.Visible = false;
                    txtStdAnnualFeeCalc.Visible = false;
                }
                txtDiscount.Visible = true;
                lbldiscount.Visible = true;
                lbldecimal.Visible = true;
            }
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
    protected void txtTotalAUM_TextChanged(object sender, EventArgs e)
    {
        string value = txtTotalAUM.Text.Replace(",", "").Replace("$", "");
        decimal ul;

        if (decimal.TryParse(value, out ul))
        {
            txtTotalAUM.TextChanged -= txtTotalAUM_TextChanged;//event handling 
            /* used Globalization To Format totalAum in (en-US) with 2 decimal */
            txtTotalAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
            txtTotalAUM.TextChanged += txtTotalAUM_TextChanged;
            txtBillingAUM.Focus();//set Focus To BiilingAum
        }
    }

    protected void txtBillingAUM_TextChanged(object sender, EventArgs e)
    {
        /* Replace , to $ for BillingAum*/
        string value = txtBillingAUM.Text.Replace(",", "").Replace("$", "");


        if (txtBillingAUM.Text.Trim() != "")
        {
            DB clsDB = new DB();//class Library
            /* Return AnnualFee  into dataset where Billing Amount and clientType Selected */
            DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@TotalAUMNmb='" + txtTotalAUM.Text.Replace(",", "").Replace("$", "") + "' ");
            if (dsBilling.Tables[0].Rows.Count > 0)
            {
                /* check AnnualFee if Null set AnnualFee="0.0" else display AnnualFee */
                string AnnualFee = Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]);

                /* used Globalization To Format StdAnnualFeeCalc in (en-US) with 2 decimal */
                txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));

                ViewState["BillingId"] = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingId"]);
            }

            CalulateFeeRate();//function to calculateFeeRate
            decimal ul;

            /*->out keyword used to display billingAUM
            ->TryParse used to convert into Decimal*/
            if (decimal.TryParse(value, out ul))
            {
                txtBillingAUM.TextChanged -= txtBillingAUM_TextChanged;
                txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                txtBillingAUM.TextChanged += txtBillingAUM_TextChanged;
            }
        }
        else
        {

            /* clear Following Control*/
            txtCustFeeAmount.Text = "";
            txtFeeRateCalc.Text = "";
            txtFeesPerMonth.Text = "";
            txtQuaterlyFeeCalc.Text = "";
            txtStdAnnualFeeCalc.Text = "";
        }

        if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
            txtRelationshipFee.Focus();
        else
            txtCustFeeAmount.Focus();
    }
    protected void txtCustFeeAmount_TextChanged(object sender, EventArgs e)
    {
        if (txtCustFeeAmount.Text != "" && txtDiscount.Text == "")
        {
            if (ddlClientType.SelectedValue != "100000001") //Standard with Relationship Fees
            {
                ddlClientType.SelectedValue = "100000007"; //Custom
                txtAdjAmt.Text = "";
                #region Standard Min-Max Changes- 3_25_2019
                lblbpsfees.Visible = false;
                lblMinValu.Visible = false;
                lblMaxVal.Visible = false;

                txtbpsfee.Visible = false;
                txtMinVal.Visible = false;
                txtMaxVal.Visible = false;

                txtbpsfee.Text = "";
                txtMinVal.Text = "";
                txtMaxVal.Text = "";
                btnStandardCalculate.Visible = false;
                ViewState["MaxValue"] = "";
                ViewState["MinValue"] = "";
                ViewState["bpsfeeValue"] = "";


                #endregion
            }
        }
        string value = txtCustFeeAmount.Text.Replace(",", "").Replace("$", "");
        decimal ul;

        if (decimal.TryParse(value, out ul))
        {
            txtCustFeeAmount.TextChanged -= txtCustFeeAmount_TextChanged;
            txtCustFeeAmount.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
            txtCustFeeAmount.TextChanged += txtCustFeeAmount_TextChanged;
        }
        CalulateFeeRate();
        btnSubmit.Focus();

    }
    protected void txtFeeRateCalc_TextChanged(object sender, EventArgs e)
    {
        /* check FeeRateCalc not Equals To Null*/
        if (txtFeeRateCalc.Text != "")
        {
            double doubleValue;
            /* Convert FeeRateCalc into double*/
            if (double.TryParse(txtFeeRateCalc.Text, out doubleValue))
            {
                doubleValue = doubleValue / 100;
                txtFeeRateCalc.Text = doubleValue.ToString("0.00%");//Convert doublevalue into String and Display in Percentage 
            }
        }
    }

    /*Calculate FeeRate*/
    private void CalulateFeeRate()
    {

        checkbillfortype();

        /* Replace ',' to '$' to Display String in $ Format*/
        string strAnnualFeeCalc = Convert.ToString(txtStdAnnualFeeCalc.Text.Replace(",", "").Replace("$", ""));
        // string stBillingAum = Convert.ToString(txtBillingAUM.Text.Replace(",", "").Replace("$", ""));
       // string stBillingAum = Convert.ToString(txtTotalBillingAUM.Text.Replace(",", "").Replace("$", ""));//added  7_18_2019

        string stBillingAum = Convert.ToString(txtBillAUM.Text.Replace(",", "").Replace("$", ""));//added  7_18_2019
        string stCustomFeeAmt = Convert.ToString(txtCustFeeAmount.Text.Replace(",", "").Replace("$", ""));
        string stFeesPerMonth = Convert.ToString(txtFeesPerMonth.Text.Replace(",", "").Replace("$", ""));
        string strRelationshipFee = Convert.ToString(txtRelationshipFee.Text.Replace(",", "").Replace("$", ""));
        string strDiscount = Convert.ToString(txtDiscount.Text.Replace(",", "").Replace("$", "").Replace("%", ""));

        //string replationfeee = txtBox1.Text.Replace(",", "").Replace("$", "");
        //string Setupfee = txtBox2.Text.Replace(",", "").Replace("$", "");
        //string Otherfee = txtBox3.Text.Replace(",", "").Replace("$", "");
        string FeeAmt = txtSecurityFee.Text.Replace(",", "").Replace("$", "");

        string TotalAnnualFee = Convert.ToString(txtDiscount.Text.Replace(",", "").Replace("$", "").Replace("%", ""));

        string TotalFlatFee = "0";
        if (ViewState["TotalFlatFee"] != null)
            TotalFlatFee = ViewState["TotalFlatFee"].ToString();

        /* Initialize following value = 0.0*/
        double dblAnnualFeeCalc = 0.0;
        double dblBillingAum = 0.0;
        double dblFeeRate = 0.0;
        double dblCustomFeeAmt = 0.0;
        double dblQuarterlyFee = 0.0;
        double dblFeePerMonth = 0.0;
        double dblRelationshipFee = 0.0;
        decimal dblQuaterlyFeeCalc = 0.0M;
        decimal dblAdjAmt = 0.0M;
        decimal dblAdjQtrFee = 0.0M;
        double dblDiscount = 0.0;
        double relationFee = 0.0;
        double SetupFee = 0.0;
        double OtherFee = 0.0;
        double feePctValue = 0;

        double dbTotalAnnualFeeCal = 0.0;
        double TotalFaltFees = 0.0;

        if (!string.IsNullOrEmpty(strAnnualFeeCalc))  //AnnualFee
            dblAnnualFeeCalc = Convert.ToDouble(strAnnualFeeCalc);

        if (!string.IsNullOrEmpty(stBillingAum)) //Billing Amount
            dblBillingAum = Convert.ToDouble(stBillingAum);

        if (!string.IsNullOrEmpty(stCustomFeeAmt)) //Custom Fee Amount
            dblCustomFeeAmt = Convert.ToDouble(stCustomFeeAmt);

        if (!string.IsNullOrEmpty(strRelationshipFee)) //Relationship Fee                       
            dblRelationshipFee = Convert.ToDouble(strRelationshipFee);

        if (!string.IsNullOrEmpty(strDiscount)) //Discount                      
            dblDiscount = Convert.ToDouble(strDiscount);

        //if (!string.IsNullOrEmpty(replationfeee))
        //    relationFee = Convert.ToDouble(replationfeee);

        //if (!string.IsNullOrEmpty(Setupfee))
        //    SetupFee = Convert.ToDouble(Setupfee);

        //if (!string.IsNullOrEmpty(Otherfee))
        //    OtherFee = Convert.ToDouble(Otherfee);

        if (!string.IsNullOrEmpty(TotalFlatFee))
            TotalFaltFees = Convert.ToDouble(TotalFlatFee);


        if (!string.IsNullOrEmpty(FeeAmt))
            feePctValue = Convert.ToDouble(FeeAmt);

        //if (!string.IsNullOrEmpty(TotalAnnualFee)) //Discount                      
        //    dbTotalAnnualFeeCal = Convert.ToDouble(TotalAnnualFee) + relationFee + SetupFee + feePctValue;
        //else
        //  dbTotalAnnualFeeCal = dblAnnualFeeCalc + relationFee + SetupFee +OtherFee+ feePctValue;
        dbTotalAnnualFeeCal = dblAnnualFeeCalc + TotalFaltFees + feePctValue; ; //relationFee + SetupFee + OtherFee + feePctValue;


        if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
        {
            dblCustomFeeAmt = dblAnnualFeeCalc + dblRelationshipFee;
            // dblCustomFeeAmt = dbTotalAnnualFeeCal + dblRelationshipFee;
            txtCustFeeAmount.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblCustomFeeAmt);
        }
        else
        {
            if (dblAnnualFeeCalc != 0.0 && dblDiscount != 0.0)
            {
                dblDiscount = dblDiscount / 100;
                dblCustomFeeAmt = dblAnnualFeeCalc - (dblDiscount * dblAnnualFeeCalc);
                // dblCustomFeeAmt = dbTotalAnnualFeeCal - (dblDiscount * dblAnnualFeeCalc);
            }
        }



        if (dblCustomFeeAmt != 0.0 || stCustomFeeAmt != "")
        {
            /* Calculating FeeRate in Percentage by Dividing (dblCustomFeeAmt / dblBillingAum) */
            /*  CalCulate QuaterlyFee divide CustomFeeAmt by 4*/
            /*  CalCulate FeePerMonth divide QuaterlyFee by 3*/
            dbTotalAnnualFeeCal = dblCustomFeeAmt + TotalFaltFees + feePctValue; ;// relationFee + SetupFee + OtherFee + feePctValue; ;
            if (dblBillingAum != 0)
                dblFeeRate = ((dbTotalAnnualFeeCal) / dblBillingAum) * 100;
            else
                dblFeeRate = 0;

            dblQuarterlyFee = (dbTotalAnnualFeeCal) / 4;
            dblFeePerMonth = (dblQuarterlyFee) / 3;

            //dblQuarterlyFee = dblCustomFeeAmt / 4;
            //dblFeePerMonth = dblQuarterlyFee / 3;
        }
        else
        {
            dbTotalAnnualFeeCal = dblAnnualFeeCalc + TotalFaltFees + feePctValue; ;// relationFee + SetupFee + OtherFee + feePctValue;



           // dbTotalAnnualFeeCal = dblAnnualFeeCalc;// relationFee + SetupFee + OtherFee + feePctValue;
            if (dblBillingAum != 0)
                //dblFeeRate = (dbTotalAnnualFeeCal / dblBillingAum) * 100;
                dblFeeRate = (dblAnnualFeeCalc / dblBillingAum) * 100;
            else
                dblFeeRate = 0;

            if (dbTotalAnnualFeeCal != 0)
            {
                dblQuarterlyFee = dbTotalAnnualFeeCal / 4;
                dblFeePerMonth = dbTotalAnnualFeeCal / 3;
            }
            else
            {

            }
        }


        if (dblFeeRate.ToString() != "Infinity" && dblFeeRate != 0)
        {
            txtFeeRateCalc.Text = dblFeeRate.ToString();
        }
        else
        {
            txtFeeRateCalc.Text = "NA";
        }
        //  txtQuaterlyFeeCalc.Text = dblQuarterlyFee.ToString();
        // txtFeesPerMonth.Text = dblFeePerMonth.ToString();

        if (dblCustomFeeAmt != 0.0)
            txtCustFeeAmount.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblCustomFeeAmt);
        //else
        //    txtCustFeeAmount.Text = "";

        txtQuaterlyFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblQuarterlyFee);
        txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblFeePerMonth);

        if (dbTotalAnnualFeeCal != 0.0)
            txtTotalAnnualFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dbTotalAnnualFeeCal);
        else
            txtTotalAnnualFee.Text = "";



        if (txtFeeRateCalc.Text != "") //Fee Rate Calc
        {
            double doubleValue;
            if (double.TryParse(txtFeeRateCalc.Text, out doubleValue))//converts value into double(TryParse)
            {
                doubleValue = doubleValue / 100;
                txtFeeRateCalc.Text = doubleValue.ToString("0.00%");
            }
        }

        string strQuaterlyFeeCalc = Convert.ToString(txtQuaterlyFeeCalc.Text.Replace(",", "").Replace("$", ""));
        string strAdjAmt = Convert.ToString(txtAdjAmt.Text.Replace(",", "").Replace("$", ""));

        if (!string.IsNullOrEmpty(strQuaterlyFeeCalc) && strQuaterlyFeeCalc != "0.0")
            dblQuaterlyFeeCalc = Convert.ToDecimal(strQuaterlyFeeCalc);

        if (!string.IsNullOrEmpty(strAdjAmt))
        {
            decimal dec = Decimal.Parse(strAdjAmt, System.Globalization.NumberStyles.Currency);
            // dblAdjAmt = double.Parse(dec);
            dblAdjQtrFee = dblQuaterlyFeeCalc + dec;
            txtAdjQtrFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblAdjQtrFee);

        }

        string strAdjQtrFee = Convert.ToString(txtAdjQtrFee.Text.Replace(",", "").Replace("$", ""));

        /* Check AdjQtrFee not equal To 0.0m*/
        if (dblAdjQtrFee != 0.0m)

            /* to calculate FeePerMonth ----> Divide AdjQtrFee by 3 and Convert into Double*/
            dblFeePerMonth = Convert.ToDouble(dblAdjQtrFee) / 3;
        else

            /* to calculate FeePerMonth ----> Divide QuarterlyFee by 3 and Convert into Double*/
            dblFeePerMonth = Convert.ToDouble(dblQuarterlyFee) / 3;

        txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblFeePerMonth);

        //if (txtBox1.Text.Replace("$", "").Replace(",", "") != "")
        //{
        //    decimal val = Convert.ToDecimal(txtBox1.Text.Replace("$", "").Replace(",", ""));
        //    txtBox1.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", val);
        //}

        //if (txtBox2.Text.Replace("$", "").Replace(",", "") != "")
        //{
        //    decimal val = Convert.ToDecimal(txtBox2.Text.Replace("$", "").Replace(",", ""));
        //    txtBox2.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", val);
        //}
        //if (txtBox3.Text.Replace("$", "").Replace(",", "") != "")
        //{
        //    decimal val = Convert.ToDecimal(txtBox3.Text.Replace("$", "").Replace(",", ""));
        //    txtBox3.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", val);
        //}
        //  LinkButton1.Visible = true;
    }
    #endregion

    protected void ddlAdvisor_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindHouseHold();
        BindBillingName();
        ClearControls();
    }
    protected void ddlHH_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindBillingName();
        ClearControls();
        CheckandGetExistingData();
    }
    protected void txtStdAnnualFeeCalc_TextChanged(object sender, EventArgs e)
    {
        CalulateFeeRate();
    }
    protected void ddlBillFor_SelectedIndexChanged(object sender, EventArgs e)
    {
        checkbillfortype();
        ClearControls();
    }
    protected void btnGeneratePDF_Click(object sender, EventArgs e)
    {

    }

    private void CheckandGetExistingData()
    {
        try
        {
            if (ddlHH.SelectedValue != "" && txtAUMDate.Text != "" && ddlBillFor.SelectedValue != "")
            {
                DB clsDB = new DB();//class Library 

                /* Check ddlHH SelectedValue To Format if true select null else SelectedValue*/
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

                /* Return BillingInvoice into dataset where HouseHold value AumAsodfDate and BillFor Selected */
                //DataSet dsBillingChk = clsDB.getDataSet("SP_S_BILLINGINVOICE_CHECK @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");
                //    string SQLQuery = "SP_S_BILLINGINVOICE_CHECK_test @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ";
                //DataSet dsBillingChk = clsDB.getDataSet("SP_S_BILLINGINVOICE_CHECK_test @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");

                string SQLQuery = "SP_S_BILLINGINVOICE_CHECK @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ";
                DataSet dsBillingChk = clsDB.getDataSet(SQLQuery);
                ViewState["CheckDS"] = dsBillingChk;
                #region checks existing records first
                if (dsBillingChk.Tables[0].Rows.Count > 0)
                {


                    /*retrieve following values and initialize to appropriate variables*/

                    string strAUMAsofDate = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_AUMAsofDate"]);
                    string strtotalaum = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_totalaum"]);
                    string straum = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_aum"]);
                    string strannualfee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_annualfee"]);
                    string strfeerate = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_feerate"]);
                    string strquarterlyfee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_quarterlyfee"]);
                    string strmonth1fee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_month1fee"]);
                    string strmonth2fee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_month2fee"]);
                    string strmonth3fee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_month3fee"]);
                    // string strFeeScheduleId = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_FeeScheduleId"]);
                    string strFeeScheduleType = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_FeeScheduleType"]);
                    string strAdjustment = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_Adjustment"]);
                    string strAdjustmentReason = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_AdjustmentReason"]);
                    string strAdjustedFee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_AdjustedFee"]);
                    //string strRelationshipFee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_RelationshipFee"]);
                    string strCustomFee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_CustomFee"]);
                    string strDiscount = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_Discount"]);
                    string Accrued = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_Accrued"]);

                    string ssi_minimumfeein = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_minimumfeein"]);
                    string ssi_maximumfeeasa = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_maximumfeeasa"]);
                    string ssi_feeonfirst25mminbps = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_feeonfirst25mminbps"]);

                    //added 7_15_2019
                    string ScheduleAUM = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ScheduleAUM"]);

                    // Maximum fee as a %- added 3_25_2019 Standard Min-Max Changes
                    if (ssi_maximumfeeasa != "")
                    {
                        txtMaxVal.Text = ssi_maximumfeeasa.ToString();
                        if (txtMaxVal.Text != "")
                        {
                            double doubleValue;
                            if (double.TryParse(txtMaxVal.Text, out doubleValue))
                            {
                                /*Convert doubleValue to percentage*/
                                doubleValue = doubleValue / 100;
                                txtMaxVal.Text = doubleValue.ToString("0.00%");
                            }
                        }
                    }
                    //Minimum fee in $ - added 3_25_2019 Standard Min-Max Changes
                    if (ssi_minimumfeein != "")
                    {
                        txtMinVal.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(ssi_minimumfeein));
                    }
                    // Fee on first $25MM in bps- added 3_25_2019 Standard Min-Max Changes
                    if (ssi_feeonfirst25mminbps != "")
                    {
                        txtbpsfee.Text = ssi_feeonfirst25mminbps.ToString();
                        if (txtbpsfee.Text != "")
                        {
                            double doubleValue;
                            if (double.TryParse(txtbpsfee.Text, out doubleValue))
                            {
                                txtbpsfee.Text = doubleValue.ToString("0.00");
                            }
                        }
                    }


                    //Aum as of Date
                    /* convert to date format 
                      convert to short date
                     display AumasofDate in AumDate Textbox */
                    if (strAUMAsofDate != "")
                    {
                        DateTime dt = Convert.ToDateTime(strAUMAsofDate);
                        strAUMAsofDate = dt.ToShortDateString();
                        txtAUMDate.Text = strAUMAsofDate.ToString();
                    }
                    //total amount 
                    /* used Globalization To Format totalAum in (en-US) with 2 decimal */
                    if (strtotalaum != "")
                        txtTotalAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strtotalaum));

                    //Aum
                    /* used Globalization To Format aum in (en-US) with 2 decimal */
                    if (straum != "")
                    {
                        txtTotalBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(straum));
                        txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(straum));
                        txtBillingAUM.Visible = false;



                    }

                    if (ScheduleAUM != "")
                    {
                        txtBillAUM.Enabled = false; //added 7_15_2019
                        txtBillAUM.Visible = true; //added 7_15_2019
                        lblStandardFeeAssets.Visible = true;//added 7_15_2019
                        txtBillAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(ScheduleAUM)); //added 7_15_2019
                        txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(ScheduleAUM));		//added 7_15_2019
                        txtBillingAUM.Visible = false;																																		  //added 7_15_2019
                    }

                    //Annual Fee
                    /* used Globalization To Format annualfee in (en-US) with 2 decimal */
                    if (strannualfee != "")
                        txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strannualfee));

                    //Discount
                    if (strDiscount != "")
                        txtDiscount.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:P2}", Convert.ToDouble(strDiscount));

                    //FeeRate
                    /*Convert FeeRateCalc to string
                    ->out keyword used to display billingAUM
                    ->TryParse used to convert into Decimal
                    */
                    if (strfeerate != "")
                    {
                        txtFeeRateCalc.Text = strfeerate.ToString();
                        if (txtFeeRateCalc.Text != "")
                        {
                            double doubleValue;
                            if (double.TryParse(txtFeeRateCalc.Text, out doubleValue))
                            {
                                /*Convert doubleValue to percentage*/
                                doubleValue = doubleValue / 100;
                                txtFeeRateCalc.Text = doubleValue.ToString("0.00%");
                            }
                        }
                    }

                    //Quarterly Fee
                    /*check quarterlyfee not equals to null*/
                    /* used Globalization To Format quarterlyfee in (en-US) with 2 decimal */

                    if (strquarterlyfee != "")
                        txtQuaterlyFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strquarterlyfee));

                    //1 Month Fee
                    /*check month1fee not equals to null*/
                    /* used Globalization To Format month1fee in (en-US) with 2 decimal */
                    if (strmonth1fee != "")
                        txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strmonth1fee));

                    //2 Month Fee
                    /*check month2fee not equals to null*/
                    /* used Globalization To Format month2fee in (en-US) with 2 decimal */
                    if (strmonth2fee != "")
                        txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strmonth2fee));

                    //3 Month Fee
                    /*check month3fee not equals to null*/
                    /* used Globalization To Format month3fee in (en-US) with 2 decimal */
                    if (strmonth3fee != "")
                        txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strmonth3fee));

                    //Relationship Fee
                    /*check RelationshipFee not equals to null*/
                    /* used Globalization To Format RelationshipFee in (en-US) with 2 decimal */
                    //if (strRelationshipFee != "")
                    //    txtRelationshipFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strRelationshipFee));
                    //else
                    //    txtRelationshipFee.Text = "";

                    //Customfee
                    /*check Customfee not equals to null*/
                    /* used Globalization To Format RelationshipFee in (en-US) with 2 decimal */
                    if (strCustomFee != "")
                        txtCustFeeAmount.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strCustomFee));
                    else
                        txtCustFeeAmount.Text = "";

                    //Adjustment 
                    /*check Adjustment not equals to null*/
                    /* used Globalization To Format RelationshipFee in (en-US) with 2 decimal */

                    if (strAdjustment != "")
                    {
                        txtAdjAmt.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strAdjustment));
                        //  this.Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "showAdjustmentSection();", true);
                    }
                    //else
                    //    this.Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "hideAdjustmentSection();", true);


                    //Adjusted Fees
                    /*check Adjusted Fees not equals to null*/
                    /* used Globalization To Format RelationshipFee in (en-US) with 2 decimal */
                    if (strAdjustedFee != "")
                    {
                        txtAdjQtrFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strAdjustedFee));
                        double dblAdjQtrFee = 0.0;
                        double dblFeePerMonth = 0.0;
                        dblAdjQtrFee = Convert.ToDouble(strAdjustedFee);
                        dblFeePerMonth = dblAdjQtrFee / 3;
                        txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblFeePerMonth);
                    }
                    else
                    {
                        double dblQuarterlyFee = 0.0;
                        double dblFeePerMonth = 0.0;
                        dblQuarterlyFee = Convert.ToDouble(strquarterlyfee);
                        dblFeePerMonth = dblQuarterlyFee / 3;
                        txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblFeePerMonth);

                    }
                    //Adjustment Reason
                    /*check Adjustment Reason not equals to null
                     Remove Readonly from CustomFeeAmount */

                    if (strAdjustmentReason != "")
                        txtAdjReason.Text = strAdjustmentReason.ToString();
                    txtCustFeeAmount.ReadOnly = false;
                    txtCustFeeAmount.BackColor = System.Drawing.Color.White;

                    if (strFeeScheduleType != "")
                    {
                        /*Remove items from ClientType DropDown where value =0 */
                        System.Web.UI.WebControls.ListItem itemToRemove = ddlClientType.Items.FindByValue("0");
                        if (itemToRemove != null)
                        {
                            ddlClientType.Items.Remove(itemToRemove);
                        }

                        BindClientType();
                        ddlClientType.SelectedValue = strFeeScheduleType;
                        if (strFeeScheduleType == "100000007") //Custom
                        {

                            /* for FeeType=Custom 
                              Hide AnnualFeeCalc
                              RequiredValidator7 
                              RequiredValidator8 
                             */
                            txtStdAnnualFeeCalc.Text = "";
                            // trStdAnnualFeeCalc.Style.Add("display", "none");
                            lblStdAnnFeeCalc.Visible = false;
                            txtStdAnnualFeeCalc.Visible = false;
                            // txtCustFeeAmount.Focus();
                            //  RequiredFieldValidator7.Visible = false;
                            RequiredFieldValidator8.Visible = true;
                            if (strannualfee != "")
                                txtCustFeeAmount.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strannualfee));
                        }
                        else if (strFeeScheduleType == "100000001") //Standard with Relationship Fees
                        {
                            /* for FeeType=Standard with Relationship Fees
                              Set CustFeeAmount to Readonly
                              Display StdAnnualFeeCalc
                              RequiredValidator7 
                              RequiredValidator8 
                              RelationshipFee
                             */
                            txtCustFeeAmount.ReadOnly = true;
                            txtCustFeeAmount.BackColor = System.Drawing.Color.LightGray;
                            // trStdAnnualFeeCalc.Style.Add("display", "");
                            lblStdAnnFeeCalc.Visible = true;
                            txtStdAnnualFeeCalc.Visible = true;
                            // RequiredFieldValidator7.Visible = true;
                            //    RequiredFieldValidator11.Visible = true;
                            trRelationshipFee.Style.Add("display", "");
                            txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strannualfee));

                        }
                        else
                        {
                            /* for FeeType=other than Custom and Standard with Relationship Fees
                              Display StdAnnualFeeCalc 
                              RequiredValidator7 
                              
                              Hide
                              RequiredValidator8 
                              RelationshipFee
                             */
                            #region Standard Min-Max Changes - 3_25_2019
                            if (ddlClientType.SelectedValue == "100000009") //Standard Fee or $180K(greater of the two)// if (ddlClientType.SelectedValue == "100000000")
                            {
                                lblbpsfees.Visible = true;
                                lblMinValu.Visible = true;
                                lblMaxVal.Visible = true;

                                txtbpsfee.Visible = true;
                                txtMinVal.Visible = true;
                                txtMaxVal.Visible = true;
                                btnStandardCalculate.Visible = true;
                            }
                            else
                            {
                                lblbpsfees.Visible = false;
                                lblMinValu.Visible = false;
                                lblMaxVal.Visible = false;

                                txtbpsfee.Visible = false;
                                txtMinVal.Visible = false;
                                txtMaxVal.Visible = false;

                                //txtbpsfee.Text = "";
                                //txtMinVal.Text = "";
                                //txtMaxVal.Text = "";
                                btnStandardCalculate.Visible = false;
                            }
                            #endregion

                            trRelationshipFee.Style.Add("display", "none");
                            // trStdAnnualFeeCalc.Style.Add("display", "");
                            lblStdAnnFeeCalc.Visible = true;
                            txtStdAnnualFeeCalc.Visible = true;
                            //   RequiredFieldValidator7.Visible = true;
                            RequiredFieldValidator8.Visible = false;
                            if (strannualfee != "")
                                txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(strannualfee));
                        }

                        //string relationfee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_Relationshipfee"]);
                        //string setupfee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_Setupfee"]);
                        //string otherfee = Convert.ToString(dsBillingChk.Tables[0].Rows[0]["ssi_Otherfee"]);

                        //txtBox1.Style.Add("display", "none");
                        //lblbox1.Style.Add("display", "none");
                        //txtBox2.Style.Add("display", "none");
                        //lblbox2.Style.Add("display", "none");
                        //txtBox3.Style.Add("display", "none");
                        //lblbox3.Style.Add("display", "none");

                        //if (relationfee != "" && Convert.ToDecimal(relationfee) != 0)
                        //{
                        //    txtBox1.Style.Add("display", "");
                        //    lblbox1.Style.Add("display", "");
                        //    txtBox1.TextChanged -= txtBox1_TextChanged;
                        //    txtBox1.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(relationfee));
                        //    txtBox1.TextChanged += txtBox1_TextChanged;
                        //}

                        //if (setupfee != "" && Convert.ToDecimal(setupfee) != 0)
                        //{
                        //    Tr2.Style.Add("display", "");
                        //    txtBox2.Style.Add("display", "");
                        //    lblbox2.Style.Add("display", "");
                        //    txtBox2.TextChanged -= txtBox2_TextChanged;
                        //    txtBox2.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(setupfee));
                        //    txtBox2.TextChanged += txtBox2_TextChanged;
                        //}

                        //if (otherfee != "" && Convert.ToDecimal(otherfee) != 0)
                        //{
                        //    txtBox3.Style.Add("display", "");
                        //    lblbox3.Style.Add("display", "");
                        //    txtBox3.TextChanged -= txtBox3_TextChanged;
                        //    txtBox3.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(otherfee));
                        //    txtBox3.TextChanged += txtBox3_TextChanged;
                        //}



                        // CalulateFeeRate();
                    }
                    else
                    {
                        //Items insertion into Client Type DropDown
                        ddlClientType.Items.Insert(0, new System.Web.UI.WebControls.ListItem("--Select--", "0"));
                        ddlClientType.SelectedValue = "0";
                    }


                    if (dsBillingChk.Tables[1] != null)
                    {
                        if (dsBillingChk.Tables[1].Rows.Count > 1)
                        {
                            lbGroup.Visible = true;
                            BindListBox(dsBillingChk.Tables[1]);
                            if (dsBillingChk.Tables[1].Rows.Count > 1)
                            {
                                lbGroup.Visible = true;
                                BindListBox(dsBillingChk.Tables[1]);
                            }

                        }
                        else
                        {
                            lbGroup.Visible = false;
                        }
                        for (int i = 0; i < dsBillingChk.Tables[1].Rows.Count; i++)
                        {
                            string data = dsBillingChk.Tables[1].Rows[i]["ssi_billingId"].ToString();
                            // SourceFileArray[i] = data;
                            ls.Add(data);
                        }
                        ViewState["sourceLsit"] = ls;
                        ViewState["BillingForList"] = dsBillingChk.Tables[1];

                        ViewState["Billing"] = dsBillingChk.Tables[2];

                    }
                    else
                    {
                        lbGroup.Visible = false;
                    }

                    if (dsBillingChk.Tables[3] != null)
                    {
                        ViewState["CustomeFeeDelete"] = dsBillingChk.Tables[3];
                    }

                    if (dsBillingChk.Tables[4] != null)
                    {
                        ViewState["CustomeFeeInsert"] = dsBillingChk.Tables[4];
                    }

                    btnDeleteAndCal.Visible = true;

                    if (dsBillingChk.Tables[5] != null)
                    {
                        txtNotes.Text = dsBillingChk.Tables[5].Rows[0]["AdvisoryNotes"].ToString();
                    }
                    ViewState["FamilyNotes"] = dsBillingChk.Tables[5];
                    if (dsBillingChk.Tables[6] != null)
                    {
                        FeeAmount = Convert.ToString(dsBillingChk.Tables[6].Rows[0]["CustomFee"]) == "" ? "0.0" : Convert.ToString(dsBillingChk.Tables[6].Rows[0]["CustomFee"]);
                        if (Convert.ToDecimal(FeeAmount) != 0)
                        {
                            txtSecurityFee.Visible = true;
                            Label5.Visible = true;
                            txtSecurityFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(FeeAmount));
                        }
                        else
                        {
                            txtSecurityFee.Visible = false;
                            Label5.Visible = false;
                        }
                    }
                    else
                    {
                        FeeAmount = "0";
                        txtSecurityFee.Visible = false;
                        Label5.Visible = false;
                    }

                    //added 7_15_2019
                    string FeeAmount1 = Convert.ToString(dsBillingChk.Tables[6].Rows[0]["CustomFeeAUM"]) == "" ? "0.0" : Convert.ToString(dsBillingChk.Tables[6].Rows[0]["CustomFeeAUM"]);
                    if (Convert.ToDecimal(FeeAmount1) != 0)
                    {
                        txtFeeAUM.Visible = true;
                        lblAssetsUnderAdministration.Visible = true;
                        txtFeeAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(FeeAmount1));
                    }
                    else
                    {
                        txtFeeAUM.Visible = false;
                        lblAssetsUnderAdministration.Visible = false;
                    }



                    if (dsBillingChk.Tables[8].Rows.Count != 0)
                    {

                        gridviewBind(dsBillingChk.Tables[8]);
                        string TotalFlatFee = dsBillingChk.Tables[8].Rows[0]["totFlatFee"].ToString();
                        ViewState["TotalFlatFee"] = TotalFlatFee;
                    }
                    if (Accrued == "True")
                        chkAccured.Checked = true;
                    else
                        chkAccured.Checked = false;
                    CalulateFeeRate();

                }
                #endregion
                else
                {
                    // get data for new record
                    NewData();
                    btnDeleteAndCal.Visible = false;

                    #region Standard Min-Max Changes - 3_25_2019
                    if (ddlClientType.SelectedValue == "100000009")//Standard Fee or $180K(greater of the two)//  if (ddlClientType.SelectedValue == "100000000")
                    {
                        lblbpsfees.Visible = true;
                        lblMinValu.Visible = true;
                        lblMaxVal.Visible = true;

                        txtbpsfee.Visible = true;
                        txtMinVal.Visible = true;
                        txtMaxVal.Visible = true;
                        btnStandardCalculate.Visible = true;
                    }
                    else
                    {
                        lblbpsfees.Visible = false;
                        lblMinValu.Visible = false;
                        lblMaxVal.Visible = false;

                        txtbpsfee.Visible = false;
                        txtMinVal.Visible = false;
                        txtMaxVal.Visible = false;

                        txtbpsfee.Text = "";
                        txtMinVal.Text = "";
                        txtMaxVal.Text = "";

                        btnStandardCalculate.Visible = false;
                    }
                    #endregion
                }
            }
        }
        catch (Exception ex)
        {
            // lblMessage.Text = ex.ToString();
        }
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        try
        {
            if (!Page.IsValid)
            {
                //Show Validation Summary
                ValidationSummary1.ShowMessageBox = true;
                return;
            }

            DB clsDB = new DB();
            object HHValue = ddlHH.SelectedValue == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlHH.SelectedValue + "'";
            //string sqlstr = "[SP_S_BILLING_NAME] @HouseHoldUUID = " + HHValue + " ";
            //DataSet ds1 = clsDB.getDataSet(sqlstr);
            //BillingUUID = Convert.ToString(ds1.Tables[0].Rows[0]["Ssi_billingid"]);

            string strBillFor = ddlBillFor.SelectedValue;
            char[] delimiterChars = { '|' };
            string[] words = strBillFor.ToString().Split(delimiterChars);
            int len = words.Length;
            if (len > 0)
            {
                strBillFor = words[0];
                BillingUUID = words[0];
            }

            DataTable dtBilling = (DataTable)ViewState["Billing"];
            DataSet ds = (DataSet)ViewState["dsData"];
            DataSet dsBillingChk = clsDB.getDataSet("SP_S_BILLINGINVOICE_CHECK @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");

            if (dsBillingChk.Tables[0].Rows.Count > 0)
            {
                ///  DataSet dsBillingChk = (DataSet)ViewState["CheckDS"];

                string ssi_relationshipfee = "0";// dsBillingChk.Tables[0].Rows[0]["ssi_relationshipfee"].ToString();
                string ssi_relationshipfeeFlg = "0";//   dsBillingChk.Tables[0].Rows[0]["ssi_relationshipfeeFlg"].ToString();

                string ssi_SetUpFee = "0";// dsBillingChk.Tables[0].Rows[0]["ssi_SetUpFee"].ToString();
                string ssi_SetUpFeeFlg = "0";// dsBillingChk.Tables[0].Rows[0]["ssi_SetUpFeeFlg"].ToString();

                string ssi_OtherFee = "0";//  dsBillingChk.Tables[0].Rows[0]["ssi_OtherFee"].ToString();
                string ssi_OtherFeeFlg = "0";// dsBillingChk.Tables[0].Rows[0]["ssi_OtherFeeFlg"].ToString();


                foreach (DataRow row in dsBillingChk.Tables[2].Rows)
                {

                    //string BillingUUID = row["ssi_BillingId"].ToString();
                    //string name = row["BillingName"].ToString();
                    //string BillingID = row["BillingID"].ToString();

                    string BillingInvoiceID = row["ssi_billinginvoiceid"].ToString();
                    string persent = row["ssi_allocationpercent"].ToString();
                    string vBillingUID = row["Ssi_billingid"].ToString();
                    string vAumValue = row["ssi_totalaum"].ToString();
                    string vFinalBillingValue = row["ssi_aum"].ToString();
                    //added 7_15_2019
                    string ssi_FeePctAmt = row["ssi_SecurityFeeAUM"].ToString();
                    string ssi_ScheduleAUM = row["ssi_ScheduleAUM"].ToString();

                    UpdateInvoiceNew(vBillingUID, BillingInvoiceID, persent, ssi_relationshipfee, ssi_relationshipfeeFlg, ssi_SetUpFee, ssi_SetUpFeeFlg, ssi_OtherFee, ssi_OtherFeeFlg, vAumValue, vFinalBillingValue, ssi_FeePctAmt, ssi_ScheduleAUM);
                }
                DataTable dtDataDelete = (DataTable)ViewState["CustomeFeeDelete"];
                //   DataTable dtDataInsert = (DataTable)ViewState["CustomeFeeInsert"];

                if (dtDataDelete != null)     // delete all old custome fee data
                {
                    if (dtDataDelete.Rows.Count > 0)
                    {
                        foreach (DataRow row in dtDataDelete.Rows)
                        {
                            string id = row["ssi_billingcustomfeeId"].ToString();
                            deleteCustomeFee(id);

                        }
                        //dtDataDelete = null;
                        ViewState["CustomeFeeDelete"] = null;
                    }
                }

                insertCustomefee();  // insert all custome fee data

                updateInvoicewithTotalandPer();


            }
            else
            {
                if (dtBilling.Rows.Count > 0)
                {
                    foreach (DataRow row in dtBilling.Rows)
                    {
                        string persent = row["BillingPct"].ToString();
                        string BillingUUID = row["ssi_BillingId"].ToString();
                        string name = row["BillingName"].ToString();
                        // string BillingID = row["BillingID"].ToString();
                        string BillingID = null;
                        string vBillingMarketValue = row["FinalBillingMarketValue"].ToString();

                        string ssi_relationshipfee = "0";// row["ssi_RelationshipFee"].ToString();
                        string ssi_relationshipfeeFlg = "0";// row["ssi_RelationshipFeeFlg"].ToString();
                        string ssi_setupfee = "0";// row["ssi_SetUpFee"].ToString();
                        string ssi_SetUpFeeFlg = "0";// row["ssi_SetUpFeeFlg"].ToString();

                        string ssi_OtherFee = "0";// row["ssi_OtherFee"].ToString();
                        string ssi_OtherFeeFlg = "0";// row["ssi_OtherFeeFlg"].ToString();

                        string FinalAUMMarketValue = row["FinalAUMMarketValue"].ToString();

                        //added 7_15_2019
                        string ssi_FeePctAmt = row["ssi_FeePctAmt"].ToString();
                        string ssi_ScheduleAUM = row["ssi_ScheduleAUM"].ToString();
                        CreateInvoiceNew(BillingUUID, BillingID, persent, ssi_relationshipfee, ssi_relationshipfeeFlg, ssi_setupfee, ssi_SetUpFeeFlg, ssi_OtherFee, ssi_OtherFeeFlg, FinalAUMMarketValue, vBillingMarketValue, ssi_FeePctAmt, ssi_ScheduleAUM);

                    }


                    InsertFlatFee(); // insert all flat fee 

                    updateInvoicewithTotalandPer();   // after insert, updae


                }
            }

            UpdateNotes();   // update notes in family entity

            PDFBillingWorksheetAndInvoice();

            ClearControls();
            BindClientType();
            CheckandGetExistingData();

            #region SEC PROCESS
            try
            {
                //Call Excel PRocedure for batch
                DataSet ds2 = new DataSet();
                ds2 = ExcelProcedure(txtAUMDate.Text.ToString(), ddlBillFor.SelectedValue.ToString());
                if (ds2.Tables[1].Rows.Count > 0)
                {
                    //Generate Excel   
                    string filepath = GenerateExcel(ds2);
                    string fileName = Path.GetFileName(filepath);

                    //Upload File To  Sharepoint 
                    string FolderPath = UploadToLib(filepath, txtAUMDate.Text.ToString());
                    if (FolderPath != "")
                    {
                        //Update Feilds in BILLING INVOICE ENTITY
                        DataTable dtInvoiceId = InvoiceId();
                        if (dtInvoiceId.Rows.Count > 0)
                        {
                            foreach (DataRow row in dtInvoiceId.Rows)
                            {
                                string BillingInvoiceID = row["ssi_billinginvoiceid"].ToString();
                                updateFields(BillingInvoiceID, FolderPath, fileName);
                            }

                        }
                        //Delete File form local folder
                        System.IO.File.Delete(filepath);
                    }
                }
            }
            catch (Exception ex)
            {
                lblMessage.Text = "File Upload to sharepoint(Billing) Failed";
            }
            #endregion
        }
        catch
        {
            lblMessage.Text = "Records Update Fail";
        }
        ViewState["MaxValue"] = "";
        ViewState["MinValue"] = "";
        ViewState["bpsfeeValue"] = "";

    }

    #region Sharepoint

    public bool CheckFolderPathExists(String folderPath)
    {
        string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        string filename = @"E:\devlopment\GP\SharepointCode\DemoTest.txt";
        ClientContext context = new ClientContext(siteUrl);


        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword);
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);


        Web site = context.Web;
        try
        {
            //Get the required RootFolder
            //string barRootFolderRelativeUrl = "Shared Documents/test 2/";
            //  string barRootFolderRelativeUrl = folderPath;

            //folderPath = folderPath.Replace("\\", "/");
            //folderPath = "Documents/" + folderPath;
            //Folder barFolder = site.GetFolderByServerRelativeUrl(folderPath);

            //int len = folderPath.Length;
            //int indexlen = folderPath.IndexOf("Documents");
            //indexlen = indexlen + 10;
            //int cnt = len - indexlen;
            //string vNewSharePointReportFolder = folderPath.Substring(indexlen, cnt);
            //vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/").Replace(@"\", "/");




            //vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/");
            //vNewSharePointReportFolder = "Documents/" + vNewSharePointReportFolder;
            Folder barFolder = site.GetFolderByServerRelativeUrl(folderPath);

            // context.Load(barFolder);
            context.ExecuteQuery();

            return true;
        }
        catch (Exception ex)
        {

            return false;
        }

        //  return true;
    }

    public bool CopyFilenew(string Ssi_SharePointReportFolder, string destFilename, string vSourcrFile)  // string vSourcefile, string vDestinationFile
    {

        #region not used
        //string Ssi_ClientPortalFolder = @"\\sp02\\Client%20Portal\Documents\Scalise%20Test\Gresham%20Statements\2016";

        // Ssi_SharePointReportFolder = @"\\sp02\ClientServ\Documents\Test%20JMASA\";

        //string Filename = "test_Masa T 2016-0630.pdf";

        //string filename = @"E:\devlopment\GP\SharepointCode\DemoTest.txt";

        //Ssi_SharePointReportFolder = Ssi_SharePointReportFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
        ////  Ssi_ClientPortalFolder = Ssi_ClientPortalFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();

        //int len = Ssi_SharePointReportFolder.Length;
        //int indexlen = Ssi_SharePointReportFolder.IndexOf("Documents");
        //indexlen = indexlen + 10;
        //int cnt = len - indexlen;
        //string vNewSharePointReportFolder = Ssi_SharePointReportFolder.Substring(indexlen, cnt);
        //vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/").Replace(@"\", "/");

        //len = Ssi_ClientPortalFolder.Length;
        //indexlen = Ssi_ClientPortalFolder.IndexOf("Documents");
        //indexlen = indexlen + 10;
        //cnt = len - indexlen;
        //string vNewClientFolderPath = Ssi_ClientPortalFolder.Substring(indexlen, cnt);

        //  string FilePath = vNewSharePointReportFolder + Filename;

        // FilePath = FilePath.Replace("\\", "/");

        //   vNewSharePointReportFolder = "Documents/" + vNewSharePointReportFolder;
        //  Response.Write(vNewSharePointReportFolder);
        #endregion

        try
        {
            // string vNewSharePointReportFolder = sharepointFolderPath(Ssi_SharePointReportFolder);
            // vSourcrFile = @"E:\AdventReport\BatchReport\ExcelTemplate\TEST COVER LETTER.PDF";
            string destFilename1 = "";
            string url = Request.Url.AbsoluteUri;
            //if (url.Contains("localhost") || url.Contains("crm-test3") || url.Contains("gp-crm2016") || url.Contains("testcrm.greshampartners")) // added for UPGRADE2016 (9_8_2017) // added testcrm.greshampartners(IFD) - 12_18_2018
            //{
            //    destFilename1 = destFilename + "_test.pdf";
            //}
            //else if (url.Contains("grpao1-vwcrm") || url.Contains("gp-crm1") || url.Contains("crm.greshampartners"))// added crm.greshampartners(IFD) - 12_18_2018
            //{
            //    destFilename1 = destFilename + ".pdf";
            //}

            //changed 7_17_2019 - Server added to Params.config
            if (Server.ToLower() == "prod")
            {
                destFilename1 = destFilename + ".pdf";
            }
            else
            {
                destFilename1 = destFilename + "_test.pdf";
            }

            string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);
            Web site = context.Web;

            byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
            System.IO.Stream stream = new System.IO.MemoryStream(bytes);

            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(Ssi_SharePointReportFolder);
            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(destFilename1), Overwrite = true };
            currentRunFolder.Files.Add(newFile);

            currentRunFolder.Update();

            context.ExecuteQuery();
            return true;
            //Response.Write("FileUpload");
        }
        catch (Exception ex)
        {
            Response.Write(ex.Message);
            return false;

        }//  bool result= CheckFolderPathExis(vNewSharePointReportFolder);



    }

    public bool checkSharepouintFileExist(string FilePath, string filename)
    {

        string filePath = "/clientserv/" + FilePath + "/" + filename;


        string siteUrl = "https://greshampartners.sharepoint.com";
        ClientContext clientContext = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //clientContext.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword);
        foreach (var c in Pass) passWord.AppendChar(c);
        clientContext.Credentials = new SharePointOnlineCredentials(user, passWord);
        // Web site = context.Web;

        Web web = clientContext.Web;
        Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filePath);
        bool bExists = false;
        try
        {
            clientContext.Load(file);
            clientContext.ExecuteQuery();
            bExists = file.Exists;
        }
        catch
        {
            bExists = false;
        }



        return bExists;

    }

    public void CopytoSharepoint(string fileName, string DestinationPath)
    {

        string folderpath = null, filepath;
        DateTime dtAUMdate = Convert.ToDateTime(txtAUMDate.Text);
        int yearNow = DateTime.Now.Year;
        int year = Convert.ToDateTime(dtAUMdate).Year;
        // DateTime datee = Convert.ToDateTime(dtAUMdate.AddYears(1));
        //int day = Convert.ToDateTime(dtAUMdate).Day;
        string month_day = dtAUMdate.ToString("MM/dd");

        Dictionary<string, string> val = Quater(dtAUMdate);
        string years = val.Keys.ElementAt(0).ToString();
        string Quarter = val.Values.ElementAt(0).ToString();

        folderpath = "Documents/Billing/" + years + "/" + Quarter + " " + years + "/Fee Calcs";


        //switch (month_day)
        //{
        //    case "03/31":
        //        folderpath = "Documents/Billing/" + year + "/Q2 " + year + "/Fee Calcs";

        //        break;

        //    case "06/30":
        //        folderpath = "Documents/Billing/" + year + "/Q3 " + year + "/Fee Calcs";

        //        break;
        //    case "09/30":
        //        folderpath = "Documents/Billing/" + year + "/Q4 " + year + "/Fee Calcs";
        //        break;
        //    case "12/31":
        //        dtAUMdate = dtAUMdate.AddYears(1);
        //        int year1 = Convert.ToDateTime(dtAUMdate).Year;
        //        folderpath = "Documents/Billing/" + year1 + "/Q1 " + year1 + "/Fee Calcs";
        //        break;
        //}

        if (CheckFolderPathExists(folderpath))
        {
            //if (checkSharepouintFileExist(folderpath, DestinationPath))
            //{
            CopyFilenew(folderpath, fileName, DestinationPath);
            //}
        }

        //}

    }

    #endregion
    #region SEC PROCESS
    public DataTable InvoiceId()
    {
        DataTable dtInvoiceId = null;
        try
        {
            DB clsDB = new DB();
            object HHValue = ddlHH.SelectedValue == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlHH.SelectedValue + "'";

            string strBillFor = ddlBillFor.SelectedValue;
            char[] delimiterChars = { '|' };
            string[] words = strBillFor.ToString().Split(delimiterChars);
            int len1 = words.Length;
            if (len1 > 0)
            {
                strBillFor = words[0];
                BillingUUID = words[0];
            }


            DataSet dsBillingChk1 = clsDB.getDataSet("SP_S_BILLINGINVOICE_CHECK @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");
            dtInvoiceId = dsBillingChk1.Tables[2];

        }
        catch (Exception ex)
        {
            dtInvoiceId = null;
        }
        return dtInvoiceId;
    }
    public string UploadToLib(string DestinationFilePath, string EndDate)//, Microsoft.SharePoint.Client.ListItem oldlist
    {
        bool Upload = false;
        string FolderPath = string.Empty;
        int dtYear = DateTime.Now.Year;
        DateTime dtAUMdate = Convert.ToDateTime(EndDate.Replace("'", ""));
        string Quarter = dtAUMdate.ToString("MM/dd");
        int Month = dtAUMdate.Month;

        Dictionary<string, string> val = Quater(dtAUMdate);

        //if (Quarter == "03/31")
        //{
        //    Quarter = dtYear + " " + "Q2";
        //}
        //else if (Quarter == "06/30")
        //{
        //    Quarter = dtYear + " " + "Q3";
        //}
        //else if (Quarter == "09/30")
        //{
        //    Quarter = dtYear + " " + "Q4";
        //}
        //else if (Quarter == "12/31")
        //{
        //    Quarter = dtYear + " " + "Q1";
        //}
        //else
        //{
        //if (Month < 2)
        //{
        //   Quarter = dtYear + " " + "Q4";
        //}
        //else if (Month >= 3 && Month <=5)
        //{
        //    Quarter = dtYear + " " + "Q4";
        //}
        //else if (Month >=6 && Month <= 8)
        //{
        //    Quarter = dtYear + " " + "Q3";
        //}
        //else if (Month >= 9 && Month <= 12)
        //{
        //    Quarter = dtYear + " " + "Q2";
        //}
        //}

        string year = val.Keys.ElementAt(0).ToString();
        Quarter = year + " " + val.Values.ElementAt(0).ToString(); ;
        try
        {

            FolderPath = "Billing" + "/" + year + "/" + Quarter;
            string siteUrl = "https://greshampartners.sharepoint.com/sites/DataBackup/";
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);
            Web site = context.Web;

            byte[] bytes = System.IO.File.ReadAllBytes(DestinationFilePath);
            System.IO.Stream stream = new System.IO.MemoryStream(bytes);

            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(DestinationFilePath), Overwrite = true };
            currentRunFolder.Files.Add(newFile);
            currentRunFolder.Update();

            context.ExecuteQuery();
            Upload = true;

            if (Upload = true)
            {
                #region TAgging
                FileCollection filecoll = currentRunFolder.Files;
                context.Load(filecoll);
                context.ExecuteQuery();
                if (filecoll.Count > 0)
                {
                    foreach (Microsoft.SharePoint.Client.File f in filecoll)
                    {
                        string FileName = string.Empty;
                        Microsoft.SharePoint.Client.ListItem listItems = f.ListItemAllFields;
                        context.Load(listItems);
                        context.ExecuteQuery();
                        try
                        {
                            if (listItems["FileLeafRef"].ToString() != null)
                            {
                                FileName = listItems["FileLeafRef"].ToString();

                                if (FileName == Path.GetFileName(DestinationFilePath))
                                {
                                    string CurrentUserName = GetcurrentUser("Name");
                                    listItems["Generated_x0020_By"] = CurrentUserName;
                                    //listItems["As_x0020_of_x0020_Date"] = EndDate.Replace("'", "");
                                    //listItems["Batch_x0020_Name"] = "";
                                    listItems.Update();
                                    context.ExecuteQuery();
                                    break;
                                }
                            }
                        }
                        catch
                        {
                        }
                    }
                }

                #endregion
            }
            return FolderPath;

        }
        catch (Exception Ex)
        {
            lblMessage.Text = "Error Uploading File to sharepoint(Billing)";
            return "";

        }

    }

    public Dictionary<string, string> Quater(DateTime dt)
    {
        DateTime Q1;
        string year = dt.Year.ToString();


        string value = "";
        // Q1
        DateTime Qt1 = DateTime.Parse("10/31/" + year);
        DateTime Qt2 = DateTime.Parse("12/31/" + year);

        // Q2
        DateTime Qt3 = DateTime.Parse("01/31/" + year);
        DateTime Qt4 = DateTime.Parse("03/31/" + year);

        // Q3
        DateTime Qt5 = DateTime.Parse("4/30/" + year);
        DateTime Qt6 = DateTime.Parse("6/30/" + year);

        // Q4
        DateTime Qt7 = DateTime.Parse("7/31/" + year);
        DateTime Qt8 = DateTime.Parse("9/30/" + year);


        if (dt >= Qt1 && dt <= Qt2)
        {

            value = "Q1";
            year = (Convert.ToInt32(year) + 1).ToString();
        }
        else if (dt >= Qt3 && dt <= Qt4)
            value = "Q2";
        else if (dt >= Qt5 && dt <= Qt6)
            value = "Q3";
        else if (dt >= Qt7 && dt <= Qt8)
            value = "Q4";


        Dictionary<string, string> data = new Dictionary<string, string>();
        data.Add(year, value);

        return data;
    }

    private string GetcurrentUser(string Type)
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        //string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        ////Response.Write("Name1 =" + strName);
        //if (Request.Url.AbsoluteUri.Contains("localhost"))
        //{
        //    strName = AppLogic.GetParam(AppLogic.ConfigParam.UserName);
        //    strName = "corp\\" + strName;
        //}

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
        sqlstr = "select top 1 internalemailaddress,systemuserid,fullName from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            if (Type == "Name")
                return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["fullName"]);
            else
                return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
            //return UserID = "DFCE21B1-B81E-E211-A2B7-0002A5443D86";
        }
        else
        {
            return UserID = "";
        }
    }
    private DataSet ExcelProcedure(string EndDate, string BillingId)
    {
        string greshamquery;
        int totalCount = 0;
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {
            greshamquery = "SP_S_SEC_DATADUMP @BillingForUUID='" + BillingId + "',@AsofDT='" + EndDate + "',@BillingFlg = 1";

            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            dagersham.SelectCommand.CommandTimeout = 600;
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);


        }
        catch (Exception exc)
        {
            totalCount = 0;

        }

        return ds_gresham;
    }
    public string GenerateExcel(DataSet ds)
    {
        try
        {
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


            if (!Directory.Exists(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder"))
            {
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder");
            }

            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            string strhour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMin = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSec = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
            string append_timestamp = strMonth + "_" + strDay + "_" + strYear + "_" + strhour + "_" + strMin + "_" + strSec;
            String lsFileNamforFinalXls = string.Empty;
            if (Server.ToLower() == "prod")
            {
                lsFileNamforFinalXls = ddlHH.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd") + "_" + append_timestamp;
            }
            else
            {
                lsFileNamforFinalXls = ddlHH.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd") + "_" + append_timestamp + "_test";
            }

            //commented and changed 7_17_2019 - added Server to Params.config
            //String lsFileNamforFinalXls = ddlHH.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd") + "_" + append_timestamp;
            string ExcelFilePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls + ".xlsx";

            if (System.IO.File.Exists(ExcelFilePath))
            {
                System.IO.File.Delete(ExcelFilePath);
            }

            //  string ExcelFilePath = ExcelfilePath + "TradingAppRecon" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
            if (ds.Tables.Count > 0)
            {
                #region EPP
                //FileInfo newFile = new FileInfo(ExcelFilePath);

                //using (OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(newFile))
                //{

                //    for (int i = 0; i < ds.Tables.Count; i++)
                //    {
                //        string SheetNme = ds.Tables[i].Rows[0][0].ToString();
                //        i++;
                //        OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add(SheetNme);
                //        if (ds.Tables[i].Rows.Count > 0)
                //        {
                //            ws.Cells["A1"].LoadFromDataTable(ds.Tables[i], true);
                //            WorksheetFormatting(ws);
                //        }
                //        else
                //        {
                //            ws.Cells["J9:L10"].Merge = true;
                //            ws.Cells["J9:L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                //            ws.Cells["J9:L10"].Value = "No Data Found";
                //            ws.Cells["J9:L10"].Style.Font.Size = 16;
                //            //ws.Cells["J9:L10"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                //        }

                //    }
                //    pck.Save();

                //}
                #endregion
                #region Spire License Code
                string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
                Spire.License.LicenseProvider.SetLicenseKey(License);
                Spire.License.LicenseProvider.LoadLicense();
                #endregion

                string SheetNme = ds.Tables[0].Rows[0][0].ToString();

                Workbook book = new Workbook();
                Worksheet sheet = book.Worksheets[0];
                sheet.Name = SheetNme;
                sheet.Range[1, 1, 1, ds.Tables[1].Columns.Count].Style.Font.IsBold = true;

                sheet.InsertDataTable(ds.Tables[1], true, 1, 1);

                sheet.Range[2, 12, ds.Tables[1].Rows.Count + 1, 12].NumberFormat = "$ #,##0.00_);($ #,##0.00)";
                sheet.Range[2, 13, ds.Tables[1].Rows.Count + 1, 13].NumberFormat = "$ #,##0.00_);($ #,##0.00)";
                sheet.Range[2, 14, ds.Tables[1].Rows.Count + 1, 14].NumberFormat = "$ #,##0.00_);($ #,##0.00)";

                sheet.Range[1, 1, ds.Tables[1].Rows.Count + 1, ds.Tables[1].Columns.Count].AutoFitColumns();
                sheet.Range[1, 1, ds.Tables[1].Rows.Count + 1, ds.Tables[1].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;

                book.SaveToFile(ExcelFilePath, ExcelVersion.Version2016);
                string vContain = "Excel Report Generated Succesfully ";
            }
            return ExcelFilePath;

        }
        catch (Exception e)
        {

            string vContain = "Excel Report Genration Fail,  Error " + e.ToString();

            return "";
        }
    }
    public void WorksheetFormatting(OfficeOpenXml.ExcelWorksheet ws)
    {
        int totalCols = ws.Dimension.End.Column;
        var headerCells = ws.Cells[1, 1, 1, totalCols];
        var headerFont = headerCells.Style.Font;
        headerFont.Bold = true;
        //ws.DeleteColumn(7);//added on 02_27_2018 by brijesh
        //ws.DeleteColumn(7);//added on 02_27_2018 by brijesh
        int totalRows = ws.Dimension.End.Row;
        //added on 02_27_2018 by brijesh
        ws.Cells[1, 12, totalRows, 12].Style.Numberformat.Format = "$ #,##0.00_);($ #,##0.00)";
        ws.Cells[1, 13, totalRows, 13].Style.Numberformat.Format = "$ #,##0.00_);($ #,##0.00)";
        ws.Cells[1, 14, totalRows, 14].Style.Numberformat.Format = "$ #,##0.00_);($ #,##0.00)";
        var allCells = ws.Cells[1, 1, totalRows, totalCols];
        allCells.AutoFitColumns();
        allCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
    }
    public void updateFields(string billinginvoiceId, string BillingPath, string fileName)
    {
        #region crm Connection

        IOrganizationService service = null;
        string strDescription = "";

        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        #endregion

        string BillingPath1 = SharepointPath + "/" + BillingPath + "/" + fileName;

        Entity objssi_billinginvoice = new Entity("ssi_billinginvoice");

        Guid InvoiceId = new Guid(Convert.ToString(billinginvoiceId));
        objssi_billinginvoice["ssi_billinginvoiceid"] = InvoiceId;

        objssi_billinginvoice["ssi_billingauditbackup"] = Convert.ToString(BillingPath1);

        service.Update(objssi_billinginvoice);
    }
    #endregion
    public void submitold()
    {
        if (!Page.IsValid)
        {
            //Show Validation Summary
            ValidationSummary1.ShowMessageBox = true;
            return;
        }

        DB clsDB = new DB();
        object HHValue = ddlHH.SelectedValue == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlHH.SelectedValue + "'";
        //string sqlstr = "[SP_S_BILLING_NAME] @HouseHoldUUID = " + HHValue + " ";
        //DataSet ds1 = clsDB.getDataSet(sqlstr);
        //BillingUUID = Convert.ToString(ds1.Tables[0].Rows[0]["Ssi_billingid"]);

        string strBillFor = ddlBillFor.SelectedValue;
        char[] delimiterChars = { '|' };
        string[] words = strBillFor.ToString().Split(delimiterChars);
        int len = words.Length;
        if (len > 0)
        {
            strBillFor = words[0];
            BillingUUID = words[0];
        }

        /* Return BillingInvoice  into dataset where HouseHold AumAsofDate and BillingForId is Selected */
        DataSet dsBillingChk = clsDB.getDataSet("SP_S_BILLINGINVOICE_CHECK @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");
        if (dsBillingChk.Tables[0].Rows.Count > 0)
        {
            ViewState["dsBillingCheck"] = dsBillingChk;
            this.Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "Confirm();", true);
            // System.Threading.Thread.Sleep(10000);


            //string confirmValue = Request.Form["confirm_value"];
            //if (confirmValue == "Yes")
            //{
            //    this.Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('You clicked YES!')", true);
            //}
            //else
            //{
            //    this.Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('You clicked NO!')", true);
            //}

            Timer1.Enabled = true;


        }
        else
        {
            // CreateInvoice(BillingUUID);//Creating New Invoice
        }


    }

    // created By abhi
    protected void CreateInvoiceNew(string BillingUUID, string BillingID, string Persent, string ssi_relationshipfee, string ssi_relationshipfeeFlg, string ssi_setupfee, string ssi_SetUpFeeFlg, string ssi_OtherFee, string ssi_OtherFeeFlg, string FinalAUMMarketValue, string vBillingMarketValue, string ssi_FeePctAmt, string ssi_ScheduleAUM)
    {
        try
        {
            #region crm Connection
            string orgName = "GreshamPartners";

            //CRM SERVICE
            IOrganizationService service = null;
            //  string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

            string strDescription = "";
            try
            {
                service = GM.GetCrmService();
                strDescription = "Crm Service starts successfully";
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                lblMessage.Text = strDescription;
            }
            catch (Exception exc)
            {
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                lblMessage.Text = strDescription;
            }


            try
            {
                //Response.Write(service.Url);
                //service.PreAuthenticate = true;
                //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
            }
            catch (NullReferenceException ne)
            {
                //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
            }

            #endregion

            string ssiName = ddlBillFor.SelectedItem.Text + "-" + txtAUMDate.Text;
            // string BillingID = ViewState["BillingId"].ToString();
            string AUMDate = txtAUMDate.Text;
            string TotalAum = txtTotalAUM.Text;
            string BillingAUM = txtBillingAUM.Text;
            string StdAnnualFeeCalc = txtStdAnnualFeeCalc.Text;
            string CustFeeAmount = txtCustFeeAmount.Text;
            string FeeRateCalc = txtFeeRateCalc.Text;
            string QuaterlyFeeCalc = txtQuaterlyFeeCalc.Text;
            string FeesPerMonth = txtFeesPerMonth.Text;
            string AdjQtrFee = txtAdjQtrFee.Text;
            string AdjAmt = txtAdjAmt.Text;
            string AdjReason = txtAdjReason.Text;
            string ClientType = ddlClientType.SelectedValue;
            string RelationshipFee = txtRelationshipFee.Text;
            // string CustFeeAmount = txtCustFeeAmount.Text;
            string Discount = txtDiscount.Text;

            //if (Persent != "")
            //{
            //    StdAnnualFeeCalc = (Convert.ToDecimal(StdAnnualFeeCalc) * Convert.ToDecimal(Persent)).ToString();
            //    CustFeeAmount = (Convert.ToDecimal(CustFeeAmount) * Convert.ToDecimal(Persent)).ToString();
            //    FeeRateCalc = (Convert.ToDecimal(FeeRateCalc) * Convert.ToDecimal(Persent)).ToString();
            //    FeesPerMonth = (Convert.ToDecimal(FeesPerMonth) * Convert.ToDecimal(Persent)).ToString();
            //}
            DataSet dsData;
            int count = 0;
            string relationshipFees = "", setupFee = "", text3 = "";


            dsData = (DataSet)ViewState["dsData"];

            count = dsData.Tables[1].Rows.Count;

            //txtRelationshipFee   txtCustFeeAmount  txtDiscount
            string InvoiceLoadId = Guid.NewGuid().ToString();

            // ssi_billinginvoice objBillingInvoice = new ssi_billinginvoice();
            Entity objBillingInvoice = new Entity("ssi_billinginvoice");

            //  objBillingInvoice.ssi_invoiceloadid = InvoiceLoadId;
            objBillingInvoice["ssi_invoiceloadid"] = Guid.NewGuid().ToString();

            string invoidcID = Guid.NewGuid().ToString();
            //objBillingInvoice.ssi_billinginvoiceid = new Key();
            //objBillingInvoice.ssi_billinginvoiceid.Value = new Guid(invoidcID);
            objBillingInvoice["ssi_billinginvoiceid"] = new Guid(invoidcID);


            //  objBillingInvoice.ssi_billinginvoiceid = new Key();
            // objBillingInvoice.ssi_billinginvoiceid.Value = Guid.NewGuid();

            //Name
            //objBillingInvoice.ssi_name = ssiName;// ddlBillFor.SelectedItem.Text + "-" + txtAUMDate.Text;
            objBillingInvoice["ssi_name"] = ssiName;
            // ssi_billingid
            if (Convert.ToString(BillingUUID) != "")
            {
                //objBillingInvoice.ssi_billingprimaryid = new Lookup();
                //objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
                //objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(BillingUUID));
                objBillingInvoice["ssi_billingprimaryid"] = new EntityReference("ssi_billing", new Guid(Convert.ToString(BillingUUID)));
            }

            //                //InvoiceId
            //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]) != "")
            //                {
            //                    objBillingInvoice.ssi_invoiceid = Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]);
            //                }

            //BillingId
            /* */
            //  objBillingInvoice.ssi_billingid = new CrmNumber();

            int BillID = GetBillingID();
            if (BillingID != "")
            {
                //objBillingInvoice.ssi_billingid = new CrmNumber();   //ssi_billingid-->value comes from Crm
                //objBillingInvoice.ssi_billingid.Value = Convert.ToInt32(BillID);
                objBillingInvoice["ssi_billingid"] = Convert.ToInt32(BillID);
            }

            //AUM Date 
            if (txtAUMDate.Text != "")
            {
                //objBillingInvoice.ssi_aumasofdate = new CrmDateTime();//ssi_aumasofdate-->value comes from Crm
                //objBillingInvoice.ssi_aumasofdate.Value = Convert.ToString(txtAUMDate.Text);
                objBillingInvoice["ssi_aumasofdate"] = Convert.ToDateTime(Convert.ToString(txtAUMDate.Text));
            }

            //Total AUM
            if (Convert.ToDecimal(FinalAUMMarketValue) != 0)
            {
                //objBillingInvoice.ssi_totalaum = new CrmMoney();//ssi_totalaum-->value comes from Crm
                //objBillingInvoice.ssi_totalaum.Value =Convert.ToDecimal(FinalAUMMarketValue);
                objBillingInvoice["ssi_totalaum"] = new Money(Convert.ToDecimal(FinalAUMMarketValue));

            }
            else
            {
                if (Convert.ToString(txtTotalAUM.Text) != "")
                {
                    decimal total = Convert.ToDecimal(txtTotalAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                    total = total * Convert.ToDecimal(Persent);

                    //objBillingInvoice.ssi_totalaum = new CrmMoney();//ssi_totalaum-->value comes from Crm
                    //objBillingInvoice.ssi_totalaum.Value = Convert.ToDecimal(total);
                    objBillingInvoice["ssi_totalaum"] = new Money(Convert.ToDecimal(total));
                }
            }
            //Billing AUM
            //if (Convert.ToString(txtBillingAUM.Text) != "")
            //{
            //decimal total = Convert.ToDecimal(txtBillingAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //total = total * Convert.ToDecimal(Persent);
            if (Convert.ToDecimal(vBillingMarketValue) != 0)
            {
                //objBillingInvoice.ssi_aum = new CrmMoney();//ssi_aum-->value comes from Crm
                //objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(vBillingMarketValue);
                objBillingInvoice["ssi_totalbillableassets"] = new Money(Convert.ToDecimal(vBillingMarketValue));
            }

            else
            {
                decimal total = Convert.ToDecimal(txtBillingAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_aum = new CrmMoney();              //ssi_aum-->value comes from Crm
                //objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_totalbillableassets"] = new Money(Convert.ToDecimal(total));
            }

            if (Convert.ToDecimal(ssi_ScheduleAUM) != 0)
            {
                //objBillingInvoice.ssi_aum = new CrmMoney();//ssi_aum-->value comes from Crm
                //objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(vBillingMarketValue);
                objBillingInvoice["ssi_aum"] = new Money(Convert.ToDecimal(ssi_ScheduleAUM));
            }

            else
            {
                decimal total = Convert.ToDecimal(txtBillAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_aum = new CrmMoney();              //ssi_aum-->value comes from Crm
                //objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_aum"] = new Money(Convert.ToDecimal(total));
            }


            //if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
            //{
            //    objBillingInvoice.ssi_annualfee = new CrmMoney();
            //    objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //}
            //else
            //{
            //Annual Fee == isnull(custom fee, standard annual fee calc)

            if (Convert.ToString(txtDiscount.Text) != "")
            {
                decimal total = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_annualfee = new CrmMoney();
                //objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_annualfee"] = new Money(Convert.ToDecimal(total));
            }
            else
            {
                if (Convert.ToString(txtCustFeeAmount.Text) != "" && Convert.ToString(txtCustFeeAmount.Text).Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "") != "0.00")
                {
                    string fees = txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                    fees = (Convert.ToDecimal(fees) * Convert.ToDecimal(Persent)).ToString();

                    //objBillingInvoice.ssi_annualfee = new CrmMoney();
                    //objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(fees);
                    objBillingInvoice["ssi_annualfee"] = new Money(Convert.ToDecimal(fees));
                }
                else
                {
                    if (Convert.ToString(txtStdAnnualFeeCalc.Text) != "")// && Convert.ToString(txtStdAnnualFeeCalc.Text).Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "") != "0.00")
                    {
                        string fees = txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                        fees = (Convert.ToDecimal(fees) * Convert.ToDecimal(Persent)).ToString();

                        //objBillingInvoice.ssi_annualfee = new CrmMoney();
                        //objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(fees);
                        objBillingInvoice["ssi_annualfee"] = new Money(Convert.ToDecimal(fees));
                    }
                }
            }
            //}

            //Fee Rate 
            if (Convert.ToString(txtFeeRateCalc.Text) != "")
            {
                string fee = txtFeeRateCalc.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace("$", "");
                if (fee != "NA")
                {
                    //objBillingInvoice.ssi_feerate = new CrmDecimal();
                    //objBillingInvoice.ssi_feerate.Value = Convert.ToDecimal(fee);
                    objBillingInvoice["ssi_feerate"] = Convert.ToDecimal(fee);
                }
                else
                {
                    //objBillingInvoice.ssi_feerate = new CrmDecimal();
                    //objBillingInvoice.ssi_feerate.IsNull = true;
                    //objBillingInvoice.ssi_feerate.IsNullSpecified = true;
                    objBillingInvoice["ssi_feerate"] = null;

                }
            }

            //Quarterly Fee  
            if (Convert.ToString(txtQuaterlyFeeCalc.Text) != "")
            {
                string fees = txtQuaterlyFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                fees = (Convert.ToDecimal(fees) * Convert.ToDecimal(Persent)).ToString();

                //objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
                //objBillingInvoice.ssi_quarterlyfee.Value = Convert.ToDecimal(fees);
                objBillingInvoice["ssi_quarterlyfee"] = new Money(Convert.ToDecimal(fees));
            }

            # region NotUsed
            //if (count >= 1)
            //{
            //    //relationshipAlocationflag = dsData.Tables[1].Rows[0]["AllocatedFlg"].ToString();

            //    if (ssi_relationshipfeeFlg == "0")
            //    {
            //        relationshipFees = ssi_relationshipfee; // txtBox1.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");


            //    }
            //    else
            //    {
            //        //  relationshipFees = dsData.Tables[2].Rows[0]["ssi_relationshipfee"].ToString();
            //        objBillingInvoice.ssi_flatfeeallocationflag1 = new CrmBoolean();
            //        objBillingInvoice.ssi_flatfeeallocationflag1.Value = true;

            //        relationshipFees = txtBox1.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            //        relationshipFees = (Convert.ToDecimal(relationshipFees) * Convert.ToDecimal(Persent)).ToString();
            //    }

            //}
            //if (count >= 2)
            //{
            //    // setupAlocationflag = dsData.Tables[1].Rows[1]["AllocatedFlg"].ToString();
            //    if (ssi_SetUpFeeFlg == "0")
            //    {
            //        setupFee = ssi_setupfee;// txtBox2.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            //    }
            //    else
            //    {
            //        objBillingInvoice.ssi_flatfeeallocationflag2 = new CrmBoolean();
            //        objBillingInvoice.ssi_flatfeeallocationflag2.Value = true;

            //        setupFee = txtBox2.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            //        setupFee = (Convert.ToDecimal(setupFee) * Convert.ToDecimal(Persent)).ToString();
            //    }
            //}
            //if (count == 3)
            //{
            //    // OtherAlocationflag = dsData.Tables[1].Rows[2]["AllocatedFlg"].ToString();
            //    if (ssi_OtherFeeFlg == "0")
            //    {
            //        text3 = ssi_OtherFee;// txtBox3.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");

            //    }
            //    else
            //    {
            //        objBillingInvoice.ssi_flatfeeallocationflag3 = new CrmBoolean();
            //        objBillingInvoice.ssi_flatfeeallocationflag3.Value = true;

            //        text3 = txtBox3.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            //        text3 = (Convert.ToDecimal(text3) * Convert.ToDecimal(Persent)).ToString();
            //    }

            //}
            #endregion

            //string Monthlyfees=txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            //Monthlyfees= (Convert.ToDecimal(txtFeesPerMonth) * Convert.ToDecimal(Persent)).ToString();
            //Month1 Fee
            if (Convert.ToString(txtFeesPerMonth.Text) != "")
            {
                string fees = txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                fees = (Convert.ToDecimal(fees) * Convert.ToDecimal(Persent)).ToString();

                //objBillingInvoice.ssi_month1fee = new CrmMoney();
                //objBillingInvoice.ssi_month1fee.Value = Convert.ToDecimal(fees);
                objBillingInvoice["ssi_month1fee"] = new Money(Convert.ToDecimal(fees));
                //Month2 Fee
                //objBillingInvoice.ssi_month2fee = new CrmMoney();
                //objBillingInvoice.ssi_month2fee.Value = Convert.ToDecimal(fees);
                objBillingInvoice["ssi_month2fee"] = new Money(Convert.ToDecimal(fees));
                //Month3 Fee
                //objBillingInvoice.ssi_month3fee = new CrmMoney();
                //objBillingInvoice.ssi_month3fee.Value = Convert.ToDecimal(fees);
                objBillingInvoice["ssi_month3fee"] = new Money(Convert.ToDecimal(fees));
            }

            //Month2 Fee
            //if (Convert.ToString(txtFeesPerMonth.Text) != "")
            //{
            //    objBillingInvoice.ssi_month2fee = new CrmMoney();
            //    objBillingInvoice.ssi_month2fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //}

            //Month3 Fee
            //if (Convert.ToString(txtFeesPerMonth.Text) != "")
            //{
            //    objBillingInvoice.ssi_month3fee = new CrmMoney();
            //    objBillingInvoice.ssi_month3fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //}

            System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");

            DateTime dtAUMdate = Convert.ToDateTime(txtAUMDate.Text);
            //  DateTime dtAUMdate = DateTime.ParseExact(txtAUMDate.Text, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            string Month1 = dtAUMdate.AddMonths(2).ToString("MMMM");//full month can be assigned--->"January"
            string Month2 = dtAUMdate.AddMonths(3).ToString("MMMM");
            string Month3 = dtAUMdate.AddMonths(4).ToString("MMMM");

            //objBillingInvoice.ssi_month1 = Month1;
            //objBillingInvoice.ssi_month2 = Month2;
            //objBillingInvoice.ssi_month3 = Month3;
            objBillingInvoice["ssi_month1"] = Month1;
            objBillingInvoice["ssi_month2"] = Month2;
            objBillingInvoice["ssi_month3"] = Month3;


            // objBillingInvoice.ssifeety
            //string relationshipFees = "", setupFee = "", text3 = "";
            //if (relationshipFees != "")
            //{

            //    objBillingInvoice.ssi_feetype1 = new Picklist();
            //    // objBillingInvoice.ssi_feetype1.
            //    objBillingInvoice.ssi_feetype1.Value = 100000000;

            //    objBillingInvoice.ssi_feeamount1 = new CrmMoney();
            //    objBillingInvoice.ssi_feeamount1.Value = Convert.ToDecimal(relationshipFees);

            //}
            //if (setupFee != "")
            //{
            //    objBillingInvoice.ssi_feetype2 = new Picklist();
            //    objBillingInvoice.ssi_feetype2.Value = 100000001;

            //    objBillingInvoice.ssi_feeamount2 = new CrmMoney();
            //    objBillingInvoice.ssi_feeamount2.Value = Convert.ToDecimal(setupFee);
            //}
            //if (text3 != "")
            //{
            //    objBillingInvoice.ssi_feetype3 = new Picklist();
            //    objBillingInvoice.ssi_feetype3.Value = 100000002;

            //    objBillingInvoice.ssi_feeamount3 = new CrmMoney();
            //    objBillingInvoice.ssi_feeamount3.Value = Convert.ToDecimal(text3);
            //}

            //ssi_AdjustedFee
            if (Convert.ToString(txtAdjQtrFee.Text) != "")
            {
                string fees = txtAdjQtrFee.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                fees = (Convert.ToDecimal(fees) * Convert.ToDecimal(Persent)).ToString();

                //objBillingInvoice.ssi_adjustedfee = new CrmMoney();
                //objBillingInvoice.ssi_adjustedfee.Value = Convert.ToDecimal(fees);
                objBillingInvoice["ssi_adjustedfee"] = new Money(Convert.ToDecimal(fees));
            }


            //ssi_adjustment
            if (Convert.ToString(txtAdjAmt.Text) != "")
            {
                string fees = txtAdjAmt.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                fees = (Convert.ToDecimal(fees) * Convert.ToDecimal(Persent)).ToString();

                //objBillingInvoice.ssi_adjustment = new CrmMoney();
                //objBillingInvoice.ssi_adjustment.Value = Convert.ToDecimal(fees);
                objBillingInvoice["ssi_adjustment"] = new Money(Convert.ToDecimal(fees));
            }

            //ssi_adjustmentreason  

            if (Convert.ToString(txtAdjReason.Text) != "")
            {
                //objBillingInvoice.ssi_adjustmentreason = Convert.ToString(txtAdjReason.Text);
                objBillingInvoice["ssi_adjustmentreason"] = Convert.ToString(txtAdjReason.Text);
            }

            // allocation persentage
            //objBillingInvoice.ssi_allocationpercent = new CrmDecimal();
            //objBillingInvoice.ssi_allocationpercent.Value = Convert.ToDecimal(Persent);
            objBillingInvoice["ssi_allocationpercent"] = Convert.ToDecimal(Persent);

            //ClientType
            if (Convert.ToString(ddlClientType.SelectedValue) != "")
            {
                //objBillingInvoice.ssi_feescheduletype = new Picklist();
                ////objBillingInvoice.ssi_feescheduleid.type = EntityName.ssi_feeschedule.ToString();
                //objBillingInvoice.ssi_feescheduletype.Value = Convert.ToInt32(ddlClientType.SelectedValue);
                objBillingInvoice["ssi_feescheduletype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ddlClientType.SelectedValue));
            }


            //Custom FeeAmount
            if (Convert.ToString(txtCustFeeAmount.Text) != "")
            {
                string fees = txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                fees = (Convert.ToDecimal(fees) * Convert.ToDecimal(Persent)).ToString();

                //objBillingInvoice.ssi_customfee = new CrmMoney();
                //objBillingInvoice.ssi_customfee.Value = Convert.ToDecimal(fees);
                objBillingInvoice["ssi_customfee"] = new Money(Convert.ToDecimal(fees));
            }
            else
            {
                //objBillingInvoice.ssi_customfee = new CrmMoney();
                //objBillingInvoice.ssi_customfee.IsNull = true;
                //objBillingInvoice.ssi_customfee.IsNullSpecified = true;
                objBillingInvoice["ssi_customfee"] = null;
            }

            //discount   
            if (Convert.ToString(txtDiscount.Text) != "")
            {
                //objBillingInvoice.ssi_discount = new CrmDecimal();
                //objBillingInvoice.ssi_discount.Value = Convert.ToDecimal(txtDiscount.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "")) / 100;
                objBillingInvoice["ssi_discount"] = Convert.ToDecimal(txtDiscount.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "")) / 100;
            }
            else
            {
                //objBillingInvoice.ssi_discount = new CrmDecimal();
                //objBillingInvoice.ssi_discount.IsNull = true;
                //objBillingInvoice.ssi_discount.IsNullSpecified = true;
                objBillingInvoice["ssi_discount"] = null;
            }

            if (chkAccured.Checked)
            {
                //objBillingInvoice.ssi_accrued = new CrmBoolean();
                //objBillingInvoice.ssi_accrued.Value = true;
                objBillingInvoice["ssi_accrued"] = true;
            }
            else
            {
                //objBillingInvoice.ssi_accrued = new CrmBoolean();
                //    objBillingInvoice.ssi_accrued.Value = false;
                objBillingInvoice["ssi_accrued"] = false;
            }
            #region added 3_25_2019 Standard Min-Max Change

            string bpsfeeValue = string.Empty;
            string MinValue = string.Empty;
            string MaxValue = string.Empty;
            try
            {
                bpsfeeValue = ViewState["bpsfeeValue"].ToString();
            }
            catch (Exception ex)
            {
                bpsfeeValue = "";
            }
            try
            {
                MinValue = ViewState["MinValue"].ToString();
            }
            catch (Exception ex1)
            {
                MinValue = "";
            }
            try
            {
                MaxValue = ViewState["MaxValue"].ToString();
            }
            catch (Exception ex2)
            {
                MaxValue = "";
            }

            //Fee on first $25MM in bps - added 3_25_2019 Standard  Min-Max Change
            if (bpsfeeValue != "")
            {
                string fee = bpsfeeValue.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace("$", "");
                if (fee != "NA")
                {
                    objBillingInvoice["ssi_feeonfirst25mminbps"] = Convert.ToDecimal(fee);
                }
                else
                {
                    objBillingInvoice["ssi_feeonfirst25mminbps"] = null;
                }
            }
            else
            {
                objBillingInvoice["ssi_feeonfirst25mminbps"] = null;
            }
            //Minimum fee in $ - added 3_25_2019 Standard Min-Max Change

            if (txtMinVal.Visible == true || txtMinVal.Text != "")
            {

                if (MinValue != "")
                {
                    string fees = MinValue.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                    fees = Convert.ToDecimal(fees).ToString();

                    objBillingInvoice["ssi_minimumfeein"] = new Money(Convert.ToDecimal(fees));
                }
            }
            else
            {
                objBillingInvoice["ssi_minimumfeein"] = null;
            }



            //Maximum fee as a %- added 3_25_2019 Standard Min-Max Change
            if (txtMaxVal.Visible == true || txtMaxVal.Text != "")
            {
                if (MaxValue != "")
                {
                    string fee = MaxValue.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace("$", "");
                    if (fee != "NA")
                    {
                        objBillingInvoice["ssi_maximumfeeasa"] = Convert.ToDecimal(fee);
                    }
                    else
                    {
                        objBillingInvoice["ssi_maximumfeeasa"] = null;
                    }
                }
            }
            else
            {
                objBillingInvoice["ssi_maximumfeeasa"] = null;
            }
            if (ddlClientType.SelectedValue == "100000009")// Standard Fee or $180K(greater of the two)// if (ddlClientType.SelectedValue == "100000000")
            {
                //Alt Fee Schedule Type- added 3_25_2019 Standard Min-Max Change
                if (bpsfeeValue != "" && MinValue != "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Discounted Standard (" + bpsfeeValue + " bps first $25mil) or at least " + MinValue + " not to exceed a maximum of " + MaxValue;
                }
                else if (bpsfeeValue == "" && MinValue != "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Standard Fee or a minimum " + MinValue + ", not to exceed a max of " + MaxValue + " of AUM";
                }
                else if (bpsfeeValue == "" && MinValue == "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Standard Fee or the lesser of " + MaxValue;
                }
                else if (bpsfeeValue == "" && MinValue != "" && MaxValue == "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Standard Fee or the greater of " + MinValue;
                }
                else if (bpsfeeValue != "" && MinValue == "" && MaxValue == "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Discounted Standard (" + bpsfeeValue + " bps first $25mil)";
                }
                else if (bpsfeeValue != "" && MinValue != "" && MaxValue == "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Discounted Standard (" + bpsfeeValue + " bps first $25mil), Standard Fee or the greater of " + MinValue;
                }
                else if (bpsfeeValue != "" && MinValue == "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Discounted Standard (" + bpsfeeValue + " bps first $25mil), Standard Fee or the lesser of " + MaxValue;
                }
                else
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = null;
                }

            }
            else
            {
                objBillingInvoice["ssi_altfeescheduletype"] = null;
            }




            #endregion

            //added 7_15_2019
            if (ssi_FeePctAmt != "")
            {
                objBillingInvoice["ssi_securityfeeaum"] = new Money(Convert.ToDecimal(ssi_FeePctAmt));
            }
            else
            {
                objBillingInvoice["ssi_securityfeeaum"] = null;
            }
            if (ssi_ScheduleAUM != "")
            {
                // objBillingInvoice["ssi_scheduleaum"] = new Money(Convert.ToDecimal(ssi_ScheduleAUM));
            }
            else
            {
                objBillingInvoice["ssi_scheduleaum"] = null;
            }


            service.Create(objBillingInvoice);

            insertCustomefeeNew(invoidcID, BillingUUID);


            /*Invoice Created*/

            System.Web.UI.WebControls.ListItem itemToRemove = ddlClientType.Items.FindByValue("0");
            if (itemToRemove != null)
            {
                ddlClientType.Items.Remove(itemToRemove);
            }

            lblMessage.Text = "Records Created Successfully";
        }


        catch (Exception ex)
        {
            //Failed to Create Invoice
            lblMessage.Text = "Billing Invoice Failed to Create : " + ex.ToString();
        }

    }
    protected void UpdateInvoiceNew(string BillingUUID, string BillingInvoiceId, string Persent, string ssi_relationshipfee, string ssi_relationshipfeeFlg, string ssi_setupfee, string ssi_SetUpFeeFlg, string ssi_OtherFee, string ssi_OtherFeeFlg, string vAumValue, string vFinalBillingValue, string ssi_FeePctAmt, string ssi_ScheduleAUM)
    {

        #region crm Connection

        //  string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//set crmserver url
        //  string orgName = "GreshamPartners";
        IOrganizationService service = null;
        string strDescription = "";

        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {

            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {

            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }



        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion


        try
        {
            if (Persent == "")
                Persent = "1";

            string InvoiceLoadId = Guid.NewGuid().ToString();

            // ssi_billinginvoice objBillingInvoice = new ssi_billinginvoice();
            Entity objBillingInvoice = new Entity("ssi_billinginvoice");

            // objBillingInvoice.ssi_invoiceloadid = InvoiceLoadId;

            objBillingInvoice["ssi_invoiceloadid"] = InvoiceLoadId;

            //objBillingInvoice.ssi_billinginvoiceid = new Key();
            //objBillingInvoice.ssi_billinginvoiceid.Value = new Guid(BillingInvoiceId);

            objBillingInvoice["ssi_billinginvoiceid"] = new Guid(BillingInvoiceId);


            //  objBillingInvoice.ssi_billinginvoiceid = new Key();
            // objBillingInvoice.ssi_billinginvoiceid.Value = Guid.NewGuid();

            //Name
            // objBillingInvoice.ssi_name = ddlBillFor.SelectedItem.Text + "-" + txtAUMDate.Text;
            objBillingInvoice["ssi_name"] = ddlBillFor.SelectedItem.Text + "-" + txtAUMDate.Text;

            // ssi_billingid
            if (Convert.ToString(BillingUUID) != "")
            {
                //objBillingInvoice.ssi_billingprimaryid = new Lookup();
                //objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
                //objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(BillingUUID));
                objBillingInvoice["ssi_billingprimaryid"] = new EntityReference("ssi_billing", new Guid(Convert.ToString(BillingUUID)));

            }

            //                //InvoiceId
            //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]) != "")
            //                {
            //                    objBillingInvoice.ssi_invoiceid = Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]);
            //                }


            ////BillingId
            //if (Convert.ToString(ViewState["BillingId"]) != "")
            //{
            //    objBillingInvoice.ssi_billingid = new CrmNumber();
            //    objBillingInvoice.ssi_billingid.Value = Convert.ToInt32(ViewState["BillingId"]);
            //}


            //AUM Date 
            if (txtAUMDate.Text != "")
            {
                //objBillingInvoice.ssi_aumasofdate = new CrmDateTime();
                //objBillingInvoice.ssi_aumasofdate.Value = Convert.ToString(txtAUMDate.Text);
                objBillingInvoice["ssi_aumasofdate"] = Convert.ToDateTime(Convert.ToString(txtAUMDate.Text));
            }
            else
            {
                //objBillingInvoice.ssi_aumasofdate.Value = "";
                objBillingInvoice["ssi_aumasofdate"] = "";
            }

            //  Total AUM

            if (Convert.ToDecimal(vAumValue) != 0)
            {
                //objBillingInvoice.ssi_totalaum = new CrmMoney();
                //objBillingInvoice.ssi_totalaum.Value = Convert.ToDecimal(vAumValue);
                objBillingInvoice["ssi_totalaum"] = new Money(Convert.ToDecimal(vAumValue));

            }
            else
            {

                if (Convert.ToString(txtTotalAUM.Text) != "")
                {
                    decimal total = Convert.ToDecimal(txtTotalAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                    total = total * Convert.ToDecimal(Persent);

                    //objBillingInvoice.ssi_totalaum = new CrmMoney();
                    //objBillingInvoice.ssi_totalaum.Value = Convert.ToDecimal(total);
                    objBillingInvoice["ssi_totalaum"] = new Money(Convert.ToDecimal(total));
                }
                else
                {
                    //objBillingInvoice.ssi_totalaum = new CrmMoney();
                    //objBillingInvoice.ssi_totalaum.IsNull = true;
                    //objBillingInvoice.ssi_totalaum.IsNullSpecified = true;
                    objBillingInvoice["ssi_totalaum"] = null;
                }
            }

            //Billing AUM
            if (Convert.ToString(txtBillingAUM.Text) != "")
            {
                //decimal total = Convert.ToDecimal(txtBillingAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                //total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_aum = new CrmMoney();
                //objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(vFinalBillingValue);
                objBillingInvoice["ssi_totalbillableassets"] = new Money(Convert.ToDecimal(vFinalBillingValue));
            }
            else
            {
                //objBillingInvoice.ssi_aum = new CrmMoney();
                //objBillingInvoice.ssi_aum.IsNull = true;
                //objBillingInvoice.ssi_aum.IsNullSpecified = true;
                objBillingInvoice["ssi_totalbillableassets"] = null;

            }

            if (Convert.ToString(txtBillAUM.Text) != "")
            {
                //decimal total = Convert.ToDecimal(txtBillingAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                //total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_aum = new CrmMoney();
                //objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(vFinalBillingValue);
                objBillingInvoice["ssi_aum"] = new Money(Convert.ToDecimal(ssi_ScheduleAUM));
            }
            else
            {
                //objBillingInvoice.ssi_aum = new CrmMoney();
                //objBillingInvoice.ssi_aum.IsNull = true;
                //objBillingInvoice.ssi_aum.IsNullSpecified = true;
                objBillingInvoice["ssi_aum"] = null;

            }

            ////if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
            ////{
            ////    objBillingInvoice.ssi_annualfee = new CrmMoney();
            ////    objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            ////}
            ////else
            ////{
            //Annual Fee == isnull(custom fee, standard annual fee calc)
            if (Convert.ToString(txtDiscount.Text) != "")
            {
                decimal total = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_annualfee = new CrmMoney();
                //objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_annualfee"] = new Money(Convert.ToDecimal(total));
            }
            else
            {
                if (Convert.ToString(txtCustFeeAmount.Text) != "" && Convert.ToString(txtCustFeeAmount.Text).Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "") != "0.00")
                {
                    decimal total = Convert.ToDecimal(txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                    total = total * Convert.ToDecimal(Persent);

                    //objBillingInvoice.ssi_annualfee = new CrmMoney();
                    //objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(total);
                    objBillingInvoice["ssi_annualfee"] = new Money(Convert.ToDecimal(total));
                }
                else
                {
                    decimal total = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                    total = total * Convert.ToDecimal(Persent);

                    //objBillingInvoice.ssi_annualfee = new CrmMoney();
                    //objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(total);
                    objBillingInvoice["ssi_annualfee"] = new Money(Convert.ToDecimal(total));
                }
            }
            ////}

            ////Fee Rate 
            if (Convert.ToString(txtFeeRateCalc.Text) != "")
            {
                //objBillingInvoice.ssi_feerate = new CrmDecimal();
                //objBillingInvoice.ssi_feerate.Value = Convert.ToDecimal(txtFeeRateCalc.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                objBillingInvoice["ssi_feerate"] = Convert.ToDecimal(txtFeeRateCalc.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            }
            else
            {
                //objBillingInvoice.ssi_feerate = new CrmDecimal();
                //objBillingInvoice.ssi_feerate.IsNull = true;
                //objBillingInvoice.ssi_feerate.IsNullSpecified = true;
                objBillingInvoice["ssi_feerate"] = null;

            }

            ////Quarterly Fee 
            if (Convert.ToString(txtQuaterlyFeeCalc.Text) != "")
            {
                decimal total = Convert.ToDecimal(txtQuaterlyFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
                //objBillingInvoice.ssi_quarterlyfee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_quarterlyfee"] = new Money(Convert.ToDecimal(total));
            }
            else
            {
                //objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
                //objBillingInvoice.ssi_quarterlyfee.IsNull = true;
                //objBillingInvoice.ssi_quarterlyfee.IsNullSpecified = true;
                objBillingInvoice["ssi_quarterlyfee"] = null; ;
            }

            //Month1 Fee
            if (Convert.ToString(txtFeesPerMonth.Text) != "")
            {
                decimal total = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_month1fee = new CrmMoney();
                //objBillingInvoice.ssi_month1fee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_month1fee"] = new Money(Convert.ToDecimal(total));

                //objBillingInvoice.ssi_month2fee = new CrmMoney();
                //objBillingInvoice.ssi_month2fee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_month2fee"] = new Money(Convert.ToDecimal(total));

                //objBillingInvoice.ssi_month3fee = new CrmMoney();
                //objBillingInvoice.ssi_month3fee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_month3fee"] = new Money(Convert.ToDecimal(total));
            }
            else
            {
                //objBillingInvoice.ssi_month1fee = new CrmMoney();
                //objBillingInvoice.ssi_month1fee.IsNull = true;
                //objBillingInvoice.ssi_month1fee.IsNullSpecified = true;
                objBillingInvoice["ssi_month1fee"] = null;

                //objBillingInvoice.ssi_month2fee = new CrmMoney();
                //objBillingInvoice.ssi_month2fee.IsNull = true;
                //objBillingInvoice.ssi_month2fee.IsNullSpecified = true;
                objBillingInvoice["ssi_month2fee"] = null;

                //objBillingInvoice.ssi_month3fee = new CrmMoney();
                //objBillingInvoice.ssi_month3fee.IsNull = true;
                //objBillingInvoice.ssi_month3fee.IsNullSpecified = true;
                objBillingInvoice["ssi_month3fee"] = null;
            }

            #region Not used
            //if (ssi_relationshipfeeFlg == "1")
            //{
            //    if (txtBox1.Text != "")
            //    {

            //        decimal total = Convert.ToDecimal(txtBox1.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

            //        if (total != 0)
            //        {
            //            total = total * Convert.ToDecimal(Persent);
            //            objBillingInvoice.ssi_feeamount1 = new CrmMoney();
            //            objBillingInvoice.ssi_feeamount1.Value = Convert.ToDecimal(total);
            //        }
            //    }
            //    else
            //    {
            //        objBillingInvoice.ssi_feeamount1 = new CrmMoney();
            //        objBillingInvoice.ssi_feeamount1.IsNull = true;
            //        objBillingInvoice.ssi_feeamount1.IsNullSpecified = true;
            //    }
            //}

            //if (ssi_SetUpFeeFlg == "1")
            //{

            //    if (txtBox2.Text != "")
            //    {
            //        decimal total = Convert.ToDecimal(txtBox2.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //        if (total != 0)
            //        {
            //            total = total * Convert.ToDecimal(Persent);

            //            objBillingInvoice.ssi_feeamount2 = new CrmMoney();
            //            objBillingInvoice.ssi_feeamount2.Value = Convert.ToDecimal(total);
            //        }
            //    }
            //    else
            //    {
            //        objBillingInvoice.ssi_feeamount2 = new CrmMoney();
            //        objBillingInvoice.ssi_feeamount2.IsNull = true;
            //        objBillingInvoice.ssi_feeamount2.IsNullSpecified = true;
            //    }
            //}
            //if (ssi_OtherFeeFlg == "1")
            //{
            //    if (txtBox3.Text != "")
            //    {

            //        decimal total = Convert.ToDecimal(txtBox3.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //        if (total != 0)
            //        {
            //            total = total * Convert.ToDecimal(Persent);

            //            objBillingInvoice.ssi_feeamount3 = new CrmMoney();
            //            objBillingInvoice.ssi_feeamount3.Value = Convert.ToDecimal(total);
            //        }
            //        else
            //        {
            //            objBillingInvoice.ssi_feeamount3 = new CrmMoney();
            //            objBillingInvoice.ssi_feeamount3.IsNull = true;
            //            objBillingInvoice.ssi_feeamount3.IsNullSpecified = true;
            //        }

            //    }
            //    else
            //    {
            //        objBillingInvoice.ssi_feeamount3 = new CrmMoney();
            //        objBillingInvoice.ssi_feeamount3.IsNull = true;
            //        objBillingInvoice.ssi_feeamount3.IsNullSpecified = true;
            //    }
            //}

            //Month2 Fee
            //if (Convert.ToString(txtFeesPerMonth.Text) != "")
            //{
            //    objBillingInvoice.ssi_month2fee = new CrmMoney();
            //    objBillingInvoice.ssi_month2fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //}
            //else
            //{
            //    objBillingInvoice.ssi_month2fee = new CrmMoney();
            //    objBillingInvoice.ssi_month2fee.IsNull = true;
            //    objBillingInvoice.ssi_month2fee.IsNullSpecified = true;
            //}

            ////Month3 Fee
            //if (Convert.ToString(txtFeesPerMonth.Text) != "")
            //{
            //    objBillingInvoice.ssi_month3fee = new CrmMoney();
            //    objBillingInvoice.ssi_month3fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
            //}
            //else
            //{
            //    objBillingInvoice.ssi_month3fee = new CrmMoney();
            //    objBillingInvoice.ssi_month3fee.IsNull = true;
            //    objBillingInvoice.ssi_month3fee.IsNullSpecified = true;
            //}
            #endregion


            System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");

            DateTime dtAUMdate = Convert.ToDateTime(txtAUMDate.Text);
            // DateTime dtAUMdate = DateTime.ParseExact(txtAUMDate.Text, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            string Month1 = dtAUMdate.AddMonths(2).ToString("MMMM");//full month can be assigned--->"january"
            string Month2 = dtAUMdate.AddMonths(3).ToString("MMMM");
            string Month3 = dtAUMdate.AddMonths(4).ToString("MMMM");

            //objBillingInvoice.ssi_month1 = Month1;
            //objBillingInvoice.ssi_month2 = Month2;
            //objBillingInvoice.ssi_month3 = Month3;
            objBillingInvoice["ssi_month1"] = Month1;
            objBillingInvoice["ssi_month2"] = Month2;
            objBillingInvoice["ssi_month3"] = Month3;
            //ssi_AdjustedFee
            if (Convert.ToString(txtAdjQtrFee.Text) != "")
            {
                decimal total = Convert.ToDecimal(txtAdjQtrFee.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_adjustedfee = new CrmMoney();
                //objBillingInvoice.ssi_adjustedfee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_adjustedfee"] = new Money(Convert.ToDecimal(total));
            }
            else
            {
                //objBillingInvoice.ssi_adjustedfee = new CrmMoney();
                //objBillingInvoice.ssi_adjustedfee.IsNull = true;
                //objBillingInvoice.ssi_adjustedfee.IsNullSpecified = true;
                objBillingInvoice["ssi_adjustedfee"] = null;
            }


            ////ssi_adjustment
            if (Convert.ToString(txtAdjAmt.Text) != "")
            {
                decimal total = Convert.ToDecimal(txtAdjAmt.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_adjustment = new CrmMoney();
                //objBillingInvoice.ssi_adjustment.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_adjustment"] = new Money(Convert.ToDecimal(total));
            }
            else
            {
                //objBillingInvoice.ssi_adjustment = new CrmMoney();
                //objBillingInvoice.ssi_adjustment.IsNull = true;
                //objBillingInvoice.ssi_adjustment.IsNullSpecified = true;
                objBillingInvoice["ssi_adjustment"] = null;
            }

            ////ssi_adjustmentreason
            if (Convert.ToString(txtAdjReason.Text) != "")
            {
                //objBillingInvoice.ssi_adjustmentreason = Convert.ToString(txtAdjReason.Text);
                objBillingInvoice["ssi_adjustmentreason"] = Convert.ToString(txtAdjReason.Text);
            }
            else
            {
                //objBillingInvoice.ssi_adjustmentreason = "";
                objBillingInvoice["ssi_adjustmentreason"] = "";
            }

            ////ssi_feeschedultype
            if (Convert.ToString(ddlClientType.SelectedValue) != "")
            {
                //objBillingInvoice.ssi_feescheduletype = new Picklist();
                ////objBillingInvoice.ssi_feescheduleid.type = EntityName.ssi_feeschedule.ToString();
                //objBillingInvoice.ssi_feescheduletype.Value = Convert.ToInt32(ddlClientType.SelectedValue);
                objBillingInvoice["ssi_feescheduletype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ddlClientType.SelectedValue));
            }

            //ssi_relationshipfee
            //if (Convert.ToString(txtRelationshipFee.Text) != "")
            //{
            //    objBillingInvoice.ssi_relationshipfee = new CrmMoney();
            //    objBillingInvoice.ssi_relationshipfee.Value = Convert.ToDecimal(txtRelationshipFee.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

            //}
            //else
            //{
            //    objBillingInvoice.ssi_relationshipfee = new CrmMoney();
            //    objBillingInvoice.ssi_relationshipfee.IsNull = true;
            //    objBillingInvoice.ssi_relationshipfee.IsNullSpecified = true;
            //}




            //ssi_customfee
            if (Convert.ToString(txtCustFeeAmount.Text) != "")
            {
                decimal total = Convert.ToDecimal(txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
                total = total * Convert.ToDecimal(Persent);

                //objBillingInvoice.ssi_customfee = new CrmMoney();
                //objBillingInvoice.ssi_customfee.Value = Convert.ToDecimal(total);
                objBillingInvoice["ssi_customfee"] = new Money(Convert.ToDecimal(total));
            }
            else
            {
                //objBillingInvoice.ssi_customfee = new CrmMoney();
                //objBillingInvoice.ssi_customfee.IsNull = true;
                //objBillingInvoice.ssi_customfee.IsNullSpecified = true;
                objBillingInvoice["ssi_customfee"] = null;
            }

            ////discount
            if (Convert.ToString(txtDiscount.Text) != "")
            {
                //objBillingInvoice.ssi_discount = new CrmDecimal();
                //objBillingInvoice.ssi_discount.Value = Convert.ToDecimal(txtDiscount.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "")) / 100;
                objBillingInvoice["ssi_discount"] = Convert.ToDecimal(txtDiscount.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "")) / 100;
            }
            else
            {
                //objBillingInvoice.ssi_discount = new CrmDecimal();
                //objBillingInvoice.ssi_discount.IsNull = true;
                //objBillingInvoice.ssi_discount.IsNullSpecified = true;
                objBillingInvoice["ssi_discount"] = null;
            }


            if (chkAccured.Checked)
            {
                //objBillingInvoice.ssi_accrued = new CrmBoolean();
                //objBillingInvoice.ssi_accrued.Value = true;
                objBillingInvoice["ssi_accrued"] = true;
            }
            else
            {
                //objBillingInvoice.ssi_accrued.Value = false;
                objBillingInvoice["ssi_accrued"] = false;
            }
            #region added 3_25_2019 Standard Min-Max Change

            string bpsfeeValue = string.Empty;
            string MinValue = string.Empty;
            string MaxValue = string.Empty;
            try
            {
                bpsfeeValue = ViewState["bpsfeeValue"].ToString();
            }
            catch (Exception ex)
            {
                bpsfeeValue = "";
            }
            try
            {
                MinValue = ViewState["MinValue"].ToString();
            }
            catch (Exception ex1)
            {
                MinValue = "";
            }
            try
            {
                MaxValue = ViewState["MaxValue"].ToString();
            }
            catch (Exception ex2)
            {
                MaxValue = "";
            }



            //Fee on first $25MM in bps - added 3_25_2019 Standard Min-Max Change
            if (bpsfeeValue != "")
            {
                string fee = bpsfeeValue.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace("$", "");
                if (fee != "NA")
                {
                    objBillingInvoice["ssi_feeonfirst25mminbps"] = Convert.ToDecimal(fee);
                }
                else
                {
                    objBillingInvoice["ssi_feeonfirst25mminbps"] = null;
                }
            }
            else
            {
                objBillingInvoice["ssi_feeonfirst25mminbps"] = null;
            }
            //Minimum fee in $ - added 3_25_2019 Standard Min-Max Change
            if (txtMinVal.Visible == true || txtMinVal.Text != "")
            {

                if (MinValue != "")
                {
                    string fees = MinValue.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "");
                    fees = Convert.ToDecimal(fees).ToString();

                    objBillingInvoice["ssi_minimumfeein"] = new Money(Convert.ToDecimal(fees));
                }
            }
            else
            {
                objBillingInvoice["ssi_minimumfeein"] = null;
            }


            //Maximum fee as a %- added 3_25_2019 Standard Min-Max Change
            if (txtMaxVal.Visible == true || txtMaxVal.Text != "")
            {
                if (MaxValue != "")
                {
                    string fee = MaxValue.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "").Replace("$", "");
                    if (fee != "NA")
                    {
                        objBillingInvoice["ssi_maximumfeeasa"] = Convert.ToDecimal(fee);
                    }
                    else
                    {
                        objBillingInvoice["ssi_maximumfeeasa"] = null;
                    }
                }
            }
            else
            {
                objBillingInvoice["ssi_maximumfeeasa"] = null;
            }
            if (ddlClientType.SelectedValue == "100000009")//  Standard Fee or $180K(greater of the two)// if (ddlClientType.SelectedValue == "100000000")
            {
                //Alt Fee Schedule Type- added 3_25_2019 Standard Min-Max Change
                if (bpsfeeValue != "" && MinValue != "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Discounted Standard (" + bpsfeeValue + " bps first $25mil) or at least " + MinValue + " not to exceed a maximum of " + MaxValue;
                }
                else if (bpsfeeValue == "" && MinValue != "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Standard Fee or a minimum " + MinValue + ", not to exceed a max of " + MaxValue + " of AUM";
                }
                else if (bpsfeeValue == "" && MinValue == "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Standard Fee or the lesser of " + MaxValue;
                }
                else if (bpsfeeValue == "" && MinValue != "" && MaxValue == "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Standard Fee or the greater of " + MinValue;
                }
                else if (bpsfeeValue != "" && MinValue == "" && MaxValue == "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = " Discounted Standard(" + bpsfeeValue + " bps first $25mil)";
                }
                else if (bpsfeeValue != "" && MinValue != "" && MaxValue == "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Discounted Standard (" + bpsfeeValue + " bps first $25mil), Standard Fee or the greater of " + MinValue;
                }
                else if (bpsfeeValue != "" && MinValue == "" && MaxValue != "")
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = "Discounted Standard (" + bpsfeeValue + " bps first $25mil), Standard Fee or the lesser of " + MaxValue;
                }
                else
                {
                    objBillingInvoice["ssi_altfeescheduletype"] = null;
                }
            }
            else
            {
                objBillingInvoice["ssi_altfeescheduletype"] = null;
            }
            #endregion


            //added 7_15_2019
            if (ssi_FeePctAmt != "")
            {
                objBillingInvoice["ssi_securityfeeaum"] = new Money(Convert.ToDecimal(ssi_FeePctAmt));
            }
            else
            {
                objBillingInvoice["ssi_securityfeeaum"] = null;
            }
            if (ssi_ScheduleAUM != "")
            {
                objBillingInvoice["ssi_scheduleaum"] = new Money(Convert.ToDecimal(ssi_ScheduleAUM));
            }
            else
            {
                objBillingInvoice["ssi_scheduleaum"] = null;
            }

            service.Update(objBillingInvoice);


            System.Web.UI.WebControls.ListItem itemToRemove = ddlClientType.Items.FindByValue("0");
            if (itemToRemove != null)
            {
                ddlClientType.Items.Remove(itemToRemove);
            }

            lblMessage.Text = "Records Updated Successfully";
        }
        catch (Exception ex)
        {
            lblMessage.Text = "Billing Invoice Failed to Update : " + ex.ToString();
        }
    }

    public void InsertFlatFee()
    {
        DB clsDB = new DB();//class Library 
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
        //  DataSet dsData = clsDB.getDataSet("SP_U_BILLINGFEERATE_Test @FlatFeeFlg=1 ,@HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");
        DataSet dsData = clsDB.getDataSet("SP_U_BILLINGFEERATE @FlatFeeFlg=1 ,@HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");
        #region crm Connection
        string orgName = "GreshamPartners";

        //CRM SERVICE
        IOrganizationService service = null;
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string strDescription = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion

        try
        {
            if (dsData.Tables[0].Rows != null)
            {
                foreach (DataRow row in dsData.Tables[0].Rows)
                {
                    string ssi_allocation = row["AllocatedFlg"].ToString();
                    bool Flag = false;
                    if (ssi_allocation == "0")
                        Flag = false;
                    else
                        Flag = true;

                    string ssi_BillingInvoiceId = row["Ssi_billingInvoiceId"].ToString();
                    string ssi_FeeTypID = row["FeeTypID"].ToString();
                    string ssi_FeeAmt = row["FeesAmt"].ToString();
                    decimal FeesAmount = Convert.ToDecimal(ssi_FeeAmt);

                    // ssi_billinginvoiceflatfee objInvoicFlatFee = new ssi_billinginvoiceflatfee();
                    Entity objInvoicFlatFee = new Entity("ssi_billinginvoiceflatfee");

                    //string BillingFlatFeeId = Guid.NewGuid().ToString();

                    //objInvoicFlatFee.ssi_billinginvoiceflatfeeid = new Key();
                    //objInvoicFlatFee.ssi_billinginvoiceflatfeeid.Value = Guid.NewGuid();

                    //objInvoicFlatFee.ssi_billinginvoiceid = new Lookup();
                    //objInvoicFlatFee.ssi_billinginvoiceid.type = EntityName.ssi_billinginvoice.ToString();
                    //objInvoicFlatFee.ssi_billinginvoiceid.Value = new Guid(Convert.ToString(ssi_BillingInvoiceId));
                    objInvoicFlatFee["ssi_billinginvoiceid"] = new EntityReference("ssi_billinginvoice", new Guid(Convert.ToString(ssi_BillingInvoiceId)));

                    //objInvoicFlatFee.ssi_allocated = new CrmBoolean();
                    //objInvoicFlatFee.ssi_allocated.Value = Flag;
                    objInvoicFlatFee["ssi_allocated"] = Convert.ToBoolean(Flag);

                    //objInvoicFlatFee.ssi_flatfeetype = new Picklist();
                    //objInvoicFlatFee.ssi_flatfeetype.Value = Convert.ToInt32(ssi_FeeTypID);
                    objInvoicFlatFee["ssi_flatfeetype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ssi_FeeTypID));

                    //objInvoicFlatFee.ssi_amount = new CrmMoney();
                    //objInvoicFlatFee.ssi_amount.Value = FeesAmount;
                    objInvoicFlatFee["ssi_amount"] = new Money(FeesAmount);

                    ////objInvoicFlatFee.ssi_amount = new CrmMoney();
                    ////objInvoicFlatFee.ssi_amount.Value = 10;

                    service.Create(objInvoicFlatFee);

                }
            }
        }
        catch (Exception e)
        {

        }

    }

    public void insertCustomefee(string vGuID = null)  // for update only
    {
        #region crm Connection
        string orgName = "GreshamPartners";

        //CRM SERVICE
        IOrganizationService service = null;
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string strDescription = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion

        try
        {
            DataTable dtDataInsert = (DataTable)ViewState["CustomeFeeInsert"];
            if (dtDataInsert != null)
            {
                foreach (DataRow row in dtDataInsert.Rows)
                {
                    string BillingInvoiceID = null;
                    string Name = row["ssi_name"].ToString();
                    string FeePercent = row["ssi_FeePercent"].ToString();
                    string amount = row["ssi_amount"].ToString();

                    if (vGuID != null)
                    {
                        BillingInvoiceID = vGuID;
                    }
                    else
                    {
                        BillingInvoiceID = row["ssi_billinginvoiceID"].ToString();
                    }

                    //ssi_billingcustomfee ObjBillingcustomfee = new ssi_billingcustomfee();
                    Entity ObjBillingcustomfee = new Entity("ssi_billingcustomfee");
                    // CustomeID
                    //ObjBillingcustomfee.ssi_billingcustomfeeid = new Key();
                    //ObjBillingcustomfee.ssi_billingcustomfeeid.Value = Guid.NewGuid();
                    //ObjBillingcustomfee["ssi_billingcustomfeeid"] =

                    // billing invoiceID
                    //ObjBillingcustomfee.ssi_billinginvoiceid = new Lookup();
                    //ObjBillingcustomfee.ssi_billinginvoiceid.type = EntityName.ssi_billinginvoice.ToString();
                    //ObjBillingcustomfee.ssi_billinginvoiceid.Value = new Guid(Convert.ToString(BillingInvoiceID));
                    ObjBillingcustomfee["ssi_billinginvoiceid"] = new EntityReference("ssi_billinginvoice", new Guid(Convert.ToString(BillingInvoiceID)));

                    // name
                    //ObjBillingcustomfee.ssi_name = Name;
                    ObjBillingcustomfee["ssi_name"] = Name;

                    //Percent
                    //ObjBillingcustomfee.ssi_feepercent = new CrmDecimal();
                    //ObjBillingcustomfee.ssi_feepercent.Value = Convert.ToDecimal(FeePercent);
                    ObjBillingcustomfee["ssi_feepercent"] = Convert.ToDecimal(FeePercent);

                    // amount
                    //ObjBillingcustomfee.ssi_amount = new CrmMoney();
                    //ObjBillingcustomfee.ssi_amount.Value = Convert.ToDecimal(amount);
                    ObjBillingcustomfee["ssi_amount"] = new Money(Convert.ToDecimal(amount));

                    service.Create(ObjBillingcustomfee);
                }
                if (vGuID == null)
                {
                    ViewState["CustomeFeeInsert"] = null;
                }
            }
        }
        catch
        {

        }


        //objBillingInvoice.ssi_billingprimaryid = new Lookup();
        //objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
        //objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(BillingUUID));

    }
    public void insertCustomefeeNew(string vGuID = null, string BillingUUID = null)
    {
        #region crm Connection
        string orgName = "GreshamPartners";

        //CRM SERVICE
        IOrganizationService service = null;
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string strDescription = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion

        try
        {
            DataTable dtDataInsert = (DataTable)ViewState["CustomeFeeInsert"];
            if (dtDataInsert != null)
            {
                foreach (DataRow row in dtDataInsert.Rows)
                {
                    string BillingInvoiceID = null;
                    string Name = row["ssi_name"].ToString();
                    string FeePercent = row["ssi_FeePercent"].ToString();
                    string amount = row["ssi_amount"].ToString();
                    string Bilid = row["Ssi_billingID"].ToString();
                    if (vGuID != null)
                    {
                        BillingInvoiceID = vGuID;
                    }
                    else
                    {
                        BillingInvoiceID = row["ssi_billinginvoiceID"].ToString();
                    }

                    if (Bilid == BillingUUID)
                    {
                        //ssi_billingcustomfee ObjBillingcustomfee = new ssi_billingcustomfee();
                        Entity ObjBillingcustomfee = new Entity("ssi_billingcustomfee");

                        // CustomeID
                        //ObjBillingcustomfee.ssi_billingcustomfeeid = new Key();
                        //ObjBillingcustomfee.ssi_billingcustomfeeid.Value = Guid.NewGuid();

                        // billing invoiceID
                        //ObjBillingcustomfee.ssi_billinginvoiceid = new Lookup();
                        //ObjBillingcustomfee.ssi_billinginvoiceid.type = EntityName.ssi_billinginvoice.ToString();
                        //ObjBillingcustomfee.ssi_billinginvoiceid.Value = new Guid(Convert.ToString(BillingInvoiceID));
                        ////ObjBillingcustomfee.ssi_billingcustomfeeid.Value = Guid.NewGuid();
                        ObjBillingcustomfee["ssi_billinginvoiceid"] = new EntityReference("ssi_billinginvoice", new Guid(Convert.ToString(BillingInvoiceID)));

                        // name
                        //ObjBillingcustomfee.ssi_name = Name;
                        ObjBillingcustomfee["ssi_name"] = Name;

                        //Percent
                        //ObjBillingcustomfee.ssi_feepercent = new CrmDecimal();
                        //ObjBillingcustomfee.ssi_feepercent.Value = Convert.ToDecimal(FeePercent);
                        ObjBillingcustomfee["ssi_feepercent"] = Convert.ToDecimal(FeePercent);

                        // amount
                        //ObjBillingcustomfee.ssi_amount = new CrmMoney();
                        //ObjBillingcustomfee.ssi_amount.Value = Convert.ToDecimal(amount);
                        ObjBillingcustomfee["ssi_amount"] = new Money(Convert.ToDecimal(amount));

                        service.Create(ObjBillingcustomfee);
                    }
                }
                if (vGuID == null)
                {
                    ViewState["CustomeFeeInsert"] = null;
                }
            }
        }
        catch
        {

        }


        //objBillingInvoice.ssi_billingprimaryid = new Lookup();
        //objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
        //objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(BillingUUID));

    }
    public void deleteCustomeFee(String UID)
    {
        #region crm Connection
        string orgName = "GreshamPartners";

        //CRM SERVICE
        IOrganizationService service = null;
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string strDescription = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion

        try
        {
            //ssi_billingcustomfee ObjBillingcustomfee = new ssi_billingcustomfee();

            //ObjBillingcustomfee.ssi_billingcustomfeeid = new Key();
            //ObjBillingcustomfee.ssi_billingcustomfeeid.Value = new Guid(UID);

            Guid gUId = new Guid(UID);


            service.Delete("ssi_billingcustomfee", gUId);

        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            Response.Write(ex.StackTrace);

        }
    }
    public void deleteInvoiceFee(String UID)
    {
        #region crm Connection
        string orgName = "GreshamPartners";

        //CRM SERVICE
        IOrganizationService service = null;
        // string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string strDescription = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
            Response.Write(exc.StackTrace + "<br/>" + exc.Message);
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
            Response.Write(exc.StackTrace + "<br/>" + exc.Message);
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion

        try
        {
            //ssi_billingcustomfee ObjBillingcustomfee = new ssi_billingcustomfee();

            //ObjBillingcustomfee.ssi_billingcustomfeeid = new Key();
            //ObjBillingcustomfee.ssi_billingcustomfeeid.Value = new Guid(UID);

            Guid gUId = new Guid(UID);


            //service.Delete(EntityName.ssi_billinginvoice.ToString(), gUId);
            service.Delete("ssi_billinginvoice", gUId);
        }
        catch (Exception exc)
        {
            Response.Write(exc.StackTrace + "<br/>" + exc.Message);
        }
    }
    public void deleteInvoiceFlatFee(String UID)
    {
        #region crm Connection
        string orgName = "GreshamPartners";

        //CRM SERVICE
        IOrganizationService service = null;
        // string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string strDescription = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion

        try
        {
            //ssi_billingcustomfee ObjBillingcustomfee = new ssi_billingcustomfee();

            //ObjBillingcustomfee.ssi_billingcustomfeeid = new Key();
            //ObjBillingcustomfee.ssi_billingcustomfeeid.Value = new Guid(UID);

            Guid gUId = new Guid(UID);


            // service.Delete(EntityName.ssi_billinginvoiceflatfee.ToString(), gUId);
            service.Delete("ssi_billinginvoiceflatfee", gUId);
        }
        catch (Exception exc)
        {
            Response.Write(exc.StackTrace + "<br/>" + exc.Message);
        }
    }
    public void UpdateNotes()    // update notes in family entity
    {
        #region crm Connection
        // string orgName = "GreshamPartners";

        //CRM SERVICE
        IOrganizationService service = null;
        //  string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string strDescription = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        #endregion

        DataTable dtData = (DataTable)ViewState["FamilyNotes"];
        if (dtData != null)
        {

            string FamilyUUID = dtData.Rows[0]["FamilyUUID"].ToString();
            string Notes = txtNotes.Text;
            string NotesTest = Notes.Replace(" ", "").Replace("\r\n", "");
            // sas_greshamfamily objfamily = new sas_greshamfamily();
            Entity objfamily = new Entity("sas_greshamfamily");


            //objfamily.sas_greshamfamilyid = new Key();
            //objfamily.sas_greshamfamilyid.Value = new Guid(FamilyUUID);
            objfamily["sas_greshamfamilyid"] = new Guid(FamilyUUID);

            if (NotesTest != "")
            {
                //  objfamily.ssi_advisorynotes = Notes;
                objfamily["ssi_advisorynotes"] = Notes;
            }
            else
            {
                objfamily["ssi_advisorynotes"] = null;
            }


            service.Update(objfamily);
        }
        //objBillingInvoice.ssi_billingprimaryid = new Lookup();
        //objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
        //objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(BillingUUID));

    }

    public void updateInvoicewithTotalandPer()
    {
        DB clsDB = new DB();//class Library 
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
        string sqlQuery = "SP_U_BILLINGFEERATE @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ";
        //DataSet dsData = clsDB.getDataSet("SP_U_BILLINGFEERATE_Test @HouseHoldUUID=" + HHValue + ",@AumAsodfDate='" + txtAUMDate.Text + "',@BillingForUUID='" + strBillFor + "' ");
        DataSet dsData = clsDB.getDataSet(sqlQuery);
        try
        {
            if (dsData != null)
            {
                foreach (DataRow row in dsData.Tables[0].Rows)
                {

                    #region crm Connection
                    //  string orgName = "GreshamPartners";

                    //CRM SERVICE
                    IOrganizationService service = null;
                    //  string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

                    string strDescription = "";
                    try
                    {
                        service = GM.GetCrmService();
                        strDescription = "Crm Service starts successfully";
                    }
                    catch (System.Web.Services.Protocols.SoapException exc)
                    {
                        strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                        lblMessage.Text = strDescription;
                    }
                    catch (Exception exc)
                    {
                        strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                        lblMessage.Text = strDescription;
                    }


                    try
                    {
                        //Response.Write(service.Url);
                        //service.PreAuthenticate = true;
                        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
                    }
                    catch (NullReferenceException ne)
                    {
                        //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
                    }

                    #endregion


                    string ssi_billingInvoiceID = row["ssi_billingInvoiceID"].ToString();
                    string ssi_TotalFee = row["ssi_TotalFee"].ToString();
                    string ssi_feerate = row["ssi_feerate"].ToString();
                    string Ssi_billingID = row["Ssi_billingID"].ToString();

                    decimal quterlyFee = Convert.ToDecimal(ssi_TotalFee) / 4;
                    decimal Monthly = Convert.ToDecimal(ssi_TotalFee) / 12;

                    string InvoiceLoadId = Guid.NewGuid().ToString();

                    // ssi_billinginvoice objBillingInvoice = new ssi_billinginvoice();
                    Entity objBillingInvoice = new Entity("ssi_billinginvoice");

                    // objBillingInvoice.ssi_invoiceloadid = InvoiceLoadId;
                    objBillingInvoice["ssi_invoiceloadid"] = InvoiceLoadId;

                    //objBillingInvoice.ssi_billinginvoiceid = new Key();
                    //objBillingInvoice.ssi_billinginvoiceid.Value = new Guid(ssi_billingInvoiceID);
                    objBillingInvoice["ssi_billinginvoiceid"] = new Guid(ssi_billingInvoiceID);

                    // ssi_billingid
                    if (Convert.ToString(Ssi_billingID) != "")
                    {
                        //objBillingInvoice.ssi_billingprimaryid = new Lookup();
                        //objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
                        //objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(Ssi_billingID));
                        objBillingInvoice["ssi_billingprimaryid"] = new EntityReference("ssi_billing", new Guid(Convert.ToString(Ssi_billingID)));

                    }

                    if (ssi_feerate != "")
                    {
                        //objBillingInvoice.ssi_feerate = new CrmDecimal();
                        //objBillingInvoice.ssi_feerate.Value = Convert.ToDecimal(ssi_feerate);
                        objBillingInvoice["ssi_feerate"] = Convert.ToDecimal(ssi_feerate);
                    }
                    else
                    {
                        //objBillingInvoice.ssi_feerate = new CrmDecimal();
                        //objBillingInvoice.ssi_feerate.IsNull = true;
                        //objBillingInvoice.ssi_feerate.IsNullSpecified = true;
                        objBillingInvoice["ssi_feerate"] = null;
                    }

                    decimal AnnualFee = Convert.ToDecimal(ssi_TotalFee);
                    //objBillingInvoice.ssi_totalannualfee = new CrmMoney();
                    //objBillingInvoice.ssi_totalannualfee.Value = AnnualFee;
                    objBillingInvoice["ssi_totalannualfee"] = new Money(AnnualFee);

                    //objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
                    //objBillingInvoice.ssi_quarterlyfee.Value = quterlyFee;
                    objBillingInvoice["ssi_quarterlyfee"] = new Money(quterlyFee);

                    //objBillingInvoice.ssi_month1fee = new CrmMoney();
                    //objBillingInvoice.ssi_month1fee.Value = Monthly;
                    objBillingInvoice["ssi_month1fee"] = new Money(Monthly);

                    //Month2 Fee
                    //objBillingInvoice.ssi_month2fee = new CrmMoney();
                    //objBillingInvoice.ssi_month2fee.Value = Monthly;
                    objBillingInvoice["ssi_month2fee"] = new Money(Monthly);
                    //Month3 Fee
                    //objBillingInvoice.ssi_month3fee = new CrmMoney();
                    //objBillingInvoice.ssi_month3fee.Value = Monthly;
                    objBillingInvoice["ssi_month3fee"] = new Money(Monthly);


                    service.Update(objBillingInvoice);

                }
            }
        }
        catch
        {

        }
    }
    #region old not used
    //Update Invoice old
    //protected void UpdateInvoice(string BillingInvoiceId, string BillingUUID)
    //{
    //    #region crm Connection
    //    try
    //    {
    //        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//set crmserver url
    //        string orgName = "GreshamPartners";
    //        CrmService service = null;
    //        string strDescription = "";

    //        try
    //        {
    //            service = GetCrmService(crmServerUrl, orgName);
    //            strDescription = "Crm Service starts successfully";
    //        }
    //        catch (System.Web.Services.Protocols.SoapException exc)
    //        {

    //            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
    //            lblMessage.Text = strDescription;
    //        }
    //        catch (Exception exc)
    //        {

    //            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
    //            lblMessage.Text = strDescription;
    //        }



    //        try
    //        {
    //            //Response.Write(service.Url);
    //            service.PreAuthenticate = true;
    //            service.Credentials = System.Net.CredentialCache.DefaultCredentials;
    //        }
    //        catch (NullReferenceException ne)
    //        {
    //            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
    //        }

    //    #endregion


    //        string InvoiceLoadId = Guid.NewGuid().ToString();

    //        ssi_billinginvoice objBillingInvoice = new ssi_billinginvoice();
    //        objBillingInvoice.ssi_invoiceloadid = InvoiceLoadId;
    //        objBillingInvoice.ssi_billinginvoiceid = new Key();
    //        objBillingInvoice.ssi_billinginvoiceid.Value = new Guid(BillingInvoiceId);


    //        //  objBillingInvoice.ssi_billinginvoiceid = new Key();
    //        // objBillingInvoice.ssi_billinginvoiceid.Value = Guid.NewGuid();

    //        //Name
    //        objBillingInvoice.ssi_name = ddlBillFor.SelectedItem.Text + "-" + txtAUMDate.Text;

    //        // ssi_billingid
    //        if (Convert.ToString(BillingUUID) != "")
    //        {
    //            objBillingInvoice.ssi_billingprimaryid = new Lookup();
    //            objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
    //            objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(BillingUUID));
    //        }

    //        //                //InvoiceId
    //        //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]) != "")
    //        //                {
    //        //                    objBillingInvoice.ssi_invoiceid = Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]);
    //        //                }


    //        ////BillingId
    //        //if (Convert.ToString(ViewState["BillingId"]) != "")
    //        //{
    //        //    objBillingInvoice.ssi_billingid = new CrmNumber();
    //        //    objBillingInvoice.ssi_billingid.Value = Convert.ToInt32(ViewState["BillingId"]);
    //        //}


    //        //AUM Date 
    //        if (txtAUMDate.Text != "")
    //        {
    //            objBillingInvoice.ssi_aumasofdate = new CrmDateTime();
    //            objBillingInvoice.ssi_aumasofdate.Value = Convert.ToString(txtAUMDate.Text);
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_aumasofdate.Value = "";
    //        }

    //        //Total AUM
    //        if (Convert.ToString(txtTotalAUM.Text) != "")
    //        {
    //            objBillingInvoice.ssi_totalaum = new CrmMoney();
    //            objBillingInvoice.ssi_totalaum.Value = Convert.ToDecimal(txtTotalAUM.Text.Replace("$", "").Replace(",", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_totalaum = new CrmMoney();
    //            objBillingInvoice.ssi_totalaum.IsNull = true;
    //            objBillingInvoice.ssi_totalaum.IsNullSpecified = true;
    //        }

    //        //Billing AUM
    //        if (Convert.ToString(txtTotalAUM.Text) != "")
    //        {
    //            objBillingInvoice.ssi_aum = new CrmMoney();
    //            objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(txtBillingAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_aum = new CrmMoney();
    //            objBillingInvoice.ssi_aum.IsNull = true;
    //            objBillingInvoice.ssi_aum.IsNullSpecified = true;

    //        }

    //        if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
    //        {
    //            objBillingInvoice.ssi_annualfee = new CrmMoney();
    //            objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            //Annual Fee == isnull(custom fee, standard annual fee calc)
    //            if (Convert.ToString(txtCustFeeAmount.Text) != "" && Convert.ToString(txtCustFeeAmount.Text).Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "") != "0.00")
    //            {
    //                objBillingInvoice.ssi_annualfee = new CrmMoney();
    //                objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //            }
    //            else
    //            {
    //                objBillingInvoice.ssi_annualfee = new CrmMoney();
    //                objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //            }

    //        }

    //        //Fee Rate 
    //        if (Convert.ToString(txtFeeRateCalc.Text) != "")
    //        {
    //            objBillingInvoice.ssi_feerate = new CrmDecimal();
    //            objBillingInvoice.ssi_feerate.Value = Convert.ToDecimal(txtFeeRateCalc.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_feerate = new CrmDecimal();
    //            objBillingInvoice.ssi_feerate.IsNull = true;
    //            objBillingInvoice.ssi_feerate.IsNullSpecified = true;

    //        }

    //        //Quarterly Fee 
    //        if (Convert.ToString(txtQuaterlyFeeCalc.Text) != "")
    //        {
    //            objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
    //            objBillingInvoice.ssi_quarterlyfee.Value = Convert.ToDecimal(txtQuaterlyFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
    //            objBillingInvoice.ssi_quarterlyfee.IsNull = true;
    //            objBillingInvoice.ssi_quarterlyfee.IsNullSpecified = true;

    //        }

    //        //Month1 Fee
    //        if (Convert.ToString(txtFeesPerMonth.Text) != "")
    //        {
    //            objBillingInvoice.ssi_month1fee = new CrmMoney();
    //            objBillingInvoice.ssi_month1fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_month1fee = new CrmMoney();
    //            objBillingInvoice.ssi_month1fee.IsNull = true;
    //            objBillingInvoice.ssi_month1fee.IsNullSpecified = true;
    //        }

    //        //Month2 Fee
    //        if (Convert.ToString(txtFeesPerMonth.Text) != "")
    //        {
    //            objBillingInvoice.ssi_month2fee = new CrmMoney();
    //            objBillingInvoice.ssi_month2fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_month2fee = new CrmMoney();
    //            objBillingInvoice.ssi_month2fee.IsNull = true;
    //            objBillingInvoice.ssi_month2fee.IsNullSpecified = true;
    //        }

    //        //Month3 Fee
    //        if (Convert.ToString(txtFeesPerMonth.Text) != "")
    //        {
    //            objBillingInvoice.ssi_month3fee = new CrmMoney();
    //            objBillingInvoice.ssi_month3fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_month3fee = new CrmMoney();
    //            objBillingInvoice.ssi_month3fee.IsNull = true;
    //            objBillingInvoice.ssi_month3fee.IsNullSpecified = true;
    //        }


    //        System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");

    //        DateTime dtAUMdate = Convert.ToDateTime(txtAUMDate.Text);
    //        // DateTime dtAUMdate = DateTime.ParseExact(txtAUMDate.Text, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
    //        string Month1 = dtAUMdate.AddMonths(2).ToString("MMMM");//full month can be assigned--->"january"
    //        string Month2 = dtAUMdate.AddMonths(3).ToString("MMMM");
    //        string Month3 = dtAUMdate.AddMonths(4).ToString("MMMM");

    //        objBillingInvoice.ssi_month1 = Month1;
    //        objBillingInvoice.ssi_month2 = Month2;
    //        objBillingInvoice.ssi_month3 = Month3;


    //        //ssi_AdjustedFee
    //        if (Convert.ToString(txtAdjQtrFee.Text) != "")
    //        {
    //            objBillingInvoice.ssi_adjustedfee = new CrmMoney();
    //            objBillingInvoice.ssi_adjustedfee.Value = Convert.ToDecimal(txtAdjQtrFee.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_adjustedfee = new CrmMoney();
    //            objBillingInvoice.ssi_adjustedfee.IsNull = true;
    //            objBillingInvoice.ssi_adjustedfee.IsNullSpecified = true;
    //        }


    //        //ssi_adjustment
    //        if (Convert.ToString(txtAdjAmt.Text) != "")
    //        {
    //            objBillingInvoice.ssi_adjustment = new CrmMoney();
    //            objBillingInvoice.ssi_adjustment.Value = Convert.ToDecimal(txtAdjAmt.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_adjustment = new CrmMoney();
    //            objBillingInvoice.ssi_adjustment.IsNull = true;
    //            objBillingInvoice.ssi_adjustment.IsNullSpecified = true;
    //        }

    //        //ssi_adjustmentreason
    //        if (Convert.ToString(txtAdjReason.Text) != "")
    //        {
    //            objBillingInvoice.ssi_adjustmentreason = Convert.ToString(txtAdjReason.Text);
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_adjustmentreason = "";
    //        }

    //        //ssi_feeschedultype
    //        if (Convert.ToString(ddlClientType.SelectedValue) != "")
    //        {
    //            objBillingInvoice.ssi_feescheduletype = new Picklist();
    //            //objBillingInvoice.ssi_feescheduleid.type = EntityName.ssi_feeschedule.ToString();
    //            objBillingInvoice.ssi_feescheduletype.Value = Convert.ToInt32(ddlClientType.SelectedValue);
    //        }

    //        //ssi_relationshipfee
    //        if (Convert.ToString(txtRelationshipFee.Text) != "")
    //        {
    //            objBillingInvoice.ssi_relationshipfee = new CrmMoney();
    //            objBillingInvoice.ssi_relationshipfee.Value = Convert.ToDecimal(txtRelationshipFee.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_relationshipfee = new CrmMoney();
    //            objBillingInvoice.ssi_relationshipfee.IsNull = true;
    //            objBillingInvoice.ssi_relationshipfee.IsNullSpecified = true;
    //        }

    //        //ssi_customfee
    //        if (Convert.ToString(txtCustFeeAmount.Text) != "")
    //        {
    //            objBillingInvoice.ssi_customfee = new CrmMoney();
    //            objBillingInvoice.ssi_customfee.Value = Convert.ToDecimal(txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_customfee = new CrmMoney();
    //            objBillingInvoice.ssi_customfee.IsNull = true;
    //            objBillingInvoice.ssi_customfee.IsNullSpecified = true;
    //        }

    //        //discount
    //        if (Convert.ToString(txtDiscount.Text) != "")
    //        {
    //            objBillingInvoice.ssi_discount = new CrmDecimal();
    //            objBillingInvoice.ssi_discount.Value = Convert.ToDecimal(txtDiscount.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "")) / 100;
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_discount = new CrmDecimal();
    //            objBillingInvoice.ssi_discount.IsNull = true;
    //            objBillingInvoice.ssi_discount.IsNullSpecified = true;
    //        }

    //        service.Update(objBillingInvoice);

    //        System.Web.UI.WebControls.ListItem itemToRemove = ddlClientType.Items.FindByValue("0");
    //        if (itemToRemove != null)
    //        {
    //            ddlClientType.Items.Remove(itemToRemove);
    //        }

    //        lblMessage.Text = "Records Updated Successfully";
    //    }
    //    catch (Exception ex)
    //    {
    //        lblMessage.Text = "Billing Invoice Failed to Update : " + ex.ToString();
    //    }
    //}
    //Crm Service
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
    /* Creating Invoice*/
    //protected void CreateInvoice(string BillingUUID)
    //{
    //    try
    //    {

    //        string orgName = "GreshamPartners";

    //        //CRM SERVICE
    //        CrmService service = null;
    //        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

    //        string strDescription = "";
    //        try
    //        {
    //            service = GetCrmService(crmServerUrl, orgName);
    //            strDescription = "Crm Service starts successfully";
    //        }
    //        catch (System.Web.Services.Protocols.SoapException exc)
    //        {
    //            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
    //            lblMessage.Text = strDescription;
    //        }
    //        catch (Exception exc)
    //        {
    //            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
    //            lblMessage.Text = strDescription;
    //        }


    //        try
    //        {
    //            //Response.Write(service.Url);
    //            service.PreAuthenticate = true;
    //            service.Credentials = System.Net.CredentialCache.DefaultCredentials;
    //        }
    //        catch (NullReferenceException ne)
    //        {
    //            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
    //        }

    //        string InvoiceLoadId = Guid.NewGuid().ToString();
    //        ssi_billinginvoice objBillingInvoice = new ssi_billinginvoice();
    //        objBillingInvoice.ssi_invoiceloadid = InvoiceLoadId;

    //        objBillingInvoice.ssi_billinginvoiceid = new Key();
    //        objBillingInvoice.ssi_billinginvoiceid.Value = Guid.NewGuid();


    //        //  objBillingInvoice.ssi_billinginvoiceid = new Key();
    //        // objBillingInvoice.ssi_billinginvoiceid.Value = Guid.NewGuid();

    //        //Name
    //        objBillingInvoice.ssi_name = ddlBillFor.SelectedItem.Text + "-" + txtAUMDate.Text;

    //        // ssi_billingid
    //        if (Convert.ToString(BillingUUID) != "")
    //        {
    //            objBillingInvoice.ssi_billingprimaryid = new Lookup();
    //            objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
    //            objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(BillingUUID));
    //        }

    //        //                //InvoiceId
    //        //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]) != "")
    //        //                {
    //        //                    objBillingInvoice.ssi_invoiceid = Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]);
    //        //                }


    //        //BillingId
    //        /* */
    //        if (Convert.ToString(ViewState["BillingId"]) != "")
    //        {
    //            objBillingInvoice.ssi_billingid = new CrmNumber();   //ssi_billingid-->value comes from Crm
    //            objBillingInvoice.ssi_billingid.Value = Convert.ToInt32(ViewState["BillingId"]);
    //        }


    //        //AUM Date 
    //        if (txtAUMDate.Text != "")
    //        {
    //            objBillingInvoice.ssi_aumasofdate = new CrmDateTime();//ssi_aumasofdate-->value comes from Crm
    //            objBillingInvoice.ssi_aumasofdate.Value = Convert.ToString(txtAUMDate.Text);
    //        }

    //        //Total AUM
    //        if (Convert.ToString(txtTotalAUM.Text) != "")
    //        {
    //            objBillingInvoice.ssi_totalaum = new CrmMoney();//ssi_totalaum-->value comes from Crm
    //            objBillingInvoice.ssi_totalaum.Value = Convert.ToDecimal(txtTotalAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        //Billing AUM
    //        if (Convert.ToString(txtTotalAUM.Text) != "")
    //        {
    //            objBillingInvoice.ssi_aum = new CrmMoney();//ssi_aum-->value comes from Crm
    //            objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(txtBillingAUM.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
    //        {
    //            objBillingInvoice.ssi_annualfee = new CrmMoney();
    //            objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }
    //        else
    //        {
    //            //Annual Fee == isnull(custom fee, standard annual fee calc)
    //            if (Convert.ToString(txtCustFeeAmount.Text) != "" && Convert.ToString(txtCustFeeAmount.Text).Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "") != "0.00")
    //            {
    //                objBillingInvoice.ssi_annualfee = new CrmMoney();
    //                objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //            }
    //            else
    //            {
    //                objBillingInvoice.ssi_annualfee = new CrmMoney();
    //                objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //            }

    //        }

    //        //Fee Rate 
    //        if (Convert.ToString(txtFeeRateCalc.Text) != "")
    //        {
    //            objBillingInvoice.ssi_feerate = new CrmDecimal();
    //            objBillingInvoice.ssi_feerate.Value = Convert.ToDecimal(txtFeeRateCalc.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        //Quarterly Fee 
    //        if (Convert.ToString(txtQuaterlyFeeCalc.Text) != "")
    //        {
    //            objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
    //            objBillingInvoice.ssi_quarterlyfee.Value = Convert.ToDecimal(txtQuaterlyFeeCalc.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        //Month1 Fee
    //        if (Convert.ToString(txtFeesPerMonth.Text) != "")
    //        {
    //            objBillingInvoice.ssi_month1fee = new CrmMoney();
    //            objBillingInvoice.ssi_month1fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        //Month2 Fee
    //        if (Convert.ToString(txtFeesPerMonth.Text) != "")
    //        {
    //            objBillingInvoice.ssi_month2fee = new CrmMoney();
    //            objBillingInvoice.ssi_month2fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        //Month3 Fee
    //        if (Convert.ToString(txtFeesPerMonth.Text) != "")
    //        {
    //            objBillingInvoice.ssi_month3fee = new CrmMoney();
    //            objBillingInvoice.ssi_month3fee.Value = Convert.ToDecimal(txtFeesPerMonth.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");

    //        DateTime dtAUMdate = Convert.ToDateTime(txtAUMDate.Text);
    //        //  DateTime dtAUMdate = DateTime.ParseExact(txtAUMDate.Text, "MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
    //        string Month1 = dtAUMdate.AddMonths(2).ToString("MMMM");//full month can be assigned--->"January"
    //        string Month2 = dtAUMdate.AddMonths(3).ToString("MMMM");
    //        string Month3 = dtAUMdate.AddMonths(4).ToString("MMMM");

    //        objBillingInvoice.ssi_month1 = Month1;
    //        objBillingInvoice.ssi_month2 = Month2;
    //        objBillingInvoice.ssi_month3 = Month3;


    //        //ssi_AdjustedFee
    //        if (Convert.ToString(txtAdjQtrFee.Text) != "")
    //        {
    //            objBillingInvoice.ssi_adjustedfee = new CrmMoney();
    //            objBillingInvoice.ssi_adjustedfee.Value = Convert.ToDecimal(txtAdjQtrFee.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }


    //        //ssi_adjustment
    //        if (Convert.ToString(txtAdjAmt.Text) != "")
    //        {
    //            objBillingInvoice.ssi_adjustment = new CrmMoney();
    //            objBillingInvoice.ssi_adjustment.Value = Convert.ToDecimal(txtAdjAmt.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));
    //        }

    //        //ssi_adjustmentreason
    //        if (Convert.ToString(txtAdjReason.Text) != "")
    //        {
    //            objBillingInvoice.ssi_adjustmentreason = Convert.ToString(txtAdjReason.Text);
    //        }



    //        //ClientType
    //        if (Convert.ToString(ddlClientType.SelectedValue) != "")
    //        {
    //            objBillingInvoice.ssi_feescheduletype = new Picklist();
    //            //objBillingInvoice.ssi_feescheduleid.type = EntityName.ssi_feeschedule.ToString();
    //            objBillingInvoice.ssi_feescheduletype.Value = Convert.ToInt32(ddlClientType.SelectedValue);
    //        }

    //        //RelationshipFee
    //        if (Convert.ToString(txtRelationshipFee.Text) != "")
    //        {
    //            objBillingInvoice.ssi_relationshipfee = new CrmMoney();
    //            objBillingInvoice.ssi_relationshipfee.Value = Convert.ToDecimal(txtRelationshipFee.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_relationshipfee = new CrmMoney();
    //            objBillingInvoice.ssi_relationshipfee.IsNull = true;
    //            objBillingInvoice.ssi_relationshipfee.IsNullSpecified = true;
    //        }

    //        //Custom FeeAmount
    //        if (Convert.ToString(txtCustFeeAmount.Text) != "")
    //        {
    //            objBillingInvoice.ssi_customfee = new CrmMoney();
    //            objBillingInvoice.ssi_customfee.Value = Convert.ToDecimal(txtCustFeeAmount.Text.Replace("$", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", ""));

    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_customfee = new CrmMoney();
    //            objBillingInvoice.ssi_customfee.IsNull = true;
    //            objBillingInvoice.ssi_customfee.IsNullSpecified = true;
    //        }

    //        //discount
    //        if (Convert.ToString(txtDiscount.Text) != "")
    //        {
    //            objBillingInvoice.ssi_discount = new CrmDecimal();
    //            objBillingInvoice.ssi_discount.Value = Convert.ToDecimal(txtDiscount.Text.Replace("%", "").Replace(",", "").Replace(",", "").Replace("(", "-").Replace(")", "")) / 100;
    //        }
    //        else
    //        {
    //            objBillingInvoice.ssi_discount = new CrmDecimal();
    //            objBillingInvoice.ssi_discount.IsNull = true;
    //            objBillingInvoice.ssi_discount.IsNullSpecified = true;
    //        }

    //        if (chkAccured.Checked)
    //        {

    //        }

    //        service.Create(objBillingInvoice);

    //        /*Invoice Created*/
    //        lblMessage.Text = "Records Created Successfully";
    //        System.Web.UI.WebControls.ListItem itemToRemove = ddlClientType.Items.FindByValue("0");
    //        if (itemToRemove != null)
    //        {
    //            ddlClientType.Items.Remove(itemToRemove);
    //        }
    //    }

    //    catch (Exception ex)
    //    {
    //        //Failed to Create Invoice
    //        lblMessage.Text = "Billing Invoice Failed to Create : " + ex.ToString();
    //    }

    //}
    #endregion

    protected void txtAUMDate_TextChanged(object sender, EventArgs e)
    {
        bool bProceed = false;

        if (txtAUMDate.Text != "")

        {
            string inputString = txtAUMDate.Text;
            DateTime dDate;

            if (DateTime.TryParse(inputString, out dDate))
            {
                String.Format("{0:MM/dd/yyyy}", dDate);
                bProceed = true; ;
                txtAUMDate.Text = dDate.ToString("MM/dd/yyyy");
            }

        }
        if (bProceed)
        {
            lblMessage.Text = "";
            ClearControls();
            //if (txtAUMDate.Text.Trim() != "")
            //{
            //    /*split up date into different variables strMont,strday,strYear*/
            //    string[] datesplit = txtAUMDate.Text.Split('/');
            //    string strMonth = "";
            //    string strday = "";
            //    string strYear = "";

            //    int month = Convert.ToInt32(datesplit[0]);
            //    int day = Convert.ToInt32(datesplit[1]);
            //    int year = Convert.ToInt32(datesplit[2]);

            //    //check the length of month entered and append if necessary
            //    if (month.ToString().Length < 2)
            //        strMonth = "0" + month;
            //    else
            //        strMonth = month.ToString();

            //    //check the length of day entered and append if necessary
            //    if (day.ToString().Length < 2)
            //        strday = "0" + day;
            //    else
            //        strday = day.ToString();

            //    //check the length of year entered and append if necessary
            //    if (year.ToString().Length == 2)
            //        strYear = "20" + year;
            //    else
            //        strYear = year.ToString();
            //    //Reformatt date using "/"
            //    txtAUMDate.Text = strMonth + "/" + strday + "/" + strYear;
            //}

            CheckandGetExistingData();

            txtTotalAUM.Focus();
        }
        else
        {
            txtAUMDate.Text = "";
            lblMessage.Text = "Please Enter Proper Date";
        }


    }
    protected void rdolstClientType_SelectedIndexChanged(object sender, EventArgs e)
    {
        // ClearControls();

        string value = txtBillingAUM.Text.Replace(",", "").Replace("$", "");
        if (txtBillingAUM.Text != "")
        {
            DB clsDB = new DB();
            object BillingAUM = txtBillingAUM.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "'";


            /* Return BillingAnnualFeeCAlc  into dataset where BillingAum,ClientType is Selected */
            DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @Billin gAumAmount=" + BillingAUM + ",@ClientType='" + ddlClientType.SelectedItem.Text + "',@TotalAUMNmb='" + txtTotalAUM.Text.Replace(",", "").Replace("$", "") + "'");

            if (dsBilling.Tables[0].Rows.Count > 0)
            {
                //check whether AnnualFee is null , if null set "0.0" or else set value
                string AnnualFee = Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]);
                txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));
                ViewState["BillingId"] = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingId"]);
            }

            CalulateFeeRate();
            decimal ul;

            if (decimal.TryParse(value, out ul))
            {

                txtBillingAUM.TextChanged -= txtBillingAUM_TextChanged;
                txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                txtBillingAUM.TextChanged += txtBillingAUM_TextChanged;
            }
        }
    }

    private void ClearControls()
    {
        System.Web.UI.WebControls.ListItem itemToRemove = ddlClientType.Items.FindByValue("0");
        if (itemToRemove != null)
        {
            ddlClientType.Items.Remove(itemToRemove);
        }

        //Clear all the following controls

        txtTotalAUM.Text = "";
        txtBillingAUM.Text = "";
        txtCustFeeAmount.Text = "";
        txtFeeRateCalc.Text = "";
        txtFeesPerMonth.Text = "";
        txtQuaterlyFeeCalc.Text = "";
        txtStdAnnualFeeCalc.Text = "";
        txtRelationshipFee.Text = "";
        txtAdjAmt.Text = "";
        txtAdjQtrFee.Text = "";
        txtAdjReason.Text = "";
        txtDiscount.Text = "";
        //txtBox1.Text = "";
        //txtBox2.Text = "";
        //txtBox3.Text = "";
        lbGroup.Visible = false;
        LinkButton1.Visible = false;
        txtTotalAnnualFee.Text = "";
        btnDeleteAndCal.Visible = false;
        txtSecurityFee.Visible = false;
        FeeAmount = "0";
        Label5.Visible = false;
        txtNotes.Text = "";
        txtSecurityFee.Text = "";
        txtTotalAUM.Enabled = false;
        txtBillingAUM.Enabled = false;
        txtTotalBillingAUM.Text = "";
        txtTotalBillingAUM.Enabled = false;

        //txtBox1.Style.Add("display", "none");
        //lblbox1.Style.Add("display", "none");
        //txtBox2.Style.Add("display", "none");
        //lblbox2.Style.Add("display", "none");
        //txtBox3.Style.Add("display", "none");
        //lblbox3.Style.Add("display", "none");
        LinkButton1.Visible = true;
        gvFlatFee.DataSource = null;
        gvFlatFee.DataBind();
        gvFlatFee.Visible = false;
        ViewState["TotalFlatFee"] = null;
        #region Standard Min-Max Changes- 3_25_2019

        ddlClientType.SelectedValue = "100000000";

        lblbpsfees.Visible = false;
        lblMinValu.Visible = false;
        lblMaxVal.Visible = false;

        txtbpsfee.Visible = false;
        txtMinVal.Visible = false;
        txtMaxVal.Visible = false;

        txtbpsfee.Text = "";
        txtMinVal.Text = "";
        txtMaxVal.Text = "";
        btnStandardCalculate.Visible = false;

        ViewState["MaxValue"] = "";
        ViewState["MinValue"] = "";
        ViewState["bpsfeeValue"] = "";



        #endregion
        //added 7_15_2019
        txtFeeAUM.Text = "";
        txtFeeAUM.Visible = false;
        txtBillAUM.Visible = false;
        txtBillAUM.Text = "";
    }

    protected void Timer1_Tick(object sender, EventArgs e)
    {
        string confirmValue = Request.Form["confirm_value"];
        if (confirmValue == "Yes")
        {
            //  lblMessage.Text = "YES";
            Timer1.Enabled = false;
            DataSet dsBillingChk = (DataSet)ViewState["dsBillingCheck"];
            for (int i = 0; i < dsBillingChk.Tables[0].Rows.Count; i++)
            {
                //Updating Invoice if Yes selected by user for the same InvoiceID
                // UpdateInvoice(dsBillingChk.Tables[0].Rows[i]["Ssi_billinginvoiceId"].ToString(), BillingUUID);
            }
        }
        else if (confirmValue == "No")
        {
            //  lblMessage.Text = "NO";
            Timer1.Enabled = false;
        }
        else
            lblMessage.Text = "";

    }
    protected void ddlBillFor_SelectedIndexChanged1(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        checkbillfortype();
        ClearControls();

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

        /*check for value or null and assign to Advisor Value*/
        /*check for value or null and assign to BillForValue */
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

        CheckandGetExistingData();
        txtAUMDate.Focus();

    }

    protected void ddlClientType_SelectedIndexChanged(object sender, EventArgs e)
    {
        FeeScheduleddlChanges();

    }
    protected void txtRelationshipFee_TextChanged(object sender, EventArgs e)
    {

        string value = txtRelationshipFee.Text.Replace(",", "").Replace("$", "");
        decimal ul;

        if (decimal.TryParse(value, out ul))
        {
            txtRelationshipFee.TextChanged -= txtRelationshipFee_TextChanged;
            txtRelationshipFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
            txtRelationshipFee.TextChanged += txtRelationshipFee_TextChanged;
        }
        //CalculateFeeRate Function called
        CalulateFeeRate();
        btnSubmit.Focus();
    }
    protected void txtAdjAmt_TextChanged(object sender, EventArgs e)
    {
        string value = txtAdjAmt.Text.Replace(",", "").Replace("$", "");
        decimal ul;

        if (decimal.TryParse(value, out ul))
        {
            txtAdjAmt.TextChanged -= txtAdjAmt_TextChanged;
            txtAdjAmt.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
            txtAdjAmt.TextChanged += txtAdjAmt_TextChanged;
        }
        //string strQuaterlyFeeCalc = Convert.ToString(txtQuaterlyFeeCalc.Text.Replace(",", "").Replace("$", ""));
        //string strAdjAmt = Convert.ToString(txtAdjAmt.Text.Replace(",", "").Replace("$", ""));

        //double dblQuaterlyFeeCalc = 0.0;
        //double dblAdjAmt = 0.0;
        //double dblAdjQtrFee = 0.0;


        //if (!string.IsNullOrEmpty(strQuaterlyFeeCalc))
        //    dblQuaterlyFeeCalc = Convert.ToDouble(strQuaterlyFeeCalc);

        //if (!string.IsNullOrEmpty(strAdjAmt))
        //    dblAdjAmt = Convert.ToDouble(strAdjAmt);

        //dblAdjQtrFee = dblQuaterlyFeeCalc - dblAdjAmt;

        CalulateFeeRate();

        //   txtAdjQtrFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", dblAdjQtrFee);

        txtAdjReason.Focus();


    }
    protected void txtAdjQtrFee_TextChanged(object sender, EventArgs e)
    {

    }
    protected void Hcheckadjvalue_ValueChanged(object sender, EventArgs e)
    {

    }
    protected void txtDiscount_TextChanged(object sender, EventArgs e)
    {
        string value = txtDiscount.Text.Replace(",", "").Replace("$", "").Replace("%", "");
        if (value == "")
            txtCustFeeAmount.Text = "";

        int ind = 0;


        decimal ul;

        if (decimal.TryParse(value, out ul))
        {
            txtDiscount.TextChanged -= txtDiscount_TextChanged;
            txtDiscount.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:P2}", ul);
            txtDiscount.TextChanged += txtDiscount_TextChanged;
        }


        CalulateFeeRate();

    }

    public void NewData()
    {
        // added By abhi
        DataSet dsBilling;
        DB clsDB = new DB();
        if (ddlHH.SelectedValue != "" && txtAUMDate.Text != "" && ddlBillFor.SelectedValue != "")
        {

            object HHValue = ddlHH.SelectedValue == "00000000-0000-0000-0000-000000000000" ? "null" : "'" + ddlHH.SelectedValue + "'";


            string strBillFor = ddlBillFor.SelectedValue;

            int len = strBillFor.IndexOf("|");
            strBillFor = strBillFor.Substring(0, len);

            //     dsBilling = clsDB.getDataSet("EXEC SP_R_BILLING_OVERALLFEE @BillingUUID = '801DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331',@ClientType = '.75% Below Minimum Client (lesser of $180K or .75% AUM)'");
            //    dsBilling = clsDB.getDataSet("SP_R_BILLING_OVERALLFEE @BillingUUID ='7E1DE384-D1A4-E511-9418-005056A0567E', @AsOfDate = '20160331',@ClientType = '.75% Below Minimum Client (lesser of $180K or .75% AUM)'");
            dsBilling = clsDB.getDataSet("EXEC SP_R_BILLING_OVERALLFEE @BillingUUID = '" + strBillFor + "',@AsOfDate = '" + txtAUMDate.Text + "',@ClientType ='" + ddlClientType.SelectedItem.Text + "'");
            ViewState["dsData"] = dsBilling;

            if (dsBilling.Tables[0].Rows.Count > 0)
            {
                //Check AnnualFee is null or not ,If null set 0.0 else assign value
                //string AnnualFee = Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]);
                //txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));
                //ViewState["BillingId"] = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingId"]);


                //Check AnnualFee is null or not ,If null set 0.0 else assign value
                string AnnualFee = Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]);

                txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));
                //    ViewState["BillingId"] = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingId"]);

                string TotalAUM = Convert.ToString(dsBilling.Tables[0].Rows[0]["TotalAUM"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["TotalAUM"]);
                if (Convert.ToDecimal(TotalAUM) == 0)
                {
                    txtTotalAUM.Enabled = true;
                    LinkButton1.Visible = false;
                }
                else
                {
                    txtTotalAUM.Enabled = false;
                    txtTotalAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(TotalAUM));
                }


                string BillAUM = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingAUM"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingAUM"]);
                if (Convert.ToDecimal(BillAUM) == 0)
                {
                    txtBillingAUM.Enabled = true;
                    LinkButton1.Visible = false;
                    txtBillingAUM.Text = "0";
                    txtBillAUM.Visible = false; //added 7_15_2019
                    lblStandardFeeAssets.Visible = false;//added 7_15_2019
                }
                else
                {
                    txtBillingAUM.Enabled = false;
                    txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(BillAUM));
                    txtBillAUM.Enabled = false; //added 7_15_2019
                    txtBillAUM.Visible = true; //added 7_15_2019
                    lblStandardFeeAssets.Visible = true;//added 7_15_2019
                    txtBillAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(BillAUM)); //added 7_15_2019
                }
                txtBillingAUM.Visible = false;


                FeeAmount = Convert.ToString(dsBilling.Tables[0].Rows[0]["FeePctValue"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["FeePctValue"]);
                if (Convert.ToDecimal(FeeAmount) != 0)
                {
                    txtSecurityFee.Visible = true;
                    Label5.Visible = true;
                    txtSecurityFee.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(FeeAmount));
                }
                else
                {
                    txtSecurityFee.Visible = false;
                    Label5.Visible = false;
                    FeeAmount = "0";
                }
                //added 7_15_2019
                string FeeAmount1 = Convert.ToString(dsBilling.Tables[0].Rows[0]["FeeAUM"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["FeeAUM"]);
                if (Convert.ToDecimal(FeeAmount1) != 0)
                {
                    txtFeeAUM.Visible = true;
                    lblAssetsUnderAdministration.Visible = true;
                    txtFeeAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(FeeAmount1));
                }
                else
                {
                    lblAssetsUnderAdministration.Visible = false;
                    txtFeeAUM.Visible = false;
                }

                string TotalBillingAUM = Convert.ToString(dsBilling.Tables[0].Rows[0]["TotalBillingAUM"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["TotalBillingAUM"]);
                if (Convert.ToDecimal(TotalBillingAUM) != 0)
                {
                    txtTotalBillingAUM.Enabled = false;
                    //  Label5.Visible = true;
                    txtTotalBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(TotalBillingAUM));
                }
                else
                {
                    txtTotalBillingAUM.Enabled = true;
                    //txtTotalBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(TotalBillingAUM));
                    //Label5.Visible = false;
                    // FeeAmount = "0";
                }



                // ddlClientType.SelectedValue = Convert.ToString(dsBilling.Tables[0].Rows[0]["FeeScheduleID"]);
                string FeeScheduleID = (Convert.ToString(dsBilling.Tables[0].Rows[0]["FeeScheduleID"]));
                if (FeeScheduleID != "")
                {
                    ddlClientType.SelectedValue = FeeScheduleID;
                }

                #region Standard Min-Max Changes - 3_25_2019

                string ssi_minimumfeein = Convert.ToString(dsBilling.Tables[0].Rows[0]["minimumfeein"]);
                string ssi_maximumfeeasa = Convert.ToString(dsBilling.Tables[0].Rows[0]["maximumfeeasa"]);
                string ssi_feeonfirst25mminbps = Convert.ToString(dsBilling.Tables[0].Rows[0]["feeonfirst25mminbps"]);


                // Maximum fee as a %- added 3_25_2019 Standard Min-Max Changes
                if (ssi_maximumfeeasa != "")
                {
                    txtMaxVal.Text = ssi_maximumfeeasa.ToString();
                    if (txtMaxVal.Text != "")
                    {
                        double doubleValue;
                        if (double.TryParse(txtMaxVal.Text, out doubleValue))
                        {
                            /*Convert doubleValue to percentage*/
                            doubleValue = doubleValue / 100;
                            txtMaxVal.Text = doubleValue.ToString("0.00%");
                        }
                    }
                }
                //Minimum fee in $ - added 3_25_2019 Standard Min-Max Changes
                if (ssi_minimumfeein != "")
                {
                    txtMinVal.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDouble(ssi_minimumfeein));
                }
                // Fee on first $25MM in bps- added 3_25_2019 Standard Min-Max Changes
                if (ssi_feeonfirst25mminbps != "")
                {
                    txtbpsfee.Text = ssi_feeonfirst25mminbps.ToString();
                    if (txtbpsfee.Text != "")
                    {
                        double doubleValue;
                        if (double.TryParse(txtbpsfee.Text, out doubleValue))
                        {
                            txtbpsfee.Text = doubleValue.ToString("0.00");
                        }
                    }
                }
            }

            string value = txtBillingAUM.Text.Replace(",", "").Replace("$", "");
            if (txtBillingAUM.Text.Trim() != "")
            {

                ViewState["bpsfeeValue"] = txtbpsfee.Text;
                ViewState["MinValue"] = txtMinVal.Text;
                ViewState["MaxValue"] = txtMaxVal.Text;
                //object BillingAUM = txtBillingAUM.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "'";           
                object txtbpsfeeValue = txtbpsfee.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtbpsfee.Text.Replace(",", "").Replace("$", "") + "'";
                string txtMinValue = txtMinVal.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtMinVal.Text.Replace(",", "").Replace("$", "") + "'";
                object txtMaxValue = txtMaxVal.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtMaxVal.Text.Replace(",", "").Replace("%", "") + "'";

                //class Library
                /* Return AnnualFee  into dataset where Billing Amount and clientType Selected */
                //  DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "' ");
                //DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + "");
                //double relatioshipfee = 0.0;
                //double trustfee = 0.0;
                //double administrativefee = 0.0;
                //double servicefee = 0.0;
                //double setupfee = 0.0;
                //double transactionfee = 0.0;
                double securityfee = 0.0;
                //string ssi_comments = string.Empty;
                //string FeeValue = string.Empty;
                //foreach (GridViewRow row in gvFlatFee.Rows)
                //{

                //    System.Web.UI.WebControls.Label relatioshipfee1 = (System.Web.UI.WebControls.Label)row.FindControl("lblFlatFee");
                //    System.Web.UI.WebControls.TextBox txtfeevalue = (System.Web.UI.WebControls.TextBox)row.FindControl("txtFlatFee");



                //    ssi_comments = relatioshipfee1.Text;
                //    FeeValue = txtfeevalue.Text;


                //    if (ssi_comments.ToLower().Trim() == "relationship fee")
                //        relatioshipfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
                //    if (ssi_comments.ToLower().Trim() == "trust fee")
                //        trustfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
                //    if (ssi_comments.ToLower().Trim() == "administrative service fee")
                //        administrativefee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
                //    if (ssi_comments.ToLower().Trim() == "service fee")
                //        servicefee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
                //    if (ssi_comments.ToLower().Trim() == "setup fee")
                //        setupfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
                //    if (ssi_comments.ToLower().Trim() == "transaction fee")
                //        transactionfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));

                //}



                if (txtSecurityFee.Text != "")
                    securityfee = Convert.ToDouble(txtSecurityFee.Text.Replace(",", "").Replace("$", ""));












                // DataSet dsBilling1 = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + ",@reCalculateFlg =1");
                DataSet dsBilling1 = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @TotalAUMNmb='" + txtTotalAUM.Text.Replace(",", "").Replace("$", "") + "',@BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + ",@reCalculateFlg =1,@SecurityFeeNum = " + securityfee);



                if (dsBilling1.Tables[0].Rows.Count > 0)
                {
                    /* check AnnualFee if Null set AnnualFee="0.0" else display AnnualFee */
                    string AnnualFee = Convert.ToString(dsBilling1.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling1.Tables[0].Rows[0]["AnnualFee"]);

                    /* used Globalization To Format StdAnnualFeeCalc in (en-US) with 2 decimal */
                    txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));

                    ViewState["BillingId"] = Convert.ToString(dsBilling1.Tables[0].Rows[0]["BillingId"]);
                }

                CalulateFeeRate();//function to calculateFeeRate
                decimal ul;

                /*->out keyword used to display billingAUM
                ->TryParse used to convert into Decimal*/
                if (decimal.TryParse(value, out ul))
                {
                    txtBillingAUM.TextChanged -= txtBillingAUM_TextChanged;
                    txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                    txtBillingAUM.TextChanged += txtBillingAUM_TextChanged;
                }
            }
            else
            {

                /* clear Following Control*/
                txtCustFeeAmount.Text = "";
                txtFeeRateCalc.Text = "";
                txtFeesPerMonth.Text = "";
                txtQuaterlyFeeCalc.Text = "";
                txtStdAnnualFeeCalc.Text = "";
            }

            #endregion
            //-------------- 2nd dataset
            if (dsBilling.Tables[1] != null)
            {
                #region not used
                //  txtBox1.Style.Add("display", "none");
                //  lblbox1.Style.Add("display", "none");
                //  txtBox2.Style.Add("display", "none");
                //  txtBox2.Style.Add("display", "none");
                //  txtBox3.Style.Add("display", "none");
                //  lblbox3.Style.Add("display", "none");
                //  //RequiredFieldValidator1.Visible = false;
                //  //RequiredFieldValidator13.Visible = false;
                //  //RequiredFieldValidator14.Visible = false;
                //  DataTable dtdata = dsBilling.Tables[1];

                //  int cnt = dtdata.Rows.Count;
                //  for (int i = 0; i < dtdata.Rows.Count; i++)
                //  {
                //      string comment = dtdata.Rows[i]["ssi_Comments"].ToString();
                //      string value = dtdata.Rows[i]["Fees"].ToString();
                //      string allocationflag = dtdata.Rows[i]["AllocatedFlg"].ToString();

                //      if (i == 0)
                //      {
                //          txtBox1.Style.Add("display", "");
                //          lblbox1.Style.Add("display", "");
                //          Tr1.Style.Add("display", "");

                //          txtBox1.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(value));
                //          lblbox1.Text = comment;

                //          //RequiredFieldValidator1.Visible = true;
                //          if (allocationflag == "0")
                //          {
                //              txtBox1.Enabled = false;
                //          }
                //          else
                //          {
                //              txtBox1.Enabled = true;
                //          }

                //          //txtBox2.Style.Add("display", "none");
                //          //txtBox2.Style.Add("display", "none");
                //          //txtBox3.Style.Add("display", "none");
                //          //lblbox3.Style.Add("display", "none");
                //          //trRelationshipFee.Style.Add("display", "none");
                //      }
                //      else if (i == 1)
                //      {
                //          txtBox2.Style.Add("display", "");
                //          lblbox2.Style.Add("display", "");
                //          //  txtBox2.Text = value;

                //          txtBox2.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(value));
                //          lblbox2.Text = comment;
                //          Tr2.Style.Add("display", "");
                //          //RequiredFieldValidator13.Visible = true;

                //          if (allocationflag == "0")
                //          {
                //              txtBox2.Enabled = false;
                //          }
                //          else
                //          {
                //              txtBox2.Enabled = true;
                //          }
                //          //txtBox3.Style.Add("display", "none");
                //          //lblbox3.Style.Add("display", "none");
                //      }
                //      if (i == 2)
                //      {
                //          txtBox3.Style.Add("display", "");
                //          lblbox3.Style.Add("display", "");
                //          txtBox3.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(value));
                //          lblbox3.Text = comment;
                //          Tr3.Style.Add("display", "");
                //          //RequiredFieldValidator14.Visible = true;
                //          if (allocationflag == "0")
                //          {
                //              txtBox3.Enabled = false;
                //          }
                //          else
                //          {
                //              txtBox3.Enabled = true;
                //          }
                //      }

                //  }
                ////  RequiredFieldValidator7.Visible = false;
                //  decimal totalfees = Convert.ToDecimal(txtStdAnnualFeeCalc.Text.Replace("$", "").Replace(",", ""));
                //  if (cnt == 1)
                //  {
                //      totalfees = totalfees + Convert.ToDecimal(txtBox1.Text.Replace("$", "").Replace(",", ""));
                //  }
                //  if (cnt == 2)
                //  {
                //      totalfees = totalfees + Convert.ToDecimal(txtBox1.Text.Replace("$", "").Replace(",", "")) + Convert.ToDecimal(txtBox2.Text.Replace("$", "").Replace(",", ""));
                //  }
                //  if (cnt == 3)
                //  {
                //      totalfees = totalfees + Convert.ToDecimal(txtBox1.Text.Replace("$", "").Replace(",", "")) + Convert.ToDecimal(txtBox2.Text.Replace("$", "").Replace(",", "")) + Convert.ToDecimal(txtBox3.Text.Replace("$", "").Replace(",", ""));
                //  }
                #endregion

                if (dsBilling.Tables[1].Rows.Count > 0)
                {
                    gridviewBind(dsBilling.Tables[1]);
                    string TotalFlatFee = dsBilling.Tables[1].Rows[0]["totFlatFee"].ToString();
                    ViewState["TotalFlatFee"] = TotalFlatFee;
                }
                else
                    ViewState["TotalFlatFee"] = null;



                //decimal feeRate =  (totalfees/Convert.ToDecimal(txtBillingAUM.Text.Replace("$", "").Replace(",", "")) ) * 100;
                //decimal quater = totalfees / 4;
                //decimal month = totalfees / 12;

                //txtFeeRateCalc.Text = feeRate.ToString();
                //txtQuaterlyFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(quater));
                //txtFeesPerMonth.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(month)); 


            }
            FeeScheduleddlChanges();

            CalulateFeeRate();


            if (dsBilling.Tables[2] != null)
            {
                if (dsBilling.Tables[2].Rows.Count > 1)
                {
                    lbGroup.Visible = true;
                    BindListBox(dsBilling.Tables[2]);

                    //for (int i = 0; i < dsBilling.Tables[2].Rows.Count; i++)
                    //{
                    //    SourceFileArray[i] = dsBilling.Tables[2].Rows[i]["BillingName"].ToString();
                    //}
                    //ViewState["SourceFileArray"] = SourceFileArray;
                }
                else
                {
                    lbGroup.Visible = false;
                }
                for (int i = 0; i < dsBilling.Tables[2].Rows.Count; i++)
                {
                    string data = dsBilling.Tables[2].Rows[i]["ssi_billingId"].ToString();
                    // SourceFileArray[i] = data;
                    ls.Add(data);
                }

            }
            else
            {
                lbGroup.Visible = false;
            }

            if (dsBilling.Tables[3] != null)
            {
                ViewState["CustomeFeeInsert"] = dsBilling.Tables[3];
            }
            ViewState["sourceLsit"] = ls;
            ViewState["Billing"] = dsBilling.Tables[2];
            ViewState["BillingForList"] = dsBilling.Tables[2];


            if (dsBilling.Tables[4] != null)
            {
                txtNotes.Text = dsBilling.Tables[4].Rows[0]["AdvisoryNotes"].ToString();
                ViewState["FamilyNotes"] = dsBilling.Tables[4];
            }

        }
    }

    public void FeeScheduleddlChanges()
    {

        // ClearControls();
        txtCustFeeAmount.Text = "";
        txtAdjAmt.Text = "";
        lblMessage.Text = "";
        checkbillfortype();
        DataSet dsBilling;
        // ClearControls();
        // txtAUMDate.Focus();
        // lbGroup.Visible = false;
        //Replace , with $ to disply amount in $ format
        string value = txtBillingAUM.Text.Replace(",", "").Replace("$", "");
        if (txtBillingAUM.Text != "")
        {
            DB clsDB = new DB();//class libarary
            /*check for value or null and assign to BillingAum*/
            object BillingAUM = txtBillingAUM.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "'";

            object txtbpsfeeValue = txtbpsfee.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtbpsfee.Text.Replace(",", "").Replace("$", "") + "'";
            string txtMinValue = txtMinVal.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtMinVal.Text.Replace(",", "").Replace("$", "") + "'";
            object txtMaxValue = txtMaxVal.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtMaxVal.Text.Replace(",", "").Replace("%", "") + "'";
            object textboxvaltxtTotalAUM = txtTotalAUM.Text.Replace(",", "").Replace("$", "") == "" ? "0" : "'" + txtTotalAUM.Text.Replace(",", "").Replace("$", "") + "'"; //changed 29_8_2019-sasmit
            /* Return BillingAnnualFeeCAlc  into dataset where BillingAum,ClientType is Selected */
            //double relatioshipfee = 0.0;
            //double trustfee = 0.0;
            //double administrativefee = 0.0;
            //double servicefee = 0.0;
            //double setupfee = 0.0;
            //double transactionfee = 0.0;
            double securityfee = 0.0;
            //string ssi_comments = string.Empty;
            //string FeeValue = string.Empty;
            //foreach (GridViewRow row in gvFlatFee.Rows)
            //{

            //    System.Web.UI.WebControls.Label relatioshipfee1 = (System.Web.UI.WebControls.Label)row.FindControl("lblFlatFee");
            //    System.Web.UI.WebControls.TextBox txtfeevalue = (System.Web.UI.WebControls.TextBox)row.FindControl("txtFlatFee");



            //    ssi_comments = relatioshipfee1.Text;
            //    FeeValue = txtfeevalue.Text;


            //    if (ssi_comments.ToLower().Trim() == "relationship fee")
            //        relatioshipfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "trust fee")
            //        trustfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "administrative service fee")
            //        administrativefee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "service fee")
            //        servicefee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "setup fee")
            //        setupfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "transaction fee")
            //        transactionfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));

            //}



            if (txtSecurityFee.Text != "")
                securityfee = Convert.ToDouble(txtSecurityFee.Text.Replace(",", "").Replace("$", ""));





            // dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount=" + BillingAUM + ",@ClientType='" + ddlClientType.SelectedItem.Text + "' ,@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + ",@reCalculateFlg =1");
            //dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @TotalAUMNmb='" + txtTotalAUM.Text.Replace(",", "").Replace("$", "") + "',@BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + ",@reCalculateFlg =1,@SecurityFeeNum = " + securityfee);
            //changed 29_8_2019 -sasmit
            dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @TotalAUMNmb=" + textboxvaltxtTotalAUM + ",@BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + ",@reCalculateFlg =1,@SecurityFeeNum = " + securityfee);


            if (dsBilling.Tables[0].Rows.Count > 0)
            {
                //Check AnnualFee is null or not ,If null set 0.0 else assign value
                string AnnualFee = Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]);
                txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));
                ViewState["BillingId"] = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingId"]);
            }


            //calculateFeeRate();
            CalulateFeeRate();
            decimal ul;

            if (decimal.TryParse(value, out ul))
            {
                txtBillingAUM.TextChanged -= txtBillingAUM_TextChanged;
                txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                txtBillingAUM.TextChanged += txtBillingAUM_TextChanged;
            }
        }

        /*Set RelationshipFee to Null
        hide RelationshipFee TextBox -->RelationshipFee.Style.Add("display", "none");
        RequiredFieldValidator11 
        remove Readonly Property from CustomFee
          */
        txtRelationshipFee.Text = "";
        trRelationshipFee.Style.Add("display", "none");
        //  RequiredFieldValidator11.Visible = false;
        txtCustFeeAmount.ReadOnly = false;
        txtCustFeeAmount.BackColor = System.Drawing.Color.White;

        string strFeeType = "";
        if (Convert.ToString(ddlBillFor.SelectedValue) != "ALL" && Convert.ToString(ddlBillFor.SelectedValue) != "")
        {
            //set selected BilFor to ---> FeeType,
            strFeeType = ddlBillFor.SelectedValue;
            char[] delimiterChars = { '|' };
            string[] words = strFeeType.ToString().Split(delimiterChars);
            int len = words.Length;
            if (len > 0)
                strFeeType = words[1];
        }


        if (ddlClientType.SelectedValue == "100000007") //Custom
        {
            /* for FeeType=Custom 
               Hide AnnualFeeCalc And
               RequiredValidator7 
               Display RequiredValidator8 
             * and discount
             * 
             */
            txtStdAnnualFeeCalc.Text = "";
            //trStdAnnualFeeCalc.Style.Add("display", "none");
            lblStdAnnFeeCalc.Visible = false;
            txtStdAnnualFeeCalc.Visible = false;
            // txtCustFeeAmount.Focus();
            RequiredFieldValidator7.Visible = false;
            RequiredFieldValidator8.Visible = true;
            txtDiscount.Visible = false;
            lbldiscount.Visible = false;
            lbldecimal.Visible = false;


            #region Standard Min-Max Changes- 3_25_2019
            lblbpsfees.Visible = false;
            lblMinValu.Visible = false;
            lblMaxVal.Visible = false;

            txtbpsfee.Visible = false;
            txtMinVal.Visible = false;
            txtMaxVal.Visible = false;

            txtbpsfee.Text = "";
            txtMinVal.Text = "";
            txtMaxVal.Text = "";
            btnStandardCalculate.Visible = false;
            #endregion
        }
        else if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
        {
            /* for FeeType=Standard with Relationship Fees
               Display AnnualFeeCalc And
               RequiredValidator7 
               RequiredValidator11
               RelationshipFee
               CustomFeeAmount
             * Dilspy discount
             */
            // trStdAnnualFeeCalc.Style.Add("display", "");


            #region Standard Min-Max Changes- 3_25_2019
            lblbpsfees.Visible = false;
            lblMinValu.Visible = false;
            lblMaxVal.Visible = false;

            txtbpsfee.Visible = false;
            txtMinVal.Visible = false;
            txtMaxVal.Visible = false;

            txtbpsfee.Text = "";
            txtMinVal.Text = "";
            txtMaxVal.Text = "";
            btnStandardCalculate.Visible = false;
            #endregion
            lblStdAnnFeeCalc.Visible = true;
            txtStdAnnualFeeCalc.Visible = true;
            RequiredFieldValidator7.Visible = true;
            //  RequiredFieldValidator11.Visible = true;
            trRelationshipFee.Style.Add("display", "none");
            txtCustFeeAmount.ReadOnly = true;
            txtCustFeeAmount.BackColor = System.Drawing.Color.Gray;
            txtDiscount.Visible = true;
            lbldiscount.Visible = true;
            lbldecimal.Visible = true;
        }
        else
        {
            /* for FeeType= other than Custom and Standard with Relationship Fees
               Hide RelationshipFee
               Display AnnualFeeCalc
               Dispaly RequiredValidator7 
               Hide RequiredValidator11
             */

            //added 3_25_2019 
            if (ddlClientType.SelectedValue == "100000009")//Standard Fee or $180K (greater of the two) //if (ddlClientType.SelectedValue == "100000000")
            {
                lblbpsfees.Visible = true;
                lblMinValu.Visible = true;
                lblMaxVal.Visible = true;

                txtbpsfee.Visible = true;
                txtMinVal.Visible = true;
                txtMaxVal.Visible = true;
                btnStandardCalculate.Visible = true;
            }
            else
            {
                lblbpsfees.Visible = false;
                lblMinValu.Visible = false;
                lblMaxVal.Visible = false;

                txtbpsfee.Visible = false;
                txtMinVal.Visible = false;
                txtMaxVal.Visible = false;

                txtbpsfee.Text = "";
                txtMinVal.Text = "";
                txtMaxVal.Text = "";
                btnStandardCalculate.Visible = false;
            }
            trRelationshipFee.Style.Add("display", "none");
            // trStdAnnualFeeCalc.Style.Add("display", "");
            lblStdAnnFeeCalc.Visible = true;
            txtStdAnnualFeeCalc.Visible = true;
            RequiredFieldValidator7.Visible = true;
            RequiredFieldValidator8.Visible = false;
            txtDiscount.Visible = true;
            lbldiscount.Visible = true;
            lbldecimal.Visible = true;

            if (strFeeType.ToUpper() == "Flat".ToUpper())
            {
                //trStdAnnualFeeCalc.Style.Add("display", "none");
                lblStdAnnFeeCalc.Visible = false;
                txtStdAnnualFeeCalc.Visible = false;
                RequiredFieldValidator8.Visible = true;
                RequiredFieldValidator7.Visible = false;
            }

        }


    }

    //protected void txtBox1_TextChanged(object sender, EventArgs e)
    //{

    //    CalulateFeeRate();
    //    decimal val = Convert.ToDecimal(txtBox1.Text.Replace("$", "").Replace(",", ""));
    //    txtBox1.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", val);
    //}
    //protected void txtBox2_TextChanged(object sender, EventArgs e)
    //{
    //    CalulateFeeRate();
    //    decimal val = Convert.ToDecimal(txtBox2.Text.Replace("$", "").Replace(",", ""));
    //    txtBox2.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", val);
    //}
    //protected void txtBox3_TextChanged(object sender, EventArgs e)
    //{
    //    CalulateFeeRate();
    //    decimal val = Convert.ToDecimal(txtBox3.Text.Replace("$", "").Replace(",", ""));
    //    txtBox3.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", val);
    //}


    protected void LinkButton1_Click(object sender, EventArgs e)
    {

        pdfBillingWorksheet();
    }
    protected void btnDeleteAndCal_Click(object sender, EventArgs e)
    {
        DataSet dsBillingChk = (DataSet)ViewState["CheckDS"];

        DataTable dtCustomeFee = dsBillingChk.Tables[3];

        if (dtCustomeFee != null)
        {
            if (dtCustomeFee.Rows.Count > 0)
            {
                foreach (DataRow row in dtCustomeFee.Rows)
                {
                    string CustFeeID = row["ssi_billingcustomfeeId"].ToString();
                    if (CustFeeID != "")
                        deleteCustomeFee(CustFeeID);

                }
            }
        }

        DataTable dtInvoice = dsBillingChk.Tables[2];

        if (dtInvoice != null)
        {
            if (dtInvoice.Rows.Count > 0)
            {
                foreach (DataRow row in dtInvoice.Rows)
                {
                    string invoiceID = row["ssi_billinginvoiceId"].ToString();
                    if (invoiceID != "")
                        deleteInvoiceFee(invoiceID);

                }
            }
        }
        DataTable dtBillingInvoieFlatfee = dsBillingChk.Tables[7];
        if (dtBillingInvoieFlatfee != null)
        {
            if (dtBillingInvoieFlatfee.Rows.Count > 0)
            {
                foreach (DataRow row in dtBillingInvoieFlatfee.Rows)
                {
                    string BillingInvoieFlatfeeId = row["ssi_BillingInvoiceFlatFeeId"].ToString();
                    if (BillingInvoieFlatfeeId != "")
                        deleteInvoiceFlatFee(BillingInvoieFlatfeeId);

                }
            }
        }


        ClearControls();
        BindClientType();
        CheckandGetExistingData();


        //deleteInvoiceFee();
        //deleteCustomeFee();
    }
    public int GetBillingID()
    {

        string cmdText = "SP_S_BillingInvoice_MaxID";
        SqlConnection conn = null;
        String lsConnectionstring = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
        try
        {

            conn = new SqlConnection(lsConnectionstring);
            conn.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;

            cmd.CommandType = CommandType.StoredProcedure;

            // Add the command parameters to the command object.

            // Execute command 
            string _returnObj = cmd.ExecuteScalar().ToString();
            //outParam = cmd.Parameters[outParamName].Value.ToString();
            cmd.Parameters.Clear();
            cmd.Dispose();
            return Convert.ToInt32(_returnObj);
        }
        finally
        {
            if (conn != null)
                conn.Close();
        }

    }

    #region PDF genration

    public void PdfMerge()
    {
        pdfBillingWorksheet();
        // string vFileInvoice=  PDFInvoic();
    }

    public void pdfBillingWorksheet()
    {
        List<string> list = new List<string>();
        list = (List<string>)ViewState["sourceLsit"];

        DataTable dtBillingforList = (DataTable)ViewState["BillingForList"];

        string sql = "EXEC SP_R_BILLING @ReportFlg = 1,@PdfFlg=1, @BillingForUUID = '";

        //foreach (string UID in SourceFileArray)
        //{
        //    sql = sql + UID + ",";
        //}

        //for (int i = 0; i < list.Count(); i++)
        //{
        //    sql = sql + list[i].ToString() + ",";
        //}
        int rows = dtBillingforList.Rows.Count;
        foreach (DataRow row in dtBillingforList.Rows)
        {
            sql = sql + row["Ssi_billingId"].ToString() + ",";
        }

        sql = sql.Substring(0, sql.LastIndexOf(","));
        sql = sql + "',@AsOfDate = '" + txtAUMDate.Text + "'";

        try
        {

            DataSet newdataset;
            DB clsDB = new DB();
            newdataset = null;
            String DestinationPath = null;
            //  string lsSql = "EXEC SP_R_BILLING @BillingForUUID = '801DE384-D1A4-E511-9418-005056A0567E,841DE384-D1A4-E511-9418-005056A0567E,821DE384-D1A4-E511-9418-005056A0567E',@AsOfDate = '20160331'";
            newdataset = clsDB.getDataSet(sql);
            string HHName = null;
            int noOfPdf = newdataset.Tables.Count;

            if (noOfPdf > 0)
            {

                SourceFileArray = new string[noOfPdf];
                for (int i = 0; i < noOfPdf; i++)
                {
                    if (newdataset.Tables[i].Rows.Count > 0)
                    {
                        DataTable dt;
                        dt = newdataset.Tables[i];
                        DataRow dr = dt.Rows[0];

                        foreach (DataRow row in dt.Rows)
                        {
                            HHName = row["BillingName"].ToString();
                            if (HHName != "")
                                break;
                        }

                        //if (dr["BillingName"].ToString() == "")
                        //{
                        //    HHName = dt.Rows[1].ToString();
                        //}
                        //else
                        //{
                        //    HHName = dt.Rows[0].ToString();
                        //}
                        // HHName = dr["BillingName"].ToString();

                        string path = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + HHName + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString();
                        string filename = GeneratePdf(dt, i);
                        SourceFileArray[i] = filename;
                    }
                }
                // String DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "MergedReport.pdf";

                //bool isExist = System.IO.File.Exists(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + ddlHH.SelectedItem.Text + "-" + ddlBillFor.SelectedItem.Text + System.DateTime.Now.ToString("yyyy-MMdd") + ".pdf");
                //if (isExist)
                //{
                //    System.IO.File.Delete(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + ddlHH.SelectedItem.Text + "-" + ddlBillFor.SelectedItem.Text + System.DateTime.Now.ToString("yyyy-MMdd") + ".pdf");
                //}
                if (rows == 1)
                {
                    //String DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "MergedReport.pdf";
                    DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ddlHH.SelectedItem.Text + "-" + ddlBillFor.SelectedItem.Text + System.DateTime.Now.ToString("yyyy-MMdd") + ".pdf";

                }
                else
                    DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ddlHH.SelectedItem.Text + System.DateTime.Now.ToString("yyyy-MMdd") + ".pdf";

                if (System.IO.File.Exists(DestinationPath))
                {
                    System.IO.File.Delete(DestinationPath);
                }

                if (SourceFileArray.Count() == 1)
                {
                    if (SourceFileArray[0] != null)
                    {
                        MergeFiles(DestinationPath, SourceFileArray);
                    }
                    else
                    {
                        lblMessage.Text = "No Data Found";
                    }
                }

                if (SourceFileArray.Count() > 0)
                {
                    MergeFiles(DestinationPath, SourceFileArray);
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

        }
        catch (Exception ex)
        {
            lblMessage.Text = "error" + ex;
        }
    }

    public void PDFBillingWorksheetAndInvoice()
    {


        List<string> list = new List<string>();
        list = (List<string>)ViewState["sourceLsit"];
        DataTable dtBillingforList = (DataTable)ViewState["BillingForList"];

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

            string vFileInvoice = PDFInvoic();
            int cnt = SourceFileArray.Length;
            if (vFileInvoice != null)
                SourceFileArray[0] = vFileInvoice;

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

                        string path = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + HHName + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString();
                        string filename = GeneratePdf(dt, i);
                        SourceFileArray[i + 1] = filename;
                    }
                }
                // String DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "MergedReport.pdf";
                if (rowCount == 1)
                {
                    //String DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "MergedReport.pdf";
                    DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ddlHH.SelectedItem.Text + "-" + ddlBillFor.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd") + ".pdf";
                    Path1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ddlHH.SelectedItem.Text + "-" + ddlBillFor.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd");
                }
                else
                {
                    DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ddlHH.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd") + ".pdf";
                    Path1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ddlHH.SelectedItem.Text + " " + dAsofDate.ToString("yyyy-MMdd");
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



                    CopytoSharepoint(Path1, DestinationPath);
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

            lblMessage.Text = "Records Updated & File Saved Successfully";
        }
        catch (Exception ex)
        {
            lblMessage.Text = "error" + ex;
        }
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
            ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "billingrpt.pdf";
            //String ls = path;
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();


            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strGUID + ".xls";

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
                    string lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + strGUID + ".pdf";

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



            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.Cache.SetExpires(DateTime.Now.AddSeconds(-1));
            Response.Cache.SetNoStore();
            //Commented  on 22_8_2019 --> ADFS LOGOUT ISSUE
            ////Clear cookies
            //string[] cookies = Request.Cookies.AllKeys;
            //foreach (string cookie in cookies)
            //{
            //    Response.Cookies[cookie].Expires = DateTime.Now.AddDays(-1);
            //}


            Random rnd = new Random();
            string no = rnd.Next().ToString();
            string filename = Path.GetFileNameWithoutExtension(destinationFile);

            FileInfo loFile = new FileInfo(destinationFile);
            filename = filename.Replace("'", "%27");
            loFile.MoveTo(destinationFile.Replace(".xls", ".pdf"));
            //Response.Write("<script>");
            //string lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + filename + ".pdf?" + no;
            ////  Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
            //Response.Write("window.open('" + lsFileNamforFinalXls + "', '_newtab')");
            //Response.Write("</script>");


            //Changed on 22_8_2019 --> ADFS LOGOUT ISSUE
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");
            sb.Append("\nwindow.open('ViewReport.aspx?" + filename + ".pdf" + "', 'mywindow');");
            sb.Append("</script>");
            ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());


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

    #region Common Methods for PDF
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

    #endregion

    public string PDFInvoic()
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
            PdfPTable loTable1 = new PdfPTable(6);
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
            int[] headerwidths = { 18, 10, 10, 13, 13, 5 };
            loTable1.WidthPercentage = 85f;
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
            int colsize1 = 6;


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

                        //if (k == 0)
                        //{ ColHeader = ""; }


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

                    if ((j == 1 || j == 3 || j == 4) && ColValue != "")
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
                    else if ((j == 2 || j == 5) && ColValue != "")
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
    protected void Button1_Click(object sender, EventArgs e)
    {
        //   updateInvoicewithTotalandPer();
        //   UpdateNotes(); 

        //   PDFInvoic();
        // pdfNew();

        //   pdfBillingWorksheet();
        PDFBillingWorksheetAndInvoice();
    }
    protected void txtTotalBillingAUM_TextChanged(object sender, EventArgs e)
    {
        string BillingAUM = txtBillingAUM.Text;
        BillingAUM = BillingAUM.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
        string TotalbillingAUM = "0";


        TotalbillingAUM = txtTotalBillingAUM.Text;
        // TotalbillingAUM = TotalbillingAUM.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
        txtBillingAUM.Text = TotalbillingAUM.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");


        string totalAUM = txtTotalBillingAUM.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
        txtTotalBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(totalAUM));

        BillingAUMChange();
    }

    public void gridviewBind(DataTable dtData)
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("ssi_Comments");
        dt.Columns.Add("Fees");
        //for (int i = 0; i < 10; i++)
        //{
        //    DataRow row = dt.NewRow();
        //    row["Colname"] = "RelationshipFee";
        //    row["Value"] = "1000";
        //    dt.Rows.Add(row);

        //}

        foreach (DataRow row in dtData.Rows)
        {
            DataRow dtRow = dt.NewRow();
            dtRow["ssi_Comments"] = row["ssi_Comments"].ToString();
            string val = row["Fees"].ToString();
            dtRow["Fees"] = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(val));
            dt.Rows.Add(dtRow);
        }

        gvFlatFee.Visible = true;
        gvFlatFee.DataSource = dt;

        gvFlatFee.DataBind();




    }

    //protected void Button1_Click1(object sender, EventArgs e)
    //{
    //    PDFBillingWorksheetAndInvoice();
    //}

    public void BillingAUMChange()
    {

        /* Replace , to $ for BillingAum*/
        string value = txtBillingAUM.Text.Replace(",", "").Replace("$", "");


        if (txtBillingAUM.Text.Trim() != "")
        {
            DB clsDB = new DB();//class Library
            /* Return AnnualFee  into dataset where Billing Amount and clientType Selected */
            DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@TotalAUMNmb='" + txtTotalAUM.Text.Replace(",", "").Replace("$", "") + "' ");
            if (dsBilling.Tables[0].Rows.Count > 0)
            {
                /* check AnnualFee if Null set AnnualFee="0.0" else display AnnualFee */
                string AnnualFee = Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]);

                /* used Globalization To Format StdAnnualFeeCalc in (en-US) with 2 decimal */
                txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));

                ViewState["BillingId"] = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingId"]);
            }

            CalulateFeeRate();//function to calculateFeeRate
            decimal ul;

            /*->out keyword used to display billingAUM
            ->TryParse used to convert into Decimal*/
            if (decimal.TryParse(value, out ul))
            {
                txtBillingAUM.TextChanged -= txtBillingAUM_TextChanged;
                txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                txtBillingAUM.TextChanged += txtBillingAUM_TextChanged;
            }
        }
        else
        {

            /* clear Following Control*/
            txtCustFeeAmount.Text = "";
            txtFeeRateCalc.Text = "";
            txtFeesPerMonth.Text = "";
            txtQuaterlyFeeCalc.Text = "";
            txtStdAnnualFeeCalc.Text = "";
        }

        if (ddlClientType.SelectedValue == "100000001") //Standard with Relationship Fees
            txtRelationshipFee.Focus();
        else
            txtCustFeeAmount.Focus();
    }


    protected void btnStandardCalculate_Click(object sender, EventArgs e)
    {
        // lblMessage.Visible = false;
        lblMessage.Text = "";
        string value = txtBillingAUM.Text.Replace(",", "").Replace("$", "");




        if (txtBillingAUM.Text.Trim() != "")
        {

            ViewState["bpsfeeValue"] = txtbpsfee.Text;
            ViewState["MinValue"] = txtMinVal.Text;
            ViewState["MaxValue"] = txtMaxVal.Text;
            //object BillingAUM = txtBillingAUM.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "'";           
            object txtbpsfeeValue = txtbpsfee.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtbpsfee.Text.Replace(",", "").Replace("$", "") + "'";
            string txtMinValue = txtMinVal.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtMinVal.Text.Replace(",", "").Replace("$", "") + "'";
            object txtMaxValue = txtMaxVal.Text.Replace(",", "").Replace("$", "") == "" ? "null" : "'" + txtMaxVal.Text.Replace(",", "").Replace("%", "") + "'";

            DB clsDB = new DB();//class Library
                                /* Return AnnualFee  into dataset where Billing Amount and clientType Selected */
                                //  DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "' ");
                                //DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + "");
                                //double relatioshipfee = 0.0;
                                //double trustfee = 0.0;
                                //double administrativefee = 0.0;
                                //double servicefee = 0.0;
                                //double setupfee = 0.0;
                                //double transactionfee = 0.0;
            double securityfee = 0.0;
            //string ssi_comments = string.Empty;
            //string FeeValue = string.Empty;
            //foreach (GridViewRow row in gvFlatFee.Rows)
            //{

            //    System.Web.UI.WebControls.Label relatioshipfee1 = (System.Web.UI.WebControls.Label)row.FindControl("lblFlatFee");
            //    System.Web.UI.WebControls.TextBox txtfeevalue = (System.Web.UI.WebControls.TextBox)row.FindControl("txtFlatFee");



            //    ssi_comments = relatioshipfee1.Text;
            //    FeeValue = txtfeevalue.Text;


            //    if (ssi_comments.ToLower().Trim() == "relationship fee")
            //        relatioshipfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "trust fee")
            //        trustfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "administrative service fee")
            //        administrativefee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "service fee")
            //        servicefee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "setup fee")
            //        setupfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));
            //    if (ssi_comments.ToLower().Trim() == "transaction fee")
            //        transactionfee = Convert.ToDouble(FeeValue.Replace(",", "").Replace("$", ""));

            //}



            if (txtSecurityFee.Text != "")
                securityfee = Convert.ToDouble(txtSecurityFee.Text.Replace(",", "").Replace("$", ""));

            DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING_ANNUALFEECALC @TotalAUMNmb='" + txtTotalAUM.Text.Replace(",", "").Replace("$", "") + "',@BillingAumAmount='" + txtBillingAUM.Text.Replace(",", "").Replace("$", "") + "',@ClientType='" + ddlClientType.SelectedItem.Text + "',@CustomBPSNum=" + txtbpsfeeValue + ",@CustomMinAmount=" + txtMinValue + ",@CustomMaxPct=" + txtMaxValue + ",@reCalculateFlg =1,@SecurityFeeNum = " + securityfee);
            if (dsBilling.Tables[0].Rows.Count > 0)
            {
                /* check AnnualFee if Null set AnnualFee="0.0" else display AnnualFee */
                string AnnualFee = Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]) == "" ? "0.0" : Convert.ToString(dsBilling.Tables[0].Rows[0]["AnnualFee"]);

                /* used Globalization To Format StdAnnualFeeCalc in (en-US) with 2 decimal */
                txtStdAnnualFeeCalc.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(AnnualFee));

                ViewState["BillingId"] = Convert.ToString(dsBilling.Tables[0].Rows[0]["BillingId"]);
            }

            CalulateFeeRate();//function to calculateFeeRate
            decimal ul;

            /*->out keyword used to display billingAUM
            ->TryParse used to convert into Decimal*/
            if (decimal.TryParse(value, out ul))
            {
                txtBillingAUM.TextChanged -= txtBillingAUM_TextChanged;
                txtBillingAUM.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", ul);
                txtBillingAUM.TextChanged += txtBillingAUM_TextChanged;
            }
        }
        else
        {

            /* clear Following Control*/
            txtCustFeeAmount.Text = "";
            txtFeeRateCalc.Text = "";
            txtFeesPerMonth.Text = "";
            txtQuaterlyFeeCalc.Text = "";
            txtStdAnnualFeeCalc.Text = "";
        }
    }

    protected void txtMaxValue_TextChanged(object sender, EventArgs e)
    {
        // lblMessage.Visible = false;
        lblMessage.Text = "";
        /* check txtMaxValue not Equals To Null*/
        if (txtMaxVal.Text != "")
        {
            double doubleValue;
            /* Convert txtMaxValue into double*/
            if (double.TryParse(txtMaxVal.Text, out doubleValue))
            {
                doubleValue = doubleValue / 100;
                txtMaxVal.Text = doubleValue.ToString("0.00%");//Convert doublevalue into String and Display in Percentage 
            }
        }
    }

    protected void txtMinVal_TextChanged(object sender, EventArgs e)
    {
        // lblMessage.Visible = false;
        lblMessage.Text = "";
        if (txtMinVal.Text != "")
        {
            string totalAUM = txtMinVal.Text.Replace("$", "").Replace(",", "").Replace("(", "-").Replace(")", "");
            txtMinVal.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(totalAUM));
        }
    }




}
