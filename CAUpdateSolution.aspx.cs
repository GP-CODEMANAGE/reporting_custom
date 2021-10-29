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
using System.Data.Common;
//using CrmSdk;
using System.IO;
using System.Drawing;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.Configuration;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;

using Microsoft.SharePoint.Client;
using System.Security;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Text;
using GemBox.Document;
using GemBox.Document.Tables;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Xls;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class CAUpdateSolution : System.Web.UI.Page
{
    ClientContext context;
    public StreamWriter sw = null;
    bool bProceed = true;
    string strDescription;
    string greshamquery;
    int totalCount = 0;
    int successcount = 0;
    // int successcount = 0;]
    GeneralMethods clsGM = new GeneralMethods();
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    String fsReportingName = "";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName, lsGAorTIAHeader;
    string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);//"Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";



    protected void Page_Load(object sender, EventArgs e)
    {
        IOrganizationService service1 = null;
        service1 = clsGM.GetCrmService();

        //Response.Cache.SetCacheability(HttpCacheability.NoCache);
        // Response.Clear();

        //Response.AddHeader("cache-control", "must-revalidate");

        //clearIECache();

        #region Commented
        /*
        //string LogFileName = "LogFile " + DateTime.Now;
        //LogFileName = LogFileName.Replace(":", "-");
        //LogFileName = LogFileName.Replace("/", "-");
        //sw = new StreamWriter(Request.PhysicalApplicationPath + "\\" + LogFileName + ".txt", true);

        string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

        string crmServerUrl = "http://Crm01/";
        //string crmServerURL = "http://server:5555/";

        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        CrmService service = null;

        try
        {
            service = GetCrmService(crmServerUrl, orgName);
            strDescription = "Crm Service starts successfully";
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
            // sw.WriteLine("step 1 ");
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }

        service.PreAuthenticate = true;
        service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        */
        #endregion
        HideControls();
        if (!IsPostBack)
        {
            Session.Remove("CurPageInBatch");
            BindFrequency();
            BindSecondaryOwner();
            BindHousehold();
            BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");
            FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");

            //GetTransactionID();
        }

        //Page.ClientScript.RegisterStartupScript(this.GetType(), "cle", "windows.history.clear", true);


        #region Commented
        //SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        //SqlCommand cmd = new SqlCommand();
        //SqlDataAdapter dagersham = new SqlDataAdapter();
        //SqlDataAdapter da_CRM;
        //DataSet ds_gresham = new DataSet();
        //DataSet ds = new DataSet();

        //string sqlstr = "[SP_S_CAUpdated_HouseHold]";
        //dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        //ds_gresham = new DataSet();
        //dagersham.Fill(ds);

        //ddlHousehold.DataTextField = "name";
        //ddlHousehold.DataValueField = "accountid";

        //ddlHousehold.DataSource = ds;
        //ddlHousehold.DataBind();
        //ds.Dispose();


        /*
        try
        {
            greshamquery = "exec SP_S_Position_CA_RollForward_Update '440B713D-6A15-DE11-8391-001D09665E8F'";
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;
            sw.WriteLine("----------------------------  Position Update Starts -------------------");
            sw.WriteLine("Batch: " + DateTime.Now.ToString());
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            totalCount = 0;
            sw.WriteLine(" error desc:" + exc.Detail.InnerText);
            // LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }
        catch (Exception exc)
        {
            bProceed = false;
            totalCount = 0;
            sw.WriteLine("error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }


        gvList.Columns[0].Visible = true;
        gvList.Columns[1].Visible = true;
        gvList.DataSource = ds_gresham;
        gvList.DataBind();
        gvList.Columns[0].Visible = false;
        gvList.Columns[1].Visible = false;

        lblUpdateMonth.Text = Convert.ToString(Convert.ToDateTime(ds_gresham.Tables[0].Rows[0]["ssi_asofdate"]).ToShortDateString());
         * */
        #endregion
    }

    private void BindFrequency()
    {
        //string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        string sqlstr = "[SP_S_ACCOUNT_CA_FREQUENCY]";
        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);

        ddlFrequency.DataTextField = "Freq_Name";
        ddlFrequency.DataValueField = "Freq_Name_ID";

        ddlFrequency.DataSource = ds;
        ddlFrequency.DataBind();

        ddlFrequency.Items.Insert(0, "All");
        ddlFrequency.Items[0].Value = "";
        ddlFrequency.SelectedIndex = 0;
    }

    private void BindHousehold()
    {
        //Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        object SecondaryOwnerId = ddlSecOwner.SelectedValue == "0" ? "null" : "'" + ddlSecOwner.SelectedValue + "'";

        string sqlstr = "[SP_S_CAUpdated_HouseHold] @SecondaryOwnerId=" + SecondaryOwnerId + ",@CurrentFlg='1,100000000'";
        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);

        ddlHousehold.DataTextField = "name";
        ddlHousehold.DataValueField = "accountid";

        ddlHousehold.DataSource = ds;
        ddlHousehold.DataBind();


    }

    private void BindSecondaryOwner()
    {
        //string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        string sqlstr = "[SP_S_CAUpdated_HouseHold_SecondaryOwner]";
        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);

        ddlSecOwner.DataTextField = "Ssi_SecondaryOwnerIdName";
        ddlSecOwner.DataValueField = "Ssi_SecondaryOwnerId";

        ddlSecOwner.DataSource = ds;
        ddlSecOwner.DataBind();

        ddlSecOwner.Items.Insert(0, "All");
        ddlSecOwner.Items[0].Value = "0";
        ddlSecOwner.SelectedIndex = 0;


    }

    private void FillGridForCashinTransit(string HouseHold, string AsOfDate)
    {
        //string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery = string.Empty;

        try
        {
            HouseHold = HouseHold == "''" ? "null" : HouseHold;
            AsOfDate = AsOfDate == "''" ? "null" : AsOfDate;
            greshamquery = "exec SP_S_POSITION_CA_ROLLFORWARD_WITH_CASH_IN_TRANSIT_UPDATE @HouseHoldId=" + HouseHold + ",@AsOfDate=" + AsOfDate;
            //Response.Write(greshamquery);
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            // totalCount = ds_gresham.Tables[0].Rows.Count;
            //sw.WriteLine("----------------------------  Position Update Starts -------------------");
            //sw.WriteLine("Batch: " + DateTime.Now.ToString());
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {

            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Detail.InnerText;
            // bProceed = false;
            //totalCount = 0;
            // sw.WriteLine(" error desc:" + exc.Detail.InnerText);
            // LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }
        catch (Exception exc)
        {
            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Message;
            //bProceed = false;
            // totalCount = 0;
            //sw.WriteLine("error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }

        gvCashinTransit.Columns[12].Visible = true;
        gvCashinTransit.Columns[13].Visible = true;
        gvCashinTransit.Columns[14].Visible = true;
        gvCashinTransit.Columns[15].Visible = true;
        gvCashinTransit.Columns[16].Visible = true;
        gvCashinTransit.Columns[17].Visible = true;
        gvCashinTransit.Columns[18].Visible = true;
        gvCashinTransit.Columns[19].Visible = true;
        gvCashinTransit.Columns[20].Visible = true;
        gvCashinTransit.Columns[21].Visible = true;
        gvCashinTransit.Columns[22].Visible = true;
        gvCashinTransit.Columns[23].Visible = true;
        gvCashinTransit.Columns[24].Visible = true;
        gvCashinTransit.Columns[25].Visible = true;
        gvCashinTransit.Columns[26].Visible = true;
        gvCashinTransit.Columns[27].Visible = true;
        gvCashinTransit.Columns[28].Visible = true;
        gvCashinTransit.Columns[29].Visible = true;

        gvCashinTransit.DataSource = ds_gresham;
        gvCashinTransit.DataBind();

        gvCashinTransit.Columns[12].Visible = false;
        gvCashinTransit.Columns[13].Visible = false;
        gvCashinTransit.Columns[14].Visible = false;
        gvCashinTransit.Columns[15].Visible = false;
        gvCashinTransit.Columns[16].Visible = false;
        gvCashinTransit.Columns[17].Visible = false;
        gvCashinTransit.Columns[18].Visible = false;
        gvCashinTransit.Columns[19].Visible = false;
        gvCashinTransit.Columns[20].Visible = false;
        gvCashinTransit.Columns[21].Visible = false;
        gvCashinTransit.Columns[22].Visible = false;
        gvCashinTransit.Columns[23].Visible = false;
        gvCashinTransit.Columns[24].Visible = false;
        gvCashinTransit.Columns[25].Visible = false;
        gvCashinTransit.Columns[26].Visible = false;
        gvCashinTransit.Columns[27].Visible = false;
        gvCashinTransit.Columns[28].Visible = false;
        gvCashinTransit.Columns[29].Visible = false;

    }


    private void BindGridView(string HouseholdId, string AsOfDate, string Freqency, string View)
    {
        //string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery = string.Empty;

        try
        {
            HouseholdId = HouseholdId == "''" ? "null" : HouseholdId;
            AsOfDate = AsOfDate == "''" ? "null" : AsOfDate;
            Freqency = Freqency == "''" ? "null" : Freqency;
            View = View == "''" ? "null" : View;

            greshamquery = "exec SP_S_Position_CA_RollForward_Update @HouseHoldId=" + HouseholdId + ",@AsOfDate=" + AsOfDate + ",@Frequency=" + Freqency + ",@View=" + View;
            //Response.Write(greshamquery);
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();

            dagersham.Fill(ds_gresham);


            // totalCount = ds_gresham.Tables[0].Rows.Count;
            //sw.WriteLine("----------------------------  Position Update Starts -------------------");
            //sw.WriteLine("Batch: " + DateTime.Now.ToString());
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {

            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Detail.InnerText;
            // bProceed = false;
            //totalCount = 0;
            // sw.WriteLine(" error desc:" + exc.Detail.InnerText);
            // LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }
        catch (Exception exc)
        {
            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Message;
            //bProceed = false;
            // totalCount = 0;
            //sw.WriteLine("error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }

        gvList.Columns[0].Visible = true;
        gvList.Columns[1].Visible = true;
        gvList.Columns[16].Visible = true;
        gvList.Columns[17].Visible = true;
        gvList.Columns[18].Visible = true;
        gvList.Columns[19].Visible = true;
        gvList.Columns[20].Visible = true;
        // gvList.Columns[21].Visible = true;
        gvList.DataSource = ds_gresham;
        gvList.DataBind();
        gvList.Columns[0].Visible = false;
        gvList.Columns[1].Visible = false;
        gvList.Columns[16].Visible = false;
        gvList.Columns[17].Visible = false;
        gvList.Columns[18].Visible = false;
        gvList.Columns[19].Visible = false;
        gvList.Columns[20].Visible = false;


        // gvList.Columns[21].Visible = false;

        if (gvList.Rows.Count > 0)
        {

            if (Convert.ToString(ds_gresham.Tables[0].Rows[0]["ssi_asofdate"]) != "")
            {
                System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
                if (Convert.ToString(ViewState["asofdate"]) == "")
                    ViewState["asofdate"] = Convert.ToString(Convert.ToDateTime(ds_gresham.Tables[0].Rows[0]["ssi_asofdate"], enUS), enUS);
                BindDate(ViewState["asofdate"].ToString());
            }
            //-   lblUpdateMonth.Text = Convert.ToString(Convert.ToDateTime(ds_gresham.Tables[0].Rows[0]["ssi_asofdate"]).ToShortDateString());
            else
            {
                ddlUpdateMonth.Items.Clear();
                ddlUpdateMonth.Items.Add("N/A");
            }

            string sqlstr = "[SP_S_CAUpdated_HouseHold] @HouseHoldId=" + HouseholdId;
            dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds);

            if (ds.Tables[0].Rows.Count > 0)
            {
                chkbCAUpcateCompleted.Checked = Convert.ToString(ds.Tables[0].Rows[0]["Ssi_CAUpdateCompleted"]).ToLower() == "true" ? true : false;
                Hidden1.Value = Convert.ToString(ds.Tables[0].Rows[0]["Ssi_CAUpdateCompleted"]).ToLower();
            }
            else
                chkbCAUpcateCompleted.Checked = false;

            btnSubmit.Visible = true;
            btnSumbitTop.Visible = true;
        }
        else
        {
            //- lblUpdateMonth.Text = "No Record found";
            chkbCAUpcateCompleted.Checked = false;

            //   btnSubmit.Visible = false;
            //   btnSumbitTop.Visible = false;
        }
    }

    private void BindDate(string AsOfdate)
    {
        string CurValue = Convert.ToString(ddlUpdateMonth.SelectedValue);

        ddlUpdateMonth.Items.Clear();
        System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
        DateTime dt1 = Convert.ToDateTime(AsOfdate, enUS);
        ddlUpdateMonth.Items.Add(Convert.ToString(dt1.ToShortDateString()));
        int Month = dt1.Month;
        for (int i = 0; i <= 2; i++)
        {
            DateTime dt2 = Convert.ToDateTime(Month + "/" + "01/" + dt1.Year, enUS);
            DateTime dt3 = dt2.AddDays(-1);
            dt1 = dt3;
            if (Month == 1)
                Month = 12;
            else
                Month--;
            ddlUpdateMonth.Items.Add(Convert.ToString(dt3.ToShortDateString()));
        }

        if (CurValue != "")
            ddlUpdateMonth.SelectedValue = CurValue;


    }

    protected void gvList_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.Header)
        {
            //Find the checkbox control in header and add an attribute
            ((CheckBox)e.Row.FindControl("chkbxNCSelectAll")).Attributes.Add("onclick", "javascript:SelectAll('" +
                    ((CheckBox)e.Row.FindControl("chkbxNCSelectAll")).ClientID + "')");
        }


        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            System.Web.UI.WebControls.TextBox txtCAUpdateValue = (System.Web.UI.WebControls.TextBox)e.Row.FindControl("txtCAUpdateValue");

            CheckBox chkbxNC1 = (CheckBox)e.Row.FindControl("chkbNC");

            //  HtmlAnchor lnkeditCommitment = (HtmlAnchor)e.Row.FindControl("lnkedit");

            LinkButton lnkeditCommitment = (LinkButton)e.Row.FindControl("lnkedit");

            chkbxNC1.Attributes.Add("onclick", "EnableDisable('" + chkbxNC1.ClientID + "','" + txtCAUpdateValue.ClientID + "')");

            string _Commitment = DataBinder.Eval(e.Row.DataItem, "Commitment").ToString();
            string _PositionId = DataBinder.Eval(e.Row.DataItem, "ssi_positionid").ToString();

            //old logic changed on 06/23/2016 Modifiedon changed to Source
            string Modifiedon = DataBinder.Eval(e.Row.DataItem, "Modifiedon").ToString();
            string Source = DataBinder.Eval(e.Row.DataItem, "dataSource").ToString();
            string Freq = DataBinder.Eval(e.Row.DataItem, "Ssi_CAUpdateFreq").ToString();
            string ddl_UpdateMonth = ddlUpdateMonth.SelectedValue;

            //lblMessage.Text = Convert.ToString( diffh);



            if (_Commitment == "1")
            {
                lnkeditCommitment.Visible = true;
                //   lnkeditCommitment.Attributes["onclick"] = "OpenChild('" + _PositionId + "');";


            }
            else if (_Commitment == "0")
            {
                lnkeditCommitment.Visible = false;
            }

            if (e.Row.RowIndex > 1)
            {
                if (e.Row.Cells[1].Text == gvList.Rows[e.Row.RowIndex - 1].Cells[1].Text)
                {
                    //gvList.Rows[e.Row.RowIndex - 1].Style["border-top-color"] = "Red";
                    //gvList.Rows[e.Row.RowIndex - 1].Style["border-top"] = "solid";
                    //gvList.Rows[e.Row.RowIndex - 1].Style["border-top-width"] = "thick";
                    // e.Row.Style.Add( border-top-color:Gray; border-top:solid; border-top-width:thick;
                    e.Row.Cells[2].Text = "";
                    e.Row.Cells[3].Text = "";
                    e.Row.Cells[4].Text = "";
                }
                else
                {

                }




            }

            if (e.Row.RowIndex > 0)
                if (e.Row.Cells[1].Text != gvList.Rows[e.Row.RowIndex - 1].Cells[1].Text)
                {
                    for (int i = 2; i < gvList.Columns.Count; i++)
                    {
                        e.Row.Cells[i].Style["border-style"] = "solid";
                        e.Row.Cells[i].Style["border-top-color"] = "#D8D8D8";
                        e.Row.Cells[i].Style["border-top-width"] = "3px";
                    }
                }





            if (e.Row.Cells[0].Text == "&nbsp;")
            {
                // Response.Write("pos1:" + gvList.Rows[e.Row.RowIndex].Cells[0].Text);

                txtCAUpdateValue.Visible = false;
                chkbxNC1.Visible = false;
            }


            if (e.Row.Cells[2].Text != "")
            {
                e.Row.BackColor = System.Drawing.Color.LightGray;

            }

            if (Modifiedon != "" && ddlUpdateMonth.SelectedValue != "")
            {
                //new logic 06/23/2016 as per ddl_UpdateMonth
                // string ddl_UpdateMonth = ddlUpdateMonth.Text.ToString();

                DateTime UpdateMonth = Convert.ToDateTime(ddlUpdateMonth.SelectedValue);

                //old logic as per modified date
                DateTime modifiedon = Convert.ToDateTime(Modifiedon);
                // DateTime modifiedon = new DateTime(5/10/2015);
                DateTime to = UpdateMonth.AddMonths(-1);
                DateTime to1 = UpdateMonth.AddMonths(-3);
                DateTime to2 = UpdateMonth.AddMonths(-6);
                DateTime to3 = UpdateMonth.AddMonths(-12);

                //int diff = DateTime.Compare(modifiedon, to);
                double diffh = (to - modifiedon).TotalDays;
                double diffh1 = (to1 - modifiedon).TotalDays;
                double diffh2 = (to2 - modifiedon).TotalDays;
                double diffh3 = (to3 - modifiedon).TotalDays;

                //for ALL Frequency
                if (ddlFrequency.SelectedItem.Text.ToUpper() == "ALL".ToUpper())
                {

                    if (Freq == "1" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }
                    if (Freq == "2" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh1 >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }
                    if (Freq == "3" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh2 >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }
                    if (Freq == "4" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh3 >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }




                }
                //For Monthly Frequency
                else if (ddlFrequency.SelectedItem.Text.ToUpper() == "Monthly".ToUpper() || ddlFrequency.SelectedItem.Text.ToUpper() == "Monthly (Lagged)".ToUpper())
                {
                    if (Freq == "1" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }
                }

                //For Quaterly Frequency
                else if (ddlFrequency.SelectedItem.Text.ToUpper() == "Quarterly".ToUpper() || ddlFrequency.SelectedItem.Text.ToUpper() == "Quarterly (Lagged)".ToUpper())
                {
                    if (Freq == "2" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh1 >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }
                }

                    //Semi Annualy Frequency
                else if (ddlFrequency.SelectedItem.Text.ToUpper() == "Semi Annually".ToUpper())
                {
                    // double month = (modifiedon - today).Days / (30 * 6);
                    if (Freq == "3" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh2 >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }
                }
                //Annualy Frequency
                else if (ddlFrequency.SelectedItem.Text.ToUpper() == "Annually".ToUpper())
                {
                    if (Freq == "4" && (Source.ToUpper() != "CA Updates".ToUpper() && Source.ToUpper() != "Manual".ToUpper()))
                    {
                        if (diffh3 >= 0)
                        {
                            if (e.Row.Cells[10].Text != "")
                            {
                                e.Row.BackColor = System.Drawing.Color.FromName("#FCFB9C");
                            }
                        }
                    }
                }
            }
        }
    }


    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");
        ddlUpdateMonth.ClearSelection();
        ddlFrequency.ClearSelection();
        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
    }

    private void CashinTransitUpdate(string UserId)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;
        lblMessage.Text = "";

        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
            // sw.WriteLine("step 1 ");
        }
        // catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }


        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        foreach (GridViewRow row in gvCashinTransit.Rows)
        {

            System.Web.UI.WebControls.TextBox txtCashinTransitUpdate = (System.Web.UI.WebControls.TextBox)row.FindControl("txtCashinTransit");
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            string ssi_PositionId = row.Cells[12].Text.Trim().Replace("ssi_positionid", "").Replace("&nbsp;", "");
            string AccName = row.Cells[13].Text.Trim().Replace("Account", "").Replace("&nbsp;", "");
            string AccountId = row.Cells[23].Text.Trim().Replace("ssi_accountid", "").Replace("&nbsp;", "");
            string _800Acctflg = row.Cells[18].Text.Trim().Replace("_800Acctflg", "").Replace("&nbsp;", "");
            string AsOfDate = row.Cells[16].Text.Trim().Replace("asofdate", "").Replace("&nbsp;", "").Replace("As Of Date", "");

            string AssetClassId = row.Cells[24].Text.Trim().Replace("sas_assetclassid ", "").Replace("&nbsp;", "");
            string subassetclassId = row.Cells[25].Text.Trim().Replace("Ssi_subassetclassId", "").Replace("&nbsp;", "");
            string BenchmarkSubAssetClassId = row.Cells[26].Text.Trim().Replace("Ssi_BenchmarkSubAssetClassId ", "").Replace("&nbsp;", "");
            string SectorFlg = row.Cells[27].Text.Trim().Replace("SectorFlg ", "").Replace("&nbsp;", "");

            try
            {

                // ssi_position objPosition = new ssi_position();
                Entity objPosition = new Entity("ssi_position");
                if (_800Acctflg.ToUpper() == "TRUE")
                {
                    if (chkSelectNC.Checked == false)
                    {
                        if (ssi_PositionId != "")
                        {
                            //objPosition.ssi_positionid = new Key();
                            //Guid Myguid = new Guid(ssi_PositionId);
                            //objPosition.ssi_positionid.Value = Myguid;

                            objPosition["ssi_positionid"] = new Guid(ssi_PositionId);
                        }

                        //objPosition.ssi_datasource = new Picklist();
                        //objPosition.ssi_datasource.Value = 6; // value of CA Update
                        objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);


                        if (txtCashinTransitUpdate.Text != "")
                        {
                            //objPosition.ssi_marketvalue = new CrmMoney();
                            //objPosition.ssi_marketvalue.Value = Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim());
                            objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim()));


                            //objPosition.ssi_quantity = new CrmDecimal();
                            //objPosition.ssi_quantity.Value = Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim());
                            objPosition["ssi_quantity"] = Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim());

                        }

                        /************************* New GA Logic **********************************/

                        //SubAssetclassid 
                        if (subassetclassId != "")
                        {
                            //objPosition.ssi_subassetclassid = new Lookup();
                            //objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                            //objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(subassetclassId));
                            objPosition["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(subassetclassId)));
                        }
                        else
                        {
                            //objPosition.ssi_subassetclassid = new Lookup();
                            //objPosition.ssi_subassetclassid.IsNull = true;
                            //objPosition.ssi_subassetclassid.IsNullSpecified = true;
                            objPosition["ssi_subassetclassid"] = null;
                        }

                        //SubAssetclassid 
                        if (BenchmarkSubAssetClassId != "")
                        {
                            //objPosition.ssi_benchmarkid = new Lookup();
                            //objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                            //objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(BenchmarkSubAssetClassId));
                            objPosition["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(BenchmarkSubAssetClassId)));
                        }
                        else
                        {
                            //objPosition.ssi_benchmarkid = new Lookup();
                            //objPosition.ssi_benchmarkid.IsNull = true;
                            //objPosition.ssi_benchmarkid.IsNullSpecified = true;
                            objPosition["ssi_benchmarkid"] = null;
                        }

                        //Assetclassid 
                        if (AssetClassId != "")
                        {
                            //objPosition.ssi_assetclassid = new Lookup();
                            //objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                            //objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(AssetClassId));
                            objPosition["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(AssetClassId)));
                        }
                        else
                        {
                            //objPosition.ssi_assetclassid = new Lookup();
                            //objPosition.ssi_assetclassid.IsNull = true;
                            //objPosition.ssi_assetclassid.IsNullSpecified = true;
                            objPosition["ssi_assetclassid"] = null;
                        }

                        //greshamadvised (sectorflg)
                        if (SectorFlg != "")
                        {
                            //objPosition.ssi_greshamadvised = new CrmBoolean();
                            //objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(SectorFlg);
                            objPosition["ssi_greshamadvised"] = Convert.ToBoolean(SectorFlg);

                        }

                        /************************* End New GA Logic *****************************/

                        if (txtCashinTransitUpdate.Text != "" && ssi_PositionId != "")
                        {
                            service.Update(objPosition);
                            successcount++;
                        }
                        else if (ssi_PositionId == "")
                        {
                            if (txtCashinTransitUpdate.Text != "")
                            {
                                //objPosition.transactioncurrencyid = new Lookup();
                                //objPosition.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                //objPosition.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                                objPosition["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));

                                //objPosition.ssi_assetclassid = new Lookup();
                                //objPosition.ssi_assetclassid.type = EntityName.ssi_assetclassallocation.ToString();
                                //objPosition.ssi_assetclassid.Value = new Guid("9776259D-0392-4DE0-8A12-0399724ABF8D");

                                //objPosition.ssi_asofdate = new CrmDateTime();
                                //objPosition.ssi_asofdate.Value = AsOfDate;
                                objPosition["ssi_asofdate"] = Convert.ToDateTime(AsOfDate);

                                //objPosition.ssi_enddate = new CrmDateTime();
                                //objPosition.ssi_enddate.Value = AsOfDate;
                                objPosition["ssi_enddate"] = Convert.ToDateTime(AsOfDate);

                                //objPosition.ssi_securityid = new Lookup();
                                //objPosition.ssi_securityid.type = EntityName.ssi_security.ToString();
                                //objPosition.ssi_securityid.Value = new Guid("47C9F791-5E66-446B-BA73-E614EC3A5137");
                                objPosition["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid("47C9F791-5E66-446B-BA73-E614EC3A5137"));

                                //objPosition.ssi_datasource = new Picklist();
                                //objPosition.ssi_datasource.Value = 6; // value of CA Update
                                objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);


                                //objPosition.ssi_approveflg = new CrmBoolean();
                                //objPosition.ssi_approveflg.Value = true;
                                objPosition["ssi_approveflg"] = true;


                                if (_800Acctflg == "True")
                                {
                                    //objPosition.ssi_accountid = new Lookup();
                                    //objPosition.ssi_accountid.type = EntityName.ssi_account.ToString();
                                    //objPosition.ssi_accountid.Value = new Guid(AccountId);
                                    objPosition["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(AccountId)));

                                }

                                if (txtCashinTransitUpdate.Text != "")
                                {
                                    //objPosition.ssi_marketvalue = new CrmMoney();
                                    //objPosition.ssi_marketvalue.Value = Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim());
                                    objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim()));


                                    //objPosition.ssi_quantity = new CrmDecimal();
                                    //objPosition.ssi_quantity.Value = Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim());
                                    objPosition["ssi_quantity"] = Convert.ToDecimal(txtCashinTransitUpdate.Text.Trim());

                                }

                                service.Create(objPosition);
                                successcount++;
                            }

                        }
                    }
                }
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                lblMessage.Text = strDescription;
                //  sw.WriteLine(strDescription);
                //LogMessage(sw, service, strDescription, 62, "GeneralError");
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                lblMessage.Text = strDescription;
                //  sw.WriteLine(strDescription);
                //LogMessage(sw, service, strDescription, 62, "GeneralError");
            }


        }

        //if (successcount > 0)
        //    lblMessage.Text = lblMessage.Text + "<br/>" + successcount.ToString() + " Records updated successfully for CASH IN TRANSIT UPDATES";
        //else
        //    lblMessage.Text = lblMessage.Text + "<br/> No Records updated for CASH IN TRANSIT UPDATES";


    }


    private void FillGridForCashinTransitPopUP(string Position)
    {
        //string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery = string.Empty;

        try
        {
            Position = Position == "''" ? "null" : Position;
            greshamquery = "EXECUTE [dbo].[SP_S_Position_CA_Commitment_Update] @PositionID =" + Position;
            //Response.Write(greshamquery);
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            // totalCount = ds_gresham.Tables[0].Rows.Count;
            //sw.WriteLine("----------------------------  Position Update Starts -------------------");
            //sw.WriteLine("Batch: " + DateTime.Now.ToString());
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {

            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Detail.InnerText;
            // bProceed = false;
            //totalCount = 0;
            // sw.WriteLine(" error desc:" + exc.Detail.InnerText);
            // LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }
        catch (Exception exc)
        {
            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Message;
            //bProceed = false;
            // totalCount = 0;
            //sw.WriteLine("error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }


        if (ds_gresham.Tables.Count > 0)
        {
            gvpopup.DataSource = ds_gresham;
            gvpopup.DataBind();
        }
        if (gvpopup.Rows.Count > 0)
        {
            lblMessage.Text = "";
            btnSubmit.Visible = true;
        }
        else
        {
            performancepanel.Visible = false;
            performancepopup.Hide();
            lblMessage.Text = "No records found.";
            btnSubmit.Visible = false;
        }
    }

    private void CreateUpdateTransaction(string UserId)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;
        lblMessage.Text = "";

        try
        {
            //  service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
            // sw.WriteLine("step 1 ");
        }
        // catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        foreach (GridViewRow row1 in gvCashinTransit.Rows)
        {
            System.Web.UI.WebControls.TextBox txtCashinTransit = (System.Web.UI.WebControls.TextBox)row1.FindControl("txtCashinTransit");
            string LegalEntity = row1.Cells[14].Text.Trim().Replace("LegalEntity", "").Replace("&nbsp;", "");
            string AccName = row1.Cells[13].Text.Trim().Replace("Account", "").Replace("&nbsp;", "");
            string AccountId = row1.Cells[23].Text.Trim().Replace("ssi_accountid", "").Replace("&nbsp;", "");
            string AsOfDate = row1.Cells[16].Text.Trim().Replace("ssi_asofdate", "").Replace("&nbsp;", "").Replace("As Of Date", "").Replace("False", "");
            string _800Acctflg = row1.Cells[18].Text.Trim().Replace("_800Acctflg", "").Replace("&nbsp;", "");

            string AssetClassId = row1.Cells[24].Text.Trim().Replace("sas_assetclassid ", "").Replace("&nbsp;", "");
            string subassetclassId = row1.Cells[25].Text.Trim().Replace("Ssi_subassetclassId", "").Replace("&nbsp;", "");
            string BenchmarkSubAssetClassId = row1.Cells[26].Text.Trim().Replace("Ssi_BenchmarkSubAssetClassId ", "").Replace("&nbsp;", "");
            string SectorFlg = row1.Cells[27].Text.Trim().Replace("SectorFlg ", "").Replace("&nbsp;", "");
            string Li_SectorFlg = row1.Cells[28].Text.Trim().Replace("Li_SectorFlg ", "").Replace("&nbsp;", "");
            string Lo_SectorFlg = row1.Cells[29].Text.Trim().Replace("Lo_SectorFlg ", "").Replace("&nbsp;", "");
            try
            {

                if (_800Acctflg.ToUpper() == "TRUE" && txtCashinTransit.Text != "")
                {
                    #region Create New Transaction With Li
                    // Delete Transaction log
                    GetTransactionID(LegalEntity, AccName, AsOfDate, UserId);

                    //ssi_transactionlog objTransaction = new ssi_transactionlog();
                    Entity objTransaction = new Entity("ssi_transactionlog");
                    if (AsOfDate != "")
                    {
                        //objTransaction.ssi_tradedate = new CrmDateTime();
                        //objTransaction.ssi_tradedate.Value = AsOfDate;
                        objTransaction["ssi_tradedate"] = Convert.ToDateTime(AsOfDate);


                    }

                    //objTransaction.ssi_transactioncodeid = new Lookup();//40541C70-600A-DE11-A38C-001D09665E8F
                    //objTransaction.ssi_transactioncodeid.type = EntityName.ssi_transactioncode.ToString();
                    //objTransaction.ssi_transactioncodeid.Value = new Guid(Convert.ToString("40541C70-600A-DE11-A38C-001D09665E8F"));
                    objTransaction["ssi_transactioncodeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_transactioncode", new Guid(Convert.ToString("40541C70-600A-DE11-A38C-001D09665E8F")));


                    //objTransaction.transactioncurrencyid = new Lookup();
                    //objTransaction.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                    //objTransaction.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                    objTransaction["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid(Convert.ToString("215A7268-A2E1-DD11-A826-001D09665E8F")));

                    //objTransaction.ssi_assetclassid = new Lookup();
                    //objTransaction.ssi_assetclassid.type = EntityName.ssi_assetclassallocation.ToString();
                    //objTransaction.ssi_assetclassid.Value = new Guid("9776259D-0392-4DE0-8A12-0399724ABF8D");

                    //objTransaction.ssi_securityid = new Lookup();
                    //objTransaction.ssi_securityid.type = EntityName.ssi_security.ToString();
                    //objTransaction.ssi_securityid.Value = new Guid("47C9F791-5E66-446B-BA73-E614EC3A5137");
                    objTransaction["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid(Convert.ToString("47C9F791-5E66-446B-BA73-E614EC3A5137")));


                    if (AccountId != "")
                    {
                        //objTransaction.ssi_accountid = new Lookup();
                        //objTransaction.ssi_accountid.type = EntityName.ssi_account.ToString();
                        //objTransaction.ssi_accountid.Value = new Guid(AccountId);
                        objTransaction["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(AccountId)));
                    }

                    if (txtCashinTransit.Text != "")
                    {
                        //objTransaction.ssi_value = new CrmMoney();
                        //objTransaction.ssi_value.Value = Convert.ToDecimal(txtCashinTransit.Text);
                        objTransaction["ssi_value"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(txtCashinTransit.Text));


                        //objTransaction.ssi_quantity = new CrmDecimal();
                        //objTransaction.ssi_quantity.Value = Convert.ToDecimal(txtCashinTransit.Text);
                        objTransaction["ssi_quantity"] = Convert.ToDecimal(txtCashinTransit.Text);

                    }
                    /*********************New GA Logic**************************/
                    if (AssetClassId != "")
                    {
                        //objTransaction.ssi_assetclassid = new Lookup();
                        //objTransaction.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                        //objTransaction.ssi_assetclassid.Value = new Guid(Convert.ToString(AssetClassId));
                        objTransaction["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(AssetClassId)));
                    }
                    else
                    {
                        //objTransaction.ssi_assetclassid = new Lookup();
                        //objTransaction.ssi_assetclassid.IsNull = true;
                        //objTransaction.ssi_assetclassid.IsNullSpecified = true;
                        objTransaction["ssi_assetclassid"] = null;
                    }
                    //Sub Assetclassid 
                    if (subassetclassId != "")
                    {
                        //objTransaction.ssi_subassetclassid = new Lookup();
                        //objTransaction.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                        //objTransaction.ssi_subassetclassid.Value = new Guid(Convert.ToString(subassetclassId));
                        objTransaction["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(subassetclassId)));
                    }
                    else
                    {

                        objTransaction["ssi_subassetclassid"] = null;
                    }

                    //Benchmarkid
                    if (BenchmarkSubAssetClassId != "")
                    {
                        //objTransaction.ssi_benchmarkid = new Lookup();
                        //objTransaction.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                        //objTransaction.ssi_benchmarkid.Value = new Guid(Convert.ToString(BenchmarkSubAssetClassId));
                        objTransaction["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(BenchmarkSubAssetClassId)));
                    }
                    else
                    {
                        //objTransaction.ssi_benchmarkid = new Lookup();
                        //objTransaction.ssi_benchmarkid.IsNull = true;
                        //objTransaction.ssi_benchmarkid.IsNullSpecified = true;
                        objTransaction["ssi_benchmarkid"] = null;
                    }
                    //greshamadvised (sectorflg)
                    if (SectorFlg != "")
                    {
                        //objTransaction.ssi_grehamadvised = new CrmBoolean();
                        //objTransaction.ssi_grehamadvised.Value = Convert.ToBoolean(Li_SectorFlg);
                        objTransaction["ssi_grehamadvised"] = Convert.ToBoolean(Li_SectorFlg);

                    }
                    /*********************New GA Logic End**************************/

                    service.Create(objTransaction);
                    #endregion

                    #region Transaction With Lo
                    if (AsOfDate != "")
                    {
                        DateTime NewAsOfDate = Convert.ToDateTime(AsOfDate);
                        NewAsOfDate = NewAsOfDate.AddDays(1);
                        //objTransaction.ssi_tradedate = new CrmDateTime();
                        //objTransaction.ssi_tradedate.Value = Convert.ToString(NewAsOfDate);
                        objTransaction["ssi_tradedate"] = Convert.ToDateTime(NewAsOfDate);

                    }

                    //objTransaction.ssi_transactioncodeid = new Lookup();
                    //objTransaction.ssi_transactioncodeid.type = EntityName.ssi_transactioncode.ToString();
                    //objTransaction.ssi_transactioncodeid.Value = new Guid(Convert.ToString("00CEFD97-600A-DE11-A38C-001D09665E8F"));// TransactionCodeid For lo
                    objTransaction["ssi_transactioncodeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_transactioncode", new Guid(Convert.ToString("00CEFD97-600A-DE11-A38C-001D09665E8F")));

                    //objTransaction.transactioncurrencyid = new Lookup();
                    //objTransaction.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                    //objTransaction.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                    objTransaction["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid(Convert.ToString("215A7268-A2E1-DD11-A826-001D09665E8F")));

                    //objTransaction.ssi_assetclassid = new Lookup();
                    //objTransaction.ssi_assetclassid.type = EntityName.ssi_assetclassallocation.ToString();
                    //objTransaction.ssi_assetclassid.Value = new Guid("9776259D-0392-4DE0-8A12-0399724ABF8D");

                    //objTransaction.ssi_securityid = new Lookup();
                    //objTransaction.ssi_securityid.type = EntityName.ssi_security.ToString();
                    //objTransaction.ssi_securityid.Value = new Guid("47C9F791-5E66-446B-BA73-E614EC3A5137");
                    objTransaction["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid(Convert.ToString("47C9F791-5E66-446B-BA73-E614EC3A5137")));



                    if (AccountId != "")
                    {
                        //objTransaction.ssi_accountid = new Lookup();
                        //objTransaction.ssi_accountid.type = EntityName.ssi_account.ToString();
                        //objTransaction.ssi_accountid.Value = new Guid(AccountId);
                        objTransaction["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(AccountId)));
                    }

                    if (txtCashinTransit.Text != "")
                    {
                        //objTransaction.ssi_value = new CrmMoney();
                        //objTransaction.ssi_value.Value = Convert.ToDecimal(txtCashinTransit.Text);//*(-1);
                        objTransaction["ssi_value"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(txtCashinTransit.Text));

                        //objTransaction.ssi_quantity = new CrmDecimal();
                        //objTransaction.ssi_quantity.Value = Convert.ToDecimal(txtCashinTransit.Text);//*(-1);
                        objTransaction["ssi_quantity"] = Convert.ToDecimal(txtCashinTransit.Text);

                    }

                    /*********************New GA Logic**************************/
                    if (AssetClassId != "")
                    {
                        //objTransaction.ssi_assetclassid = new Lookup();
                        //objTransaction.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                        //objTransaction.ssi_assetclassid.Value = new Guid(Convert.ToString(AssetClassId));
                        objTransaction["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(AssetClassId)));

                    }
                    else
                    {
                        //objTransaction.ssi_assetclassid = new Lookup();
                        //objTransaction.ssi_assetclassid.IsNull = true;
                        //objTransaction.ssi_assetclassid.IsNullSpecified = true;
                        objTransaction["ssi_assetclassid"] = null;
                    }
                    //Sub Assetclassid 
                    if (subassetclassId != "")
                    {
                        //objTransaction.ssi_subassetclassid = new Lookup();
                        //objTransaction.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                        //objTransaction.ssi_subassetclassid.Value = new Guid(Convert.ToString(subassetclassId));
                        objTransaction["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(subassetclassId)));
                    }
                    else
                    {
                        //objTransaction.ssi_subassetclassid = new Lookup();
                        //objTransaction.ssi_subassetclassid.IsNull = true;
                        //objTransaction.ssi_subassetclassid.IsNullSpecified = true;
                        objTransaction["ssi_subassetclassid"] = null;
                    }

                    //Benchmarkid
                    if (BenchmarkSubAssetClassId != "")
                    {
                        //objTransaction.ssi_benchmarkid = new Lookup();
                        //objTransaction.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                        //objTransaction.ssi_benchmarkid.Value = new Guid(Convert.ToString(BenchmarkSubAssetClassId));
                        objTransaction["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(BenchmarkSubAssetClassId)));
                    }
                    else
                    {
                        //objTransaction.ssi_benchmarkid = new Lookup();
                        //objTransaction.ssi_benchmarkid.IsNull = true;
                        //objTransaction.ssi_benchmarkid.IsNullSpecified = true;
                        objTransaction["ssi_benchmarkid"] = null;
                    }
                    //greshamadvised (sectorflg)
                    if (SectorFlg != "")
                    {
                        //objTransaction.ssi_grehamadvised = new CrmBoolean();
                        //objTransaction.ssi_grehamadvised.Value = Convert.ToBoolean(Lo_SectorFlg);
                        objTransaction["ssi_grehamadvised"] = Convert.ToBoolean(Convert.ToString(Lo_SectorFlg));

                    }
                    /*********************New GA Logic End**************************/

                    service.Create(objTransaction);
                    #endregion


                }

            }
            // catch (System.Web.Services.Protocols.SoapException exc)
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
                lblMessage.Text = strDescription;
                //  sw.WriteLine(strDescription);
                //LogMessage(sw, service, strDescription, 62, "GeneralError");
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                lblMessage.Text = strDescription;
                //  sw.WriteLine(strDescription);
                //LogMessage(sw, service, strDescription, 62, "GeneralError");
            }


        }


    }


    private void GetTransactionID(string LegalEntity1, string AccountID1, string AsOfDate1, string UserId)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
        //CrmService service = null;
        lblMessage.Text = "";
        IOrganizationService service = null;
        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
            // sw.WriteLine("step 1 ");
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;


        string TransactionId = string.Empty;
        //string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery = string.Empty;
        try
        {
            object LegalEntity = LegalEntity1 == "" ? "null" : "'" + LegalEntity1 + "'";
            object Account = AccountID1 == "" ? "null" : "'" + AccountID1 + "'";
            object AsOfDate = AsOfDate1 == "" ? "null" : "'" + AsOfDate1 + "'";
            greshamquery = "exec SP_S_TRANSACTION_CA_ROLLFORWARD_WITH_CASH_IN_TRANSIT  @LegalentityName = " + LegalEntity + ", @ClientAccountName = " + Account + ", @Asofdate = " + AsOfDate;
            //Response.Write(greshamquery + "<br/><br/><br/>");
            //Response.Write(LegalEntity + "<br/><br/><br/>" + Account + "<br/><br/><br/>" + AsOfDate + "<br/><br/><br/>");

            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            //Response.Write(ds_gresham.Tables.Count.ToString());
            for (int i = 0; i < ds_gresham.Tables.Count; i++)
            {
                for (int j = 0; j < ds_gresham.Tables[i].Rows.Count; j++)
                {
                    Guid TransactionID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[j]["ssi_transactionlogid"]));
                    // service.Delete(EntityName.ssi_transactionlog.ToString(), TransactionID);
                    service.Delete("ssi_transactionlog", TransactionID);
                }
            }

        }
        catch (Exception ex)
        {
            Response.Write("<br/>" + ex.Message);
        }

    }


    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        bool bGenreport = false;
        bool bSuccess = true;

        if (ddlReportOptions.SelectedValue == "1")  // run report
        {
            bSuccess = GenerateReport();
            return;
        }
        else if (ddlReportOptions.SelectedValue == "2")
        {
            bGenreport = true;
        }
        else if (ddlReportOptions.SelectedValue == "3")  // Open Mail Queue form
        {
            string csname2 = "ClientScript";
            System.Text.StringBuilder cstext2 = new System.Text.StringBuilder();
            cstext2.Append("<script type=\"text/javascript\"> ");
            cstext2.Append("window.open('BatchReport/ReportReviewForm.aspx?hhid=" + ddlHousehold.SelectedValue + "') </");
            cstext2.Append("script>");
            RegisterClientScriptBlock(csname2, cstext2.ToString());



            //string script = "<script language='javascript'> window.open('BatchReport\\MailQueue.aspx') </script>";
            //Page.ClientScript.RegisterStartupScript(this.GetType(), "cle", script, true);
            return;
        }

        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
        //CrmService service = null;
        lblMessage.Text = "";
        IOrganizationService service = null;
        string UserId = GetcurrentUser();

        CreateUpdateTransaction(UserId);
        CashinTransitUpdate(UserId);

        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //  catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        string currentCAUpdateCompleted = chkbCAUpcateCompleted.Checked ? "true" : "false";


        //update CAUpdate completedOn
        if (currentCAUpdateCompleted != Hidden1.Value)
        {

            // account objAccount = new account();
            Entity objAccount = new Entity("account");
            //primary key ssi_positionid
            //objAccount.accountid = new Key();
            //objAccount.accountid.Value = new Guid(Convert.ToString(ddlHousehold.SelectedValue));
            Guid accountid = new Guid(ddlHousehold.SelectedValue);
            objAccount["accountid"] = accountid;

            //objAccount.ssi_caupdatecompleted = new CrmBoolean();
            //objAccount.ssi_caupdatecompleted.Value = chkbCAUpcateCompleted.Checked ? true : false;
            objAccount["ssi_caupdatecompleted"] = Convert.ToBoolean(chkbCAUpcateCompleted.Checked ? true : false);
            service.Update(objAccount);

            //Response.Write("<br/>CAUpdatedCompleted updated for :" + ddlHousehold.SelectedItem.Text);

        }

        // update position
        for (int i = 0; i < gvList.Rows.Count; i++)
        {
            System.Web.UI.WebControls.TextBox txtCAUpdateValue2 = (System.Web.UI.WebControls.TextBox)gvList.Rows[i].FindControl("txtCAUpdateValue");

            CheckBox chkbxNC2 = (CheckBox)gvList.Rows[i].FindControl("chkbNC");

            //ssi_position objPosition = new ssi_position();
            Entity objPosition = new Entity("ssi_position");
            string DataSource = Convert.ToString(gvList.Rows[i].Cells[9].Text);

            string AssetClassId = Convert.ToString(gvList.Rows[i].Cells[17].Text.Trim().Replace("sas_assetclassid ", "").Replace("&nbsp;", ""));
            string subassetclassId = Convert.ToString(gvList.Rows[i].Cells[18].Text.Trim().Replace("Ssi_subassetclassId", "").Replace("&nbsp;", ""));
            string BenchmarkSubAssetClassId = Convert.ToString(gvList.Rows[i].Cells[19].Text.Trim().Replace("Ssi_BenchmarkSubAssetClassId ", "").Replace("&nbsp;", ""));
            string SectorFlg = Convert.ToString(gvList.Rows[i].Cells[20].Text.Trim().Replace("SectorFlg ", "").Replace("&nbsp;", ""));

            if (chkbxNC2.Checked)
            {
                // update datasource manual and Modified on to current user (default)

                //primary key ssi_positionid
                //objPosition.ssi_positionid = new Key();
                //objPosition.ssi_positionid.Value = new Guid(Convert.ToString(gvList.Rows[i].Cells[0].Text));
                Guid ssi_positionid = new Guid(Convert.ToString(gvList.Rows[i].Cells[0].Text));
                objPosition["ssi_positionid"] = ssi_positionid;

                //objPosition.ssi_datasource = new Picklist();
                //objPosition.ssi_datasource.Value = 6; // value of CA Update
                objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);

                //objPosition.modifiedon = new CrmDateTime();
                //objPosition.modifiedon.Value = System.DateTime.Now.ToString();//Current Date Time
                objPosition["modifiedon"] = DateTime.Now;

                successcount++;
                #region NewGA Logic
                /************************* New GA Logic **********************************/

                //SubAssetclassid 
                if (subassetclassId != "")
                {
                    //objPosition.ssi_subassetclassid = new Lookup();
                    //objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                    //objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(subassetclassId));
                    objPosition["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(subassetclassId)));


                }
                else
                {
                    //objPosition.ssi_subassetclassid = new Lookup();
                    //objPosition.ssi_subassetclassid.IsNull = true;
                    //objPosition.ssi_subassetclassid.IsNullSpecified = true;
                    objPosition["ssi_subassetclassid"] = null;
                }

                //SubAssetclassid 
                if (BenchmarkSubAssetClassId != "")
                {
                    //objPosition.ssi_benchmarkid = new Lookup();
                    //objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                    //objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(BenchmarkSubAssetClassId));
                    objPosition["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(BenchmarkSubAssetClassId)));
                }
                else
                {
                    //objPosition.ssi_benchmarkid = new Lookup();
                    //objPosition.ssi_benchmarkid.IsNull = true;
                    //objPosition.ssi_benchmarkid.IsNullSpecified = true;
                    objPosition["ssi_benchmarkid"] = null;
                }

                //Assetclassid 
                if (AssetClassId != "")
                {
                    //objPosition.ssi_assetclassid = new Lookup();
                    //objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                    //objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(AssetClassId));
                    objPosition["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(AssetClassId)));
                }
                else
                {
                    //objPosition.ssi_assetclassid = new Lookup();
                    //objPosition.ssi_assetclassid.IsNull = true;
                    //objPosition.ssi_assetclassid.IsNullSpecified = true;
                    objPosition["ssi_assetclassid"] = null;
                }

                //greshamadvised (sectorflg)
                if (SectorFlg != "")
                {
                    //objPosition.ssi_greshamadvised = new CrmBoolean();
                    //objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(SectorFlg);
                    objPosition["ssi_greshamadvised"] = Convert.ToBoolean(Convert.ToString(SectorFlg));

                }

                /************************* End New GA Logic *****************************/
                #endregion
                service.Update(objPosition);
            }
            else  // unchecked
            {
                if (txtCAUpdateValue2.Text.Trim() != "" && (Convert.ToString(gvList.Rows[i].Cells[12].Text) != "" && Convert.ToString(gvList.Rows[i].Cells[12].Text) != "&nbsp;") && DataSource.ToLower() != "Manual".ToLower())
                {
                    //update market value for position
                    // update source manual

                    //objPosition.ssi_positionid = new Key();
                    //objPosition.ssi_positionid.Value = new Guid(Convert.ToString(gvList.Rows[i].Cells[0].Text));
                    objPosition["ssi_positionid"] = new Guid(Convert.ToString(gvList.Rows[i].Cells[0].Text));

                    //objPosition.ssi_datasource = new Picklist();
                    //objPosition.ssi_datasource.Value = 6; // value of CA Update
                    objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);

                    //objPosition.ssi_marketvalue = new CrmMoney();
                    //objPosition.ssi_marketvalue.Value = Convert.ToDecimal(txtCAUpdateValue2.Text.Trim());
                    objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(txtCAUpdateValue2.Text.Trim()));


                    #region NewGA Logic
                    /************************* New GA Logic **********************************/

                    //SubAssetclassid 
                    if (subassetclassId != "")
                    {
                        //objPosition.ssi_subassetclassid = new Lookup();
                        //objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                        //objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(subassetclassId));
                        objPosition["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(subassetclassId)));


                    }
                    else
                    {
                        //objPosition.ssi_subassetclassid = new Lookup();
                        //objPosition.ssi_subassetclassid.IsNull = true;
                        //objPosition.ssi_subassetclassid.IsNullSpecified = true;
                        objPosition["ssi_subassetclassid"] = null;
                    }

                    //SubAssetclassid 
                    if (BenchmarkSubAssetClassId != "")
                    {
                        //objPosition.ssi_benchmarkid = new Lookup();
                        //objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                        //objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(BenchmarkSubAssetClassId));
                        objPosition["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(BenchmarkSubAssetClassId)));
                    }
                    else
                    {
                        //objPosition.ssi_benchmarkid = new Lookup();
                        //objPosition.ssi_benchmarkid.IsNull = true;
                        //objPosition.ssi_benchmarkid.IsNullSpecified = true;
                        objPosition["ssi_benchmarkid"] = null;
                    }

                    //Assetclassid 
                    if (AssetClassId != "")
                    {
                        //objPosition.ssi_assetclassid = new Lookup();
                        //objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                        //objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(AssetClassId));
                        objPosition["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(AssetClassId)));
                    }
                    else
                    {
                        //objPosition.ssi_assetclassid = new Lookup();
                        //objPosition.ssi_assetclassid.IsNull = true;
                        //objPosition.ssi_assetclassid.IsNullSpecified = true;
                        objPosition["ssi_assetclassid"] = null;
                    }

                    //greshamadvised (sectorflg)
                    if (SectorFlg != "")
                    {
                        //objPosition.ssi_greshamadvised = new CrmBoolean();
                        //objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(SectorFlg);
                        objPosition["ssi_greshamadvised"] = Convert.ToBoolean(Convert.ToString(SectorFlg));

                    }

                    /************************* End New GA Logic *****************************/
                    #endregion
                    service.Update(objPosition);
                    successcount++;

                    //Response.Write("<br/>Market value and datasource updated");
                }
                else if (txtCAUpdateValue2.Text.Trim() != "" && DataSource.ToLower() == "Manual".ToLower())
                {
                    //objPosition.ssi_positionid = new Key();
                    //objPosition.ssi_positionid.Value = new Guid(Convert.ToString(gvList.Rows[i].Cells[0].Text));
                    objPosition["ssi_positionid"] = new Guid(Convert.ToString(gvList.Rows[i].Cells[0].Text));

                    //objPosition.ssi_datasource = new Picklist();
                    //objPosition.ssi_datasource.Value = 6; // value of CA Update
                    objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);

                    //objPosition.ssi_marketvalue = new CrmMoney();
                    //objPosition.ssi_marketvalue.Value = Convert.ToDecimal(txtCAUpdateValue2.Text.Trim());
                    objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(txtCAUpdateValue2.Text.Trim()));

                    #region NewGA Logic
                    /************************* New GA Logic **********************************/

                    //SubAssetclassid 
                    if (subassetclassId != "")
                    {
                        //objPosition.ssi_subassetclassid = new Lookup();
                        //objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                        //objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(subassetclassId));
                        objPosition["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(subassetclassId)));


                    }
                    else
                    {
                        //objPosition.ssi_subassetclassid = new Lookup();
                        //objPosition.ssi_subassetclassid.IsNull = true;
                        //objPosition.ssi_subassetclassid.IsNullSpecified = true;
                        objPosition["ssi_subassetclassid"] = null;
                    }

                    //SubAssetclassid 
                    if (BenchmarkSubAssetClassId != "")
                    {
                        //objPosition.ssi_benchmarkid = new Lookup();
                        //objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                        //objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(BenchmarkSubAssetClassId));
                        objPosition["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(BenchmarkSubAssetClassId)));
                    }
                    else
                    {
                        //objPosition.ssi_benchmarkid = new Lookup();
                        //objPosition.ssi_benchmarkid.IsNull = true;
                        //objPosition.ssi_benchmarkid.IsNullSpecified = true;
                        objPosition["ssi_benchmarkid"] = null;
                    }

                    //Assetclassid 
                    if (AssetClassId != "")
                    {
                        //objPosition.ssi_assetclassid = new Lookup();
                        //objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                        //objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(AssetClassId));
                        objPosition["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(AssetClassId)));
                    }
                    else
                    {
                        //objPosition.ssi_assetclassid = new Lookup();
                        //objPosition.ssi_assetclassid.IsNull = true;
                        //objPosition.ssi_assetclassid.IsNullSpecified = true;
                        objPosition["ssi_assetclassid"] = null;
                    }

                    //greshamadvised (sectorflg)
                    if (SectorFlg != "")
                    {
                        //objPosition.ssi_greshamadvised = new CrmBoolean();
                        //objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(SectorFlg);
                        objPosition["ssi_greshamadvised"] = Convert.ToBoolean(Convert.ToString(SectorFlg));

                    }

                    /************************* End New GA Logic *****************************/
                    #endregion
                    service.Update(objPosition);
                    successcount++;
                }
                else // if unchecked and no value is entered
                {
                    // completely ignore record
                }
            }
        }

        if (successcount > 0)
            lblMessage.Text = lblMessage.Text + "<br/>" + successcount.ToString() + " Records updated successfully";// for CLIENT SPECIFIC UPDATES";
        else
            lblMessage.Text = lblMessage.Text + "<br/> No Records updated for ";//CLIENT SPECIFIC UPDATES";

        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");


        if (bGenreport)  //Generate Report 
            GenerateReport();

    }

    private bool GenerateReport()
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        bool isSuccess = true;
        try
        {
            lblMessage.Text = "";
            string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
            //string crmServerURL = "http://server:5555/";

            string orgName = "GreshamPartners";
            string currentuser = null;
            //string orgName = "Webdev";
            // CrmService service = null;
            IOrganizationService service = null;
            Boolean checkrunreport = false;
            String DestinationPath = string.Empty;
            string ConsolidatePdfFileName = string.Empty;

            DataTable dtBatch = null;

            //string[] distColName = { "Ssi_ContactIdName" };

            DateTime dt = DateTime.Now;

            string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();

            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

            //string strUserName = HttpContext.Current.User.Identity.Name.ToString();
            //Changed Windows to - ADFS Claims Login 8_9_2019
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            string strUserName = claimsIdentity.Name;

            //Response.Write(strUserName);

            strUserName = strUserName.Substring(strUserName.IndexOf("\\") + 1);

            ViewState["ParentFolder"] = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;
            string ReportOpFolder = string.Empty;
            //string ReportOpFolder = "\\\\Fs01\\_ops_C_I_R_group\\Quarterly_Reports\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
                //ReportOpFolder = @"C:\Reports\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
            {
                ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
                //ReportOpFolder = "\\\\Fs01\\_ops_C_I_R_group\\BATCH REPORTS\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }

            string ContactFolderName = string.Empty;
            FileInfo loCoversheetCheck;
            String ReportOpFolder1 = String.Empty;

            /*****Start :  Array declaration for PDF merge **************/
            PDFMerge PDF = new PDFMerge();
            int sourcefilecount = 0;//= dtBatch.Rows.Count + 1;
            string[] SourceFileArray;
            /*****End   :  Array declaration for PDF merge **************/

            ConsolidatePdfFileName = "ConsolidatedPDF" + "_" + strYear + strMonth + strDay + "_" + ".pdf";

            checkrunreport = true;
            String HouseholdIdListTxt = Convert.ToString(ddlHousehold.SelectedValue);


            DataTable dtBatchList = GetBatchList(ddlHousehold.SelectedValue, "");

            if (dtBatchList.Rows.Count < 1)
            {
                lblMessage.Text = "Report can not be generated, Batch not found for this Household";

                return false;
            }
            string strBatchId = Convert.ToString(dtBatchList.Rows[0]["Ssi_batchId"]);
            dtBatch = GetDataTable(strBatchId);

            if (dtBatch.Rows.Count < 1)
            {
                lblMessage.Text = "Report can not be generated, Batch or householdparameter not found for this Household";

                return false;
            }

            //Pdf File Name
            String PdfFileName = Convert.ToString(dtBatchList.Rows[0]["PdfFileName"]).Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

            sourcefilecount = dtBatch.Rows.Count + 2;
            SourceFileArray = new string[sourcefilecount];

            for (int i = 0; i < dtBatch.Rows.Count; i++)
            {
                ContactFolderName = Convert.ToString(dtBatchList.Rows[0]["FolderNameTxt"]).Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                //ContactFolderName = Convert.ToString(dtBatch.Rows[i]["Ssi_ContactIdName"]).Replace(",", "");
                bool isExist = System.IO.Directory.Exists(ReportOpFolder + "\\" + ContactFolderName);

                if (!isExist)
                {
                    //Response.Write("Folder: " + ReportOpFolder + "\\" + ContactFolderName);
                    System.IO.Directory.CreateDirectory(ReportOpFolder + "\\" + ContactFolderName);
                }

                ViewState["AsOfDate"] = Convert.ToString(dtBatch.Rows[i]["Ssi_EndAsOfDate2"]);
                // ViewState["PdfFileName"] = HHName = Convert.ToString(dtBatch.Rows[i]["PdfFileName"]);

                String fsAllocationGroup = Convert.ToString(dtBatch.Rows[i]["Ssi_AllocationGroup"]).Replace("'", "''");
                String fsHouseholdName = Convert.ToString(dtBatch.Rows[i]["Ssi_HouseholdIdName"]).Replace("'", "''");
                String fsAsofDate = Convert.ToString(dtBatch.Rows[i]["Ssi_EndAsOfDate2"]);
                String fsSPriorDate = Convert.ToString(dtBatch.Rows[i]["Ssi_StartPriorDate1"]);
                String fsLookthrogh = Convert.ToString(dtBatch.Rows[i]["Ssi_ConsolidateDetailLevel"]);
                String fsContactFullname = Convert.ToString(dtBatch.Rows[i]["Ssi_ContactIdName"]);
                String fsVersion = Convert.ToString(dtBatch.Rows[i]["Ssi_UnderlyingManagerDetail"]);
                String fsSummaryFlag = Convert.ToString(dtBatch.Rows[i]["Ssi_SummaryDetail"]);
                String fsAllignment = Convert.ToString(dtBatch.Rows[i]["Ssi_Alignment"]);
                String fsDisplayContactName = Convert.ToString(dtBatch.Rows[i]["ContactName"]);
                String fsContactId = Convert.ToString(dtBatch.Rows[i]["ssi_ContactID"]);
                String fsKeyContactID = Convert.ToString(dtBatch.Rows[i]["ssi_keycontactId"]);
                String fsHousholdReportTitle = Convert.ToString(dtBatch.Rows[i]["ssi_householdreporttitle"]);
                String fsGreshReportIdName = Convert.ToString(dtBatch.Rows[i]["ssi_GreshamReportIdName"]);
                String fsGAorTIAflag = Convert.ToString(dtBatch.Rows[i]["ssi_gaortia"]);
                String lsFinalTitleAfterChange = String.Empty;
                String fsDiscretionaryFlg = Convert.ToString(dtBatch.Rows[i]["Discretionary Flag"]);
                String fsReportRollupGroupIdName = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportRollupGroupIdName"]).Replace("'", "''");
                String fsHHreportparametersId = Convert.ToString(dtBatch.Rows[i]["Ssi_hhreportparametersId"]);
                fsReportingName = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportingName"]);

                if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["HouseHoldReportTitle"])))
                    lsFinalTitleAfterChange = Convert.ToString(dtBatch.Rows[i]["HouseHoldReportTitle"]);

                if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["AllocationGroupReportTitle"])))
                    lsFinalTitleAfterChange = Convert.ToString(dtBatch.Rows[i]["AllocationGroupReportTitle"]);

                String fsFooterTxt = String.Empty;
                if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["GreshamFooterTxt"])))
                    fsFooterTxt = Convert.ToString(dtBatch.Rows[i]["GreshamFooterTxt"]);

                /*Change added on 31st OCT 2010*/
                String fsReportGroupflag = "null";
                if (Convert.ToString(dtBatch.Rows[i]["ssi_report"]) == "")
                    fsReportGroupflag = "null";
                else
                    fsReportGroupflag = Convert.ToString(dtBatch.Rows[i]["ssi_report"]);
                //Convert.ToString(dtBatch.Rows[i]["ssi_report"]).Replace(",", "");
                String fsReportgroupflag2 = "null";
                if (Convert.ToString(dtBatch.Rows[i]["ssi_report2"]) == "")
                    fsReportgroupflag2 = "null";
                else
                    fsReportgroupflag2 = Convert.ToString(dtBatch.Rows[i]["ssi_report2"]);

                /* END OF CHANGE*/

                string strGUID = Guid.NewGuid().ToString();
                strGUID = strGUID.Substring(0, 5);
                //String lsExcleSavePath = ReportOpFolder + "\\" + ContactFolderName + "\\" + fsHouseholdName.Replace(",", "") + "_" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + strGUID + ".xls";
                String lsExcleSavePath = ReportOpFolder + "\\" + ContactFolderName + "\\" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + lsFinalTitleAfterChange.Replace(",", "").Replace("/", "").Replace("\\", "") + "_" + Convert.ToDateTime(fsAsofDate).ToString("yyyyMMdd") + "_" + strGUID + ".xls";
                String lsCoversheet = ReportOpFolder + "\\" + ContactFolderName + "\\Coversheet.xls";
                //String fsHouseHoldReportTitle = "";

                //Page number logic 
                if (i == 0)
                {
                    dtBatch.Columns.Add("numPageNo", typeof(System.Int32));
                    dtBatch.Rows[i]["numPageNo"] = "1";
                }
                #region added sasmit(7_14_2017)
                bool bContinueBatch = true;
                /** Attach Template PDF ---Static pdf logic  ***/
                string strTemplateFilePath = Convert.ToString(dtBatch.Rows[i]["ssi_TemplateFilePath"]);
                if (strTemplateFilePath != "")
                {
                    string strExtension = Path.GetExtension(strTemplateFilePath);


                    #region Fetch File from Sharepoint

                    if (strTemplateFilePath.Contains("https://greshampartners.sharepoint.com") || strTemplateFilePath.Contains("http://greshampartners.sharepoint.com"))
                    {

                        string FileName = Path.GetFileName(strTemplateFilePath);
                        FileName = FileName.Replace("%20", " ");
                        // string FileName2 = HttpUtility.HtmlEncode(FileName).ToString();
                        string SharepointPath = strTemplateFilePath;
                        SharepointPath = SharepointPath.Replace("//", "/");
                        SharepointPath = SharepointPath.Replace("https:/greshampartners.sharepoint.com/clientserv/", "");
                        SharepointPath = SharepointPath.Replace("http:/greshampartners.sharepoint.com/clientserv/", "");
                        SharepointPath = SharepointPath.Replace("%20", " ");
                        SharepointPath = SharepointPath.Replace(FileName, "");

                        string LocalPath = ReportOpFolder + "\\" + ContactFolderName + "\\";

                        strTemplateFilePath = sharepointFile(FileName, SharepointPath, LocalPath);
                    }
                    #endregion
                    if (strExtension.ToString().ToLower() == ".doc" || strExtension.ToString().ToLower() == ".docx")
                    {
                        strTemplateFilePath = ConvertDocument(strTemplateFilePath, lsExcleSavePath);
                        strTemplateFilePath = strTemplateFilePath.Replace(".xls", ".pdf");
                    }
                    if (strExtension.ToString().ToLower() == ".xls" || strExtension.ToString().ToLower() == ".xlsx")
                    {
                        strTemplateFilePath = ConvertSpreadsheet(strTemplateFilePath, lsExcleSavePath);
                        strTemplateFilePath = strTemplateFilePath.Replace(".xls", ".pdf");
                    }

                    //FOR -- TESTING 
                    if (Request.Url.AbsoluteUri.Contains("localhost"))
                        strTemplateFilePath = @"C:\Reports\Commentaries.pdf";

                    if (Convert.ToString(Session["CurPageInBatch"]) == "")
                        Session["CurPageInBatch"] = "0";

                    lsExcleSavePath = strTemplateFilePath.Replace(".pdf", ".xls");
                    int numofPage = objCombinedReports.GetPageCountFromPDF(strTemplateFilePath);
                    int CurPage = Convert.ToInt32(Convert.ToString(Session["CurPageInBatch"])) + 1;
                    if (numofPage > 0)
                    {
                        numofPage--;
                        dtBatch.Rows[i]["numPageNo"] = CurPage;
                        Session["CurPageInBatch"] = numofPage + CurPage;
                        bContinueBatch = false;
                    }



                    else
                        dtBatch.Rows[i]["numPageNo"] = 0;

                }

                bool CombinedFileName = false;

                /** if record is template then it will not generate report -- only static pdf will attach **/
                /** Generate report on excel and pdf **/
                if (bContinueBatch)
                {
                    if (i != 0)
                    {
                        if (Session["CurPageInBatch"] != null)
                        {
                            int CurPage = Convert.ToInt32(Convert.ToString(Session["CurPageInBatch"])) + 1;
                            dtBatch.Rows[i]["numPageNo"] = CurPage;
                        }
                    }

                    // Generate report on excel and pdf

                    // bool CombinedFileName = false;
                    if (fsGreshReportIdName != "Asset Distribution" && fsGreshReportIdName != "Asset Distribution Comparison")
                    {
                        CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, fsGAorTIAflag, fsReportRollupGroupIdName, fsHHreportparametersId);
                    }
                    else
                    {
                        SetValuesToVariable(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                        // generatesExcelsheets(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                        generatePDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                        CombinedFileName = true;
                    }

                    loCoversheetCheck = new FileInfo(lsCoversheet);
                    if (!loCoversheetCheck.Exists)
                    {
                        generateCoversheetPDF(fsAsofDate, lsCoversheet, fsAllocationGroup, fsHouseholdName, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsGAorTIAflag, fsDiscretionaryFlg);
                        //generatesCoverExcel(fsAsofDate, fsHouseholdName, fsAllocationGroup, lsCoversheet, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange);
                    }

                    /* Array fill with the PATH + Fullname of PDF*/
                }
                else
                {
                    CombinedFileName = true;
                }
                #endregion added sasmit(7_14_2017)
                if (i == 0)
                {
                    SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                    SourceFileArray[i + 1] = (Server.MapPath("") + @"\ExcelTemplate\Blank.pdf");
                    if (CombinedFileName == true)
                        SourceFileArray[i + 2] = lsExcleSavePath.Replace(".xls", ".pdf");
                }
                else
                {
                    if (CombinedFileName == true)
                        SourceFileArray[i + 2] = lsExcleSavePath.Replace(".xls", ".pdf");

                }

                /* Array fill with the PATH + Fullname of PDF*/


            }

            // Consolidate File Logic ORIGINAL
            //File.Copy(ReportOpFolder + " " + TempName + "\\" + ContactFolderName + "\\Coversheet.pdf", ReportOpFolder + " " + TempName + "\\" + ContactFolderName + "\\" + "ConsolidatedReport.pdf");
            //String DestinationPath = ReportOpFolder + " " + TempName + "\\" + ContactFolderName + "\\" + "ConsolidatedReport.pdf";

            // Consolidate File Logic NEW
            DateTime dtAsOfDate = Convert.ToDateTime(ViewState["AsOfDate"]);

            strYear = dtAsOfDate.Year.ToString().Length < 2 ? "0" + dtAsOfDate.Year.ToString() : dtAsOfDate.Year.ToString();
            strMonth = dtAsOfDate.Month.ToString().Length < 2 ? "0" + dtAsOfDate.Month.ToString() : dtAsOfDate.Month.ToString();
            strDay = dtAsOfDate.Day.ToString().Length < 2 ? "0" + dtAsOfDate.Day.ToString() : dtAsOfDate.Day.ToString();

            //string ConsolidatePdfFileName = ContactFolderName + "_" + strYear + strMonth + strDay + ".pdf";
            ConsolidatePdfFileName = PdfFileName + "_" + strYear + "-" + strMonth + strDay + ".pdf";


            if (!System.IO.File.Exists(ReportOpFolder + "\\" + ConsolidatePdfFileName))
                System.IO.File.Copy(ReportOpFolder + "\\" + ContactFolderName + "\\Coversheet.pdf", ReportOpFolder + "\\" + ConsolidatePdfFileName);

            DestinationPath = ReportOpFolder + "\\" + ConsolidatePdfFileName;



            if (ContactFolderName.Contains("MTGBK")) //generate without coversheet
            {
                string[] target = new string[sourcefilecount - 2];
                Array.Copy(SourceFileArray, 2, target, 0, sourcefilecount - 2);
                PDF.MergeFiles(DestinationPath, target);
            }
            else //generate with coversheet
            {
                PDF.MergeFiles(DestinationPath, SourceFileArray);
                string DestinationPath1 = objCombinedReports.addPageIndex(DestinationPath, dtBatch,"");
                System.IO.File.Copy(DestinationPath1, DestinationPath, true);
            }

            //Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
            Session.Remove("CurPageInBatch");


            ////////////////////////////////////

            if (1 == 1) // Output report 
            {
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName);

                System.IO.File.Copy(DestinationPath, strDirectory, true);
                Directory.Delete(ReportOpFolder, true);
                string lsFileNamforFinal;
                lsFileNamforFinal = "./ExcelTemplate/" + ConsolidatePdfFileName;
                string newWindow = string.Empty;
                try
                {
                    ////loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                    Response.Write("<script>");
                    lsFileNamforFinal = "./ExcelTemplate/" + ConsolidatePdfFileName;
                    Response.Write("window.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow')");
                    Response.Write("</script>");
                }
                catch (Exception exc)
                {
                    Response.Write(exc.Message);
                }

                //// Response.Clear();
                //Response.Buffer = false; //transmitfile self buffers
                ////Response.Clear();
                //Response.ClearContent();
                //Response.ClearHeaders();
                //Response.ContentType = "application/pdf";
                //Response.AddHeader("Content-Disposition", "inline;filename=" + ConsolidatePdfFileName);
                //Response.WriteFile(lsFileNamforFinal); //transmitfile keeps entire file from loading into memory
                //Response.End();

                //Response.Clear();
                //Response.ContentType = "application/pdf";
                //Response.AddHeader("Content-Disposition", "attachement;filename=" + ConsolidatePdfFileName);
                //Context.Response.Buffer = false;
                //FileStream file = null;
                //byte[] mybuff = new byte[1024];
                //long count;

                //string filePATH = Server.MapPath(lsFileNamforFinal);
                //file = File.OpenRead(filePATH);

                //while ((count = file.Read(mybuff, 0, mybuff.Length)) > 0)
                //{
                //    if (Context.Response.IsClientConnected)
                //    {
                //        Context.Response.OutputStream.Write(mybuff, 0, mybuff.Length);
                //        Context.Response.Flush();
                //    }
                //}



            }
            ////////////////////////////////////

            if (checkrunreport)
                lblMessage.Text = "Reports generated successfully";
            else
                lblMessage.Text = "Please Select a batch to run report.";
        }
        catch (Exception ex)
        {
            isSuccess = false;
            lblMessage.Text = "Error Generating Report " + ex.ToString();
        }

        return isSuccess;
    }
    private string ConvertDocument(string strSourcePath, string strDestPath)
    {
        try
        {

            ComponentInfo.SetLicense("D7OT-O3KE-PMVU-IXWZ");
            //ComponentInfo.FreeLimitReached += (sender1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            DocumentModel document = DocumentModel.Load(strSourcePath);

            document.Save(strDestPath.Replace(".xls", ".pdf"));

            return strDestPath.Replace(".pdf", ".xls");


        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            return "";
        }
    }

    private string ConvertSpreadsheet(string strSourcePath, string strDestPath)
    {
        try
        {

            SpreadsheetInfo.SetLicense("E43Y-7VYO-CTN8-X97J");
            // ComponentInfo.FreeLimitReached += (sender1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            ExcelFile document = ExcelFile.Load(strSourcePath);

            document.Save(strDestPath.Replace(".xls", ".pdf"));

            return strDestPath.Replace(".pdf", ".xls");


        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            return "";
        }
    }
    public string sharepointFile(string FileName, string path, string finalPath)
    {
        string Value = null;


        string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();

        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);


        Web site = context.Web;

        // Folder subFoldercol = site.GetFolderByServerRelativeUrl("Documents" + "/"+"_Test Files");
        // Folder subFoldercol = site.GetFolderByServerRelativeUrl(path.ToLower().Replace("clientserv/", ""));
        Folder subFoldercol = site.GetFolderByServerRelativeUrl(path);
        // Microsoft.SharePoint.Client.File subfile = site.GetFileByServerRelativeUrl("Anziano" + "/" + Path);
        ListCollection collList = site.Lists;

        //  FolderCollection fcolection = subFoldercol.Folders;
        Microsoft.SharePoint.Client.FileCollection fcolection = subFoldercol.Files;
        context.Load(fcolection);
        context.Load(collList);
        context.ExecuteQuery();
        foreach (Microsoft.SharePoint.Client.File f in fcolection)
        {

            string FileNAME = f.Name.ToString();
            if (FileName == FileNAME)
            {
                FileCopy(f, finalPath);
                Value = finalPath + "\\" + FileName;
                break;
            }
            else
            {
                Value = null;
            }
        }
        return Value;
    }
    public void FileCopy(Microsoft.SharePoint.Client.File files1, string finalPath)
    {
        // -- Get fIle and copy to Destination
        Stream filestrem = getFile(files1);
        string fileName = System.IO.Path.GetFileName(files1.Name);
        // string filepath = System.IO.Path.Combine(Test, fileName);
        string filepath = System.IO.Path.Combine(finalPath, fileName);
        // FileStream fileStream = System.IO.File.Create(filepath, (int)filestrem.Length); // Test Local PAth
        FileStream fileStream = System.IO.File.Create(filepath, (int)filestrem.Length); // Original PAth
        // Initialize the bytes array with the stream length and then fill it with data 
        byte[] bytesInStream = new byte[filestrem.Length];
        filestrem.Read(bytesInStream, 0, bytesInStream.Length);
        // Use write method to write to the file specified above 
        fileStream.Write(bytesInStream, 0, bytesInStream.Length);

        fileStream.Close();
    }
    public Stream getFile(Microsoft.SharePoint.Client.File files1)
    {
        context.Load(files1);
        ClientResult<Stream> stream = files1.OpenBinaryStream();
        context.ExecuteQuery();
        return this.ReadFully(stream.Value);
    }
    private Stream ReadFully(Stream input)
    {
        byte[] buffer = new byte[16 * 1024];
        using (MemoryStream ms = new MemoryStream())
        {
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                ms.Write(buffer, 0, read);
            }
            return new MemoryStream(ms.ToArray()); ;
        }
    }
    public bool generateCombinedPDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate,
        String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment,
        String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String ReportName, String GAorTIAflag, String ReportRollupGroupIdName, String fsHHreportparametersId)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();

        objCombinedReports.HouseHoldValue = "";
        objCombinedReports.HouseHoldText = fsHouseholdName;
        objCombinedReports.AllocationGroupValue = "";
        objCombinedReports.AllocationGroupText = fsAllocationGroup;
        objCombinedReports.AsOfDate = fsAsofDate;
        objCombinedReports.lsFamiliesName = fsHouseholdName;
        objCombinedReports.lsDateName = "";
        objCombinedReports.GreshamAdvisedFlag = GAorTIAflag;
        objCombinedReports.ReportRollupGroupIdName = ReportRollupGroupIdName;

        if (fsReportingName != "")
            objCombinedReports.ReportingName = fsReportingName;


        if (ReportName == "Client Goals" || ReportName == "Absolute Returns" || ReportName == "Capital Protection")
        {
            string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";


            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            String HHRPIDListTxt = Convert.ToString(fsHHreportparametersId);
            string greshamquery = "[SP_S_HH_PARAMETER_ASSETCLASS] @HHParameterListTxt='" + HHRPIDListTxt + "'";

            SqlDataAdapter dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            DataSet ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);

            if (ds_gresham.Tables[0].Rows.Count > 0)
            {
                string _assetclass = "";
                for (int i = 0; i < ds_gresham.Tables[0].Rows.Count; i++)
                {
                    _assetclass = _assetclass + "," + ds_gresham.Tables[0].Rows[i]["sas_name"].ToString();
                }

                _assetclass = _assetclass.Substring(1, _assetclass.Length - 1);
                objCombinedReports.AssetClassCSV = _assetclass;
            }
        }


        string filepdfname = objCombinedReports.MergeReports(fsFinalLocation, ReportName);

        if (filepdfname == "")
        {
            return false;
        }
        else
            return true;

    }

    public void SetValuesToVariable(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFinalReportTitle, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        String lsfamilyName = fsHouseholdName;
        int liCommaCounter = lsfamilyName.IndexOf(",");
        int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
        if (liCommaCounter > 0 && liSpaceCounter > 0)
            lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
        else
            lsfamilyName = lsfamilyName;

        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsfamilyName = fsAllocationGroup;
        }
        if (!String.IsNullOrEmpty(lsFinalReportTitle))
            lsfamilyName = lsFinalReportTitle;

        //Set for Pdf
        if (fsAllignment != "Horizontal")
            lsDistributionName = "Asset Distribution Comparison";
        else
            lsDistributionName = "Asset Distribution";

        lsFamiliesName = lsfamilyName;
        lsDateName = Convert.ToDateTime(fsAsofDate).ToString("MMMM dd, yyyy") + "";

        if (fsGAorTIAflag == "GA")
        {
            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                lsGAorTIAHeader = "GRESHAM ADVISED ASSETS - DISCRETIONARY";
            else
                lsGAorTIAHeader = "GRESHAM ADVISED ASSETS";
        }
        else
        {
            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                lsGAorTIAHeader = "TOTAL INVESTMENT ASSETS - DISCRETIONARY";
            else
                lsGAorTIAHeader = "TOTAL INVESTMENT ASSETS";
        }
    }

    public void generatePDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        liPageSize = 29;
        DataSet lodataset; DB clsDB = new DB();
        lodataset = null;

        String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, fsGAorTIAflag, fsDiscretionaryFlg);
        // Response.Write(lsSQL);
        lodataset = clsDB.getDataSet(lsSQL);
        DataSet loInsertblankRow = lodataset.Copy();
        lodataset.Tables[0].Clear();
        lodataset.Clear();
        lodataset = null;
        lodataset = loInsertblankRow.Clone();
        int liBlankCounter = 1;
        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Rows.Count; liBlankRow++)
        {
            if (liBlankRow != 0 && loInsertblankRow.Tables[0].Rows[liBlankRow]["_Ssi_BoldFlg"].ToString().ToUpper() == "TRUE" || loInsertblankRow.Tables[0].Rows[liBlankRow]["_Ssi_SuperBoldFlg"].ToString().ToUpper() == "TRUE")
            {
                //if (!String.IsNullOrEmpty(fsSPriorDate) && loInsertblankRow.Tables[0].Rows.Count - 1 != liBlankRow)
                if (loInsertblankRow.Tables[0].Rows.Count - 1 != liBlankRow)
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[0] = "test";
                    newCustomersRow[1] = "test";
                    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    liBlankCounter = liBlankCounter + 1;
                }
                else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liBlankRow][0]) == "NET WORTH")
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[0] = "test";
                    newCustomersRow[1] = "test";
                    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    liBlankCounter = liBlankCounter + 1;
                }
                else if (fsAllignment != "Horizontal")
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[0] = "test";
                    newCustomersRow[1] = "test";
                    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    liBlankCounter = liBlankCounter + 1;
                }
            }
            lodataset.Tables[0].ImportRow(loInsertblankRow.Tables[0].Rows[liBlankRow]);
        }
        lodataset.AcceptChanges();
        DataSet loInsertdataset = lodataset.Copy();
        for (int liNewdataset = lodataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        {
            if (lodataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || lodataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
            }
        }
        //    loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[1]);
        loInsertdataset.AcceptChanges();

        //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 10);
        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
        String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();


        lsTotalNumberofColumns = loInsertdataset.Tables[0].Columns.Count + "";
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, loInsertdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        setTableProperty(loTable);
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        int liTotalPage = (loInsertdataset.Tables[0].Rows.Count / liPageSize);
        int liCurrentPage = 0;
        if (loInsertdataset.Tables[0].Rows.Count % liPageSize != 0)
        {
            liTotalPage = liTotalPage + 1;
        }
        else
        {
            liPageSize = 28;
            liTotalPage = liTotalPage + 1;
        }

        //check the length of the column name to set the pagesize.
        for (int j = 0; j < loInsertdataset.Tables[0].Columns.Count; j++)
        {
            if (loInsertdataset.Tables[0].Columns[j].ColumnName.Length > 30)
            {
                liPageSize = 28;
            }
        }

        for (int liRowCount = 0; liRowCount < loInsertdataset.Tables[0].Rows.Count; liRowCount++)
        {
            if (liRowCount % liPageSize == 0)
            {
                document.Add(loTable);

                if (liRowCount != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    // document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty));//Commented -- FooterLogic
                    document.NewPage();
                    objCombinedReports.SetTotalPageCount("Asset Distribution");
                }


                setHeader(document, loInsertdataset);
                loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, loInsertdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                setTableProperty(loTable);
            }

            int colsize = loInsertdataset.Tables[0].Columns.Count;
            for (int liColumnCount = 0; liColumnCount < colsize; liColumnCount++)
            {
                iTextSharp.text.Chunk lochunk = new Chunk();
                String lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                try
                {
                    if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1 && fsAllignment == "Horizontal")
                    {
                        lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
                    }
                    else
                    {
                        lsFormatedString = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));

                    }
                }
                catch
                {

                }

                //changed on 02/25/2011
                //lochunk = new Chunk(lsFormatedString, Font8Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                lochunk = new Chunk(lsFormatedString, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                loCell = new iTextSharp.text.Cell();
                loCell.Border = 0;
                loCell.NoWrap = true;
                // loCell.VerticalAlignment=0;
                loCell.VerticalAlignment = 5;

                setGreyBorder(lodataset, loCell, liRowCount);
                loCell.Leading = 6f;//6
                loCell.UseBorderPadding = true;

                //  if (lodataset.Tables[0].Rows[liRowCount]["_Ssi_TabFlg"].ToString() == "True" && lodataset.Tables[0].Rows[liRowCount]["_Ssi_UnderlineFlg"].ToString() != "True")


                if (liColumnCount != 0)
                {
                    loCell.HorizontalAlignment = 2;
                }


                /*=========START WITH BOLD AND SUPERBOLD FLAG========*/
                if (checkTrue(lodataset, liRowCount, "_Ssi_BoldFlg") || checkTrue(lodataset, liRowCount, "_Ssi_SuperBoldFlg"))
                {
                    lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                    try
                    {
                        if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1 && fsAllignment == "Horizontal")
                        {
                            lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
                        }
                        else
                        {
                            lsFormatedString = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));

                        }
                    }
                    catch
                    {

                    }

                    //changed on 02/25/2011
                    //lochunk = new Chunk(lsFormatedString, Font9Bold());
                    lochunk = new Chunk(lsFormatedString, Font8Bold());

                    if (!lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE"))
                    {
                        //changed on 02/25/2011
                        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                        lochunk = new Chunk(lsFormatedString, Font8Bold());
                        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                        if (lsFormatedString.Length > 25)
                        {
                            if (checkTrue(lodataset, liRowCount, "_Ssi_BoldFlg"))
                            {
                                //decrease columncount by 1 to adjust the Colspan. eg: NON-INVESTMENT ASSETS/LOOK-THROUGHS
                                loCell.Colspan = 2;
                                colsize = colsize - 1;
                            }
                        }
                        setBottomWidthWhite(loCell);

                    } /*=========IF END OF BOLD AND SUPERBOLD FLAG========*/
                    else
                    {
                        if (lodataset.Tables[0].Rows[liRowCount][0].ToString() == "NET CHANGE")
                        {
                            setGreyBorder(loCell);
                            //added on 28Feb2011 to change font size for total
                            if (liColumnCount != 0)
                            {
                                lochunk = new Chunk(lsFormatedString, Font7Bold());
                            }
                        }
                    }

                    if (lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE %"))
                    {

                        lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                        try
                        {
                            lsFormatedString = String.Format("{0:#,###0.0%;(#,###0.0%)}", Convert.ToDecimal(lsFormatedString) / 100);
                        }
                        catch
                        {

                        }
                        //changed on 02/25/2011
                        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                        lochunk = new Chunk(lsFormatedString, Font8Bold());
                        //added on 28Feb2011 to change font size for total
                        if (liColumnCount != 0)
                        {
                            lochunk = new Chunk(lsFormatedString, Font7Bold());
                        }


                    }


                }
                else
                {
                    if (liColumnCount == 0 && !checkTrue(lodataset, liRowCount, "_Ssi_UnderlineFlg"))
                    {
                        String abc = "          " + lodataset.Tables[0].Rows[liRowCount][1].ToString();
                        //changed on 02/25/2011
                        //lochunk = new Chunk(abc, Font9Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                        lochunk = new Chunk(abc, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                    }



                }
                if (checkTrue(lodataset, liRowCount, "_Ssi_TabFlg") && !checkTrue(lodataset, liRowCount, "_Ssi_UnderlineFlg"))
                {
                    if (liColumnCount == 0)
                    {
                        String abc = "          " + "          " + lodataset.Tables[0].Rows[liRowCount][1].ToString();
                        //changed on 02/25/2011
                        //lochunk = new Chunk(abc, Font8Grey());
                        lochunk = new Chunk(abc, Font7Grey());
                    }
                    else
                    {
                        //changed on 02/25/2011
                        //lochunk = new Chunk(lsFormatedString, Font8Grey());
                        lochunk = new Chunk(lsFormatedString, Font7Grey());
                    }
                }

                //CONDITION FOR SUPERBOLDFLAG
                checkTrue(lodataset, liRowCount, "_Ssi_SuperBoldFlg", loCell, new iTextSharp.text.Color(183, 221, 232));
                //====added on 28Feb2011 to change font size for total====
                if (checkTrue(lodataset, liRowCount, "_Ssi_SuperBoldFlg"))
                {
                    if (liColumnCount != 0)
                    {
                        lochunk = new Chunk(lsFormatedString, Font7Bold());
                    }
                }
                /*=====END=====*/

                if (checkTrue(lodataset, liRowCount, "_Ssi_UnderlineFlg"))
                {
                    if (liColumnCount == 0)
                    {
                        String abc = "          " + "          " + "Total";
                        //changed on 02/25/2011
                        //lochunk = new Chunk(abc, Font8Normal());
                        lochunk = new Chunk(abc, Font7Normal());
                    }
                    setTopWidthBlack(loCell);
                    setBottomWidthWhite(loCell);

                }
                loCell.Add(lochunk);
                loTable.AddCell(loCell);
            }

            if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
            {
                document.Add(loTable);
                liCurrentPage = liCurrentPage + 1;
                //  document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt)); //Commented -- FooterLogic

                objCombinedReports.SetTotalPageCount("Asset Distribution");
            }
        }

        document.Close();

        FileInfo loFile = new FileInfo(ls);
        //try
        //{
        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch { }
    }

    public void generateCoversheetPDF(String lsDateString, String fsFinalLocation, String fsAllocationGroup, String fsHouseholdName, String fsContactId, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        int TotalReportCount = foTable.Rows.Count;
        int UpperspaceCount = 0;
        int RptTitleCount = 0;
        int MainTitleLengthCount = 0;

        String lsfamilyName = fsHouseholdName;
        int liCommaCounter = lsfamilyName.IndexOf(",");
        int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
        if (liCommaCounter > 0 && liSpaceCounter > 0)
            lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
        else
            lsfamilyName = lsfamilyName;

        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsfamilyName = fsAllocationGroup;
        }

        lsfamilyName = "";

        if (fsKeyContactID == fsContactId)
        {
            //lsfamilyName = fsHouseHoldTitle;
            //if (!String.IsNullOrEmpty(fsAllocationGroup))
            //    lsfamilyName = fsAllocationGroup;
            if (!String.IsNullOrEmpty(lsFinalReportTitle))
                lsfamilyName = lsFinalReportTitle;
        }
        else
        {
            lsfamilyName = "Reports for " + fsDisplayContactName;
        }

        //if (!String.IsNullOrEmpty(lsFinalReportTitle))
        //    lsfamilyName = lsFinalReportTitle;

        MainTitleLengthCount = lsfamilyName.Length;


        if (TotalReportCount > 0 && TotalReportCount < 6)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
            else
            {
                UpperspaceCount = 12;
                RptTitleCount = 11;
            }

        }
        else if (TotalReportCount >= 6 && TotalReportCount < 9)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 7;
                RptTitleCount = 13;
            }
            else
            {
                UpperspaceCount = 9;
                RptTitleCount = 14;
            }
        }
        else if (TotalReportCount >= 9 && TotalReportCount < 11)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 5;
                RptTitleCount = 12;
            }
            else
            {
                UpperspaceCount = 7;
                RptTitleCount = 13;
            }
        }
        else if (TotalReportCount >= 11 && TotalReportCount < 13)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 4;
                RptTitleCount = 16;
            }
            else
            {
                UpperspaceCount = 6;
                RptTitleCount = 17;
            }
        }
        else
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 1;
                RptTitleCount = 16;
            }
            else
            {
                UpperspaceCount = 2;
                RptTitleCount = 16;
            }
        }

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 80, 80, 31, 5);
        //String ls = Server.MapPath("") + "/a" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        String ls = fsFinalLocation.Replace(".xls", ".pdf");

        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        document.Open();
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(2);
        loTable.Width = 100;
        int[] headerwidths = { 39, 45 }; //{ 47, 35 }
        loTable.SetWidths(headerwidths);
        loTable.Border = 0;

        iTextSharp.text.Cell loCell = new Cell();
        Chunk loChunk = new Chunk();
        for (int liCounter = 0; liCounter < UpperspaceCount; liCounter++)//13//7
        {
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 1;
            loCell.Border = 0;
            loTable.AddCell(loCell);

        }

        loCell = new Cell();
        loChunk = new Chunk(lsfamilyName, setFontsAll(26, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 1;
        if (MainTitleLengthCount >= 54)
        {
            loCell.Leading = 25f;
        }
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(Convert.ToDateTime(lsDateString).ToString("MMMM dd, yyyy") + "", setFontsAll(12, 0, 1));
        loCell.Add(loChunk);
        loCell.Leading = 25f;
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 1;
        loTable.AddCell(loCell);


        for (int liCounter = 0; liCounter < 2; liCounter++)//4
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 1;
            loCell.Border = 0;
            loTable.AddCell(loCell);
        }

        int rowcount = foTable.Rows.Count;
        int rowdiff = 0;
        int j = 0;
        for (int liCounter = 0; liCounter < RptTitleCount; liCounter++)
        {
            rowdiff = RptTitleCount - rowcount;
            if (liCounter >= rowdiff)
            {
                if (fsContactId == Convert.ToString(foTable.Rows[j]["ssi_ContactID"]).Replace(",", ""))
                {
                    loCell = new Cell();
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 0;
                    loCell.HorizontalAlignment = 0;
                    loCell.Leading = 0.3f;//0.7f
                    loCell.Border = 1;
                    loTable.AddCell(loCell);

                    loCell = new Cell();
                    String lsAllocationGroupNEW = Convert.ToString(foTable.Rows[j]["Ssi_AllocationGroup"]);

                    String lsFinalTitleAfterChange = String.Empty;
                    if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"])))
                        lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"]);

                    if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"])))
                        lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"]);

                    fsGAorTIAflag = Convert.ToString(foTable.Rows[j]["ssi_gaortia"]);
                    fsDiscretionaryFlg = Convert.ToString(foTable.Rows[j]["Discretionary Flag"]);

                    if (fsGAorTIAflag == "GA")
                    {
                        if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                            loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                        else
                            loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                    }
                    else
                    {
                        if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                            loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                        else
                            loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));
                    }
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 1;
                    loCell.Border = 0;
                    loCell.Width = 45;//20                    
                    loCell.HorizontalAlignment = 0;
                    loTable.AddCell(loCell);
                    j++;
                }
            }
            else
            {
                if (liCounter == rowdiff - 1)
                {
                    loCell = new Cell();
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 0;
                    loCell.Leading = 1f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Border = 1;
                    loTable.AddCell(loCell);

                    loCell = new Cell();
                    loChunk = new Chunk("Reports included:", setFontsAll(10, 0, 1));
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 1;
                    loCell.Border = 0;
                    loCell.HorizontalAlignment = 0;
                    loTable.AddCell(loCell);
                }
                else
                {
                    loCell = new Cell();
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 2;
                    loCell.HorizontalAlignment = 1;
                    loCell.Border = 0;
                    loTable.AddCell(loCell);
                }
            }

        }

        for (int liCounter1 = 0; liCounter1 < 2; liCounter1++)
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 1;
            loCell.Border = 0;
            loTable.AddCell(loCell);

        }


        loCell = new Cell();
        loChunk = new Chunk("The values shown for the current period and the prior period are subject to the availability of information. In particular, certain non-marketable investments such as commercial real estate and private equity holdings do not provide frequent valuations. In these and other cases, we have either carried the investments at cost or used the general partner's most recent quarterly valuation estimates adjusted for subsequent investments or distributions.", setFontsAll(8, 0, 1, new iTextSharp.text.Color(150, 150, 150)));
        loCell.Add(loChunk);
        loCell.Leading = 9f;
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loTable.AddCell(loCell);
        int liFindRow = foTable.Rows.Count * 2;
        //for (int liCounterww = 0; liCounterww < 19 - liFindRow; liCounterww++)
        for (int liCounterww = 0; liCounterww < 3; liCounterww++)
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 0;
            loCell.Leading = 5f;
            loCell.Border = 0;
            loTable.AddCell(loCell);
        }

        loCell = new Cell();
        loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loCell.Add(loChunk);
        loCell.BorderWidth = 0;
        loCell.Colspan = 2;
        loCell.HorizontalAlignment = 2;
        // loTable.AddCell(loCell);//Commented -- FooterLogic

        document.Add(loTable);

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png"); //(Server.MapPath("") + @"\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);
        document.Close();
        try
        {
            FileInfo loFile = new FileInfo(ls);
            // loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        }
        catch { }
    }

    public string getFinalSp(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        String lsSQL = "";
        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsSQL = "SP_R_Advent_Report_Allocation_NEW_GA @AllocationGroupNameTxt='" + fsAllocationGroup.Replace("'", "''") + "', ";
        }
        else
        {
            lsSQL = "SP_R_Advent_Report_Other_NEW_GA";
        }
        lsSQL = lsSQL + " @UUID = '" + System.Guid.NewGuid().ToString() + "'," +
        "@HouseholdName = '" + fsHouseholdName.Replace("'", "''") + "',";

        if (!String.IsNullOrEmpty(fsAsofDate))
        {
            lsSQL += "@EndAsofDate = '" + Convert.ToDateTime(fsAsofDate).ToShortDateString() + "',";
        }
        else
        {
            lsSQL += "@EndAsofDate = " + "null" + ",";
        }
        if (!String.IsNullOrEmpty(fsSPriorDate))
        {
            lsSQL += "@StartAsofDate = '" + Convert.ToDateTime(fsSPriorDate).ToShortDateString() + "',";
        }
        else
        {
            lsSQL += "@StartAsofDate = " + "null" + ",";
        }

        if (!String.IsNullOrEmpty(fsGAorTIAflag))
        {
            lsSQL += "@PositionGAFlagTxt = '" + fsGAorTIAflag + "',";
        }
        else
        {
            lsSQL += "@PositionGAFlagTxt = " + "null" + ",";
        }

        if (fsDiscretionaryFlg.ToUpper() == "TRUE")
            fsDiscretionaryFlg = "1";
        else if (fsDiscretionaryFlg.ToUpper() == "FALSE")
            fsDiscretionaryFlg = "0";
        else
            fsDiscretionaryFlg = "null";


        lsSQL += "@LookThruDetailTxt = '" + fsLookthrogh.Replace("'", "''") + "'," +
                    "@ContactFullNameTxt = '" + fsContactFullname.Replace("'", "''") + "'," +
                    "@VersionTxt = '" + fsVersion.Replace("'", "''") + "'," +
                    "@summaryflgtxt = '" + fsSummaryFlag + "'," +
                    "@ReportType = '" + fsAllignment + "'," +
                    "@ReportGroupFlg = " + fsReportGroupflag +
                    ",@Report2GroupFlg = " + fsReportgroupflag2 +
                    ",@DiscretionaryFlg = " + fsDiscretionaryFlg;
        return lsSQL;
    }

    private DataTable GetDataTable(String BatchID)
    {
        string greshamquery;
        int totalCount = 0;
        //string ReportOpFolder2 = ConfigurationManager.AppSettings.Keys[1].ToString();
        //string Gresham_String = "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {
            object PriorDate = "null";// txtPriorDate.Text == "" ? "null" : "'" + txtPriorDate.Text + "'";
            object EndDate = ddlUpdateMonth.SelectedValue == "N/A" ? "null" : "'" + ddlUpdateMonth.SelectedValue + "'";
            //object EndDate = null;

            //object NoComparison = chkNoComparison.Checked == false ? 0 : 1;
            greshamquery = "sp_s_batch @BatchIdListTxt='" + BatchID + "',@ssi_approvalreqd=true,@PriorDT=" + PriorDate + ",@EndDT=" + EndDate;// +",@NoComparisonLineFlg=" + Convert.ToBoolean(chkNoComparison.Checked);
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;
            // Response.Write("Batch: " + DateTime.Now.ToString());
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            totalCount = 0;
            Response.Write("sp_s_batch sp fails error desc:" + exc.Detail.InnerText);
            // LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }
        catch (Exception exc)
        {
            bProceed = false;
            totalCount = 0;
            Response.Write("sp_S_batch sp fails error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }

        return ds_gresham.Tables[0];
    }

    private DataTable GetBatchList(string HouseholdID, string BatchType)
    {
        string greshamquery;
        int totalCount = 0;
        //string ReportOpFolder2 = ConfigurationManager.AppSettings.Keys[1].ToString();



        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {
            HouseholdID = HouseholdID == "0" ? "null" : "'" + HouseholdID + "'";
            BatchType = BatchType == "0" ? "null" : "'" + BatchType + "'";
            //greshamquery = "sp_s_batch_list @HouseHoldId =" + HouseholdID + ",@BatchType=" + BatchType;
            greshamquery = "sp_s_batch_list_CONSOLIDETED @HouseHoldId =" + HouseholdID + ",@ssi_approvalreqd=true";
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;
            // Response.Write("Batch: " + DateTime.Now.ToString());
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            totalCount = 0;
            Response.Write("sp_s_batch_list_CONSOLIDETED sp fails error desc:" + exc.Detail.InnerText);
            // LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }
        catch (Exception exc)
        {
            bProceed = false;
            totalCount = 0;
            Response.Write("sp_s_batch_list_CONSOLIDETED sp fails error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }

        return ds_gresham.Tables[0];
    }

    protected void chkbNCAll_CheckedChanged(object sender, EventArgs e)
    {
        for (int i = 0; i < gvList.Rows.Count; i++)
        {
            System.Web.UI.WebControls.TextBox txtCAUpdateValue2 = (System.Web.UI.WebControls.TextBox)gvList.Rows[i].FindControl("txtCAUpdateValue");
            CheckBox chkbxNC2 = (CheckBox)gvList.Rows[i].FindControl("chkbNC");

            if (chkbxNCAll.Checked)
            {
                chkbxNC2.Checked = true;
                txtCAUpdateValue2.Text = "";
            }
            else
            {
                chkbxNC2.Checked = false;
                //txtCAUpdateValue2.Text = "";
            }
        }
    }


    #region PDF Report Supporting functions
    public void setTopWidthBlack(Cell foCell)
    {
        foCell.BorderColor = iTextSharp.text.Color.BLACK;
        foCell.Border = iTextSharp.text.Rectangle.TOP_BORDER;
        foCell.BorderWidth = 0.1F;
    }
    public void setGreyBorder(DataSet foDataset, Cell foCell, int fiRowCount)
    {
        try
        {
            if (checkTrue(foDataset, fiRowCount, "_Ssi_UnderlineFlg") || checkTrue(foDataset, fiRowCount, "_Ssi_BoldFlg") || checkTrue(foDataset, fiRowCount, "_Ssi_SuperBoldFlg"))
            {
                setBottomWidthWhite(foCell);
            }
            if (checkTrue(foDataset, fiRowCount + 1, "_Ssi_UnderlineFlg") || checkTrue(foDataset, fiRowCount + 1, "_Ssi_BoldFlg") || checkTrue(foDataset, fiRowCount + 1, "_Ssi_SuperBoldFlg"))
            {
                setBottomWidthWhite(foCell);
            }
            else
            {
                foCell.BorderWidthBottom = 0.1F;
                //foCell.BorderColorBottom = new iTextSharp.text.Color(242, 242, 242);
                foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);
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
    public void setBottomWidthWhite(Cell foCell)
    {
        foCell.BorderWidthBottom = 0;
        foCell.BorderColorBottom = new iTextSharp.text.Color(255, 255, 255);
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

    public void setHeader(Document foDocument, DataSet loInsertdataset)
    {
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, 4);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk loParagraph = new Chunk();


        //     Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));
        Chunk lochunk = new Chunk(lsFamiliesName, setFontsAll(14, 1, 0));
        // loParagraph.Chunks.Add(lochunk);
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);
        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        loCell.HorizontalAlignment = 1;

        lochunk = new Chunk("\n" + lsGAorTIAHeader, setFontsAll(10, 0, 0));
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + lsDistributionName, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
        loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
        loCell.Leading = 13F;
        loTable.AddCell(loCell);



        Boolean lbCheckFoMarket = false;
        for (int liColumnCount = 0; liColumnCount < loInsertdataset.Tables[0].Columns.Count; liColumnCount++)
        {
            if (liColumnCount == 0)
            {
                //changed on 02/25/2011
                //lochunk = new Chunk("", setFontsAll(9, 1, 0));
                lochunk = new Chunk("", setFontsAll(7, 1, 0));
            }
            else
            {
                //changed on 02/25/2011
                lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Replace(" Market Value", ""), setFontsAll(7, 1, 0));
                //lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Replace(" Market Value", ""), setFontsAll(9, 1, 0));
                if (loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName.Contains(" Market Value"))
                    lbCheckFoMarket = true;

            }
            loCell = new Cell();

            loCell.Add(lochunk);
            loCell.Border = 0;
            loCell.NoWrap = true;//true;

            if (liColumnCount != 0)
            {
                loCell.HorizontalAlignment = 2;
            }
            if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
            {
                loCell.Leading = 10f;//8
                loCell.MaxLines = 5;
                //loCell.Leading = 9f;
            }
            loCell.Leading = 10f;//8
            loCell.VerticalAlignment = 6;//5 ,6 bottom : WASTE VALUES - 3,4
            loTable.AddCell(loCell);

        }


        //loCell = new Cell("");
        //lochunk = new Chunk("Market Value", FontFactory.GetFont(lsStringName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9, Font.BOLD));
        if (lbCheckFoMarket)
        {
            for (int liColumnCount = 0; liColumnCount < loInsertdataset.Tables[0].Columns.Count; liColumnCount++)
            {
                //Response.Write("<br>"+liColumnCount + "<br>");
                loCell.Border = 0;
                loCell.NoWrap = true;

                loCell = new Cell();
                if (liColumnCount != 0)
                {
                    loCell.HorizontalAlignment = 2;
                }
                if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
                {
                    loCell.NoWrap = false;
                }
                if (loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName.Contains(" Market Value"))
                {
                    // Response.Write("<br>" + liColumnCount + " In<br>");
                    //changed on 02/25/2011
                    //lochunk = new Chunk("Market Value", setFontsAll(9, 1, 0));
                    lochunk = new Chunk("Market Value", setFontsAll(7, 1, 0));
                }
                else
                {
                    //Response.Write("<br>" + liColumnCount + " Out<br>");
                    //changed on 02/25/2011
                    //lochunk = new Chunk("", setFontsAll(9, 1, 0));
                    lochunk = new Chunk("", setFontsAll(7, 1, 0));

                }
                loCell.Add(lochunk);
                loCell.Border = 0;
                loCell.NoWrap = true;
                loCell.Leading = 6f;
                loTable.AddCell(loCell);
            }
        }

        //loCell = new Cell();
        //loCell.Add(lochunk);
        //loCell.Border = 0;
        //loCell.NoWrap = true;
        //loTable.AddCell(loCell);
        //loCell = new Cell("");

        //loCell.Border = 0;
        //loCell.NoWrap = true;
        //loTable.AddCell(loCell);

        foDocument.Add(loTable);
        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        foDocument.Add(png);
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
                int[] headerwidths5 = { 30, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths5);
                fotable.Width = 67;
                break;
            case "6":
                int[] headerwidths6 = { 30, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths6);
                fotable.Width = 76;
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
                int[] headerwidths9 = { 25, 9, 9, 9, 9, 9, 9, 9, 9 };
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
    public Boolean checkTrue(DataSet foDataset, int fiRowCount, String fsField)
    {
        Boolean lblReturn = false;
        if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString().ToUpper() == "TRUE")
        {
            lblReturn = true;
        }
        return lblReturn;

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
        //return setFontsAll(7, 0, 0, new iTextSharp.text.Color(165, 165, 165));
        return setFontsAll(7, 0, 0, new iTextSharp.text.Color(0, 102, 153));
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

    public iTextSharp.text.Font Font7Bold()
    {
        return setFontsAll(7, 1, 0);
    }

    public void checkTrue(DataSet foDataset, int fiRowCount, String fsField, Cell foCell, iTextSharp.text.Color foColor)
    {

        if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString().ToUpper() == "TRUE")
        {
            foCell.BackgroundColor = foColor;
        }


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


    #region OLDCODE Commented CRM2016 UPGRADE
    ///// <summary>
    ///// Set up the CRM Service.
    ///// </summary>
    ///// <param name="organizationName">My Organization</param>
    ///// <returns>CrmService configured with AD Authentication</returns>
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
    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        //   IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        string strName = "";// claimsIdentity.Name;
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

    protected void ddlSecOwner_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindHousehold();
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");
        ddlUpdateMonth.ClearSelection();
        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
    }
    protected void gvCashinTransit_RowDataBound(object sender, GridViewRowEventArgs e)
    {


        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            decimal OrgTotal;
            decimal CurrenValue;

            //set values into string variables
            string _Headerflg = DataBinder.Eval(e.Row.DataItem, "_Headerflg").ToString();
            string _Dataflg = DataBinder.Eval(e.Row.DataItem, "_Dataflg").ToString();
            string _800Acctflg = DataBinder.Eval(e.Row.DataItem, "_800Acctflg").ToString();
            string _Totalflg = DataBinder.Eval(e.Row.DataItem, "_Totalflg").ToString();
            string _OrderbyAccount = DataBinder.Eval(e.Row.DataItem, "_OrderbyAcount").ToString();
            string asofdate = DataBinder.Eval(e.Row.DataItem, "ssi_asofdate").ToString().Replace("As Of Date", "").Replace("ssi_asofdate", "");

            System.Web.UI.WebControls.TextBox txtCAUpdateValue = (System.Web.UI.WebControls.TextBox)e.Row.FindControl("txtCashinTransit");
            CheckBox chkSelectNC = (CheckBox)e.Row.FindControl("chkSelectNC");

            chkSelectNC.Attributes.Add("onclick", "EnableDisable('" + chkSelectNC.ClientID + "','" + txtCAUpdateValue.ClientID + "')");

            DateTime AsOfDate;
            DateTime CompareMonthDate = Convert.ToDateTime("5/31/2011");
            if (asofdate != "")
            {
                AsOfDate = Convert.ToDateTime(asofdate);
                if (CompareMonthDate.CompareTo(AsOfDate) > 0)
                {
                    txtCAUpdateValue.BackColor = System.Drawing.Color.LightGray;
                    txtCAUpdateValue.Enabled = false;
                    chkSelectNC.Enabled = false;
                }
            }




            if (e.Row.Cells[21].Text == "0.5")
            {
                e.Row.Font.Bold = true;
                //e.Row.Cells[1].Text = e.Row.Cells[1].Text.Replace(e.Row.Cells[1].Text, "C." + " " + e.Row.Cells[1].Text);

                e.Row.Cells[8].Text = "";
                e.Row.Cells[9].Text = "";

                e.Row.Cells[0].Text = e.Row.Cells[0].Text.Replace("&nbsp;", e.Row.Cells[1].Text);
                e.Row.Cells[1].Text = e.Row.Cells[1].Text.Replace(e.Row.Cells[1].Text, "");
                e.Row.Cells[0].ColumnSpan = 3;
                e.Row.Cells[0].VerticalAlign = VerticalAlign.Bottom;
                e.Row.Cells[0].HorizontalAlign = HorizontalAlign.Left;
                //e.Row.Cells.RemoveAt(0);
                //e.Row.Cells.RemoveAt(2);

                e.Row.Cells[0].Height = 50;
                //e.Row.Cells[0].VerticalAlign = VerticalAlign.Bottom;

                e.Row.Cells[0].CssClass = "CellTitle";
                e.Row.Cells[1].CssClass = "CellTitle";
                e.Row.Cells[2].CssClass = "CellTitle";
                e.Row.Cells[3].CssClass = "CellTitle";
                e.Row.Cells[4].CssClass = "CellTitle";
                e.Row.Cells[5].CssClass = "CellTitle";
                e.Row.Cells[6].CssClass = "CellTitle";
                e.Row.Cells[7].CssClass = "CellTitle";
                e.Row.Cells[8].CssClass = "CellTitle";
                e.Row.Cells[9].CssClass = "CellTitle";

                chkSelectNC.Visible = false;
                txtCAUpdateValue.Visible = false;

            }

            if (e.Row.Cells[17].Text.ToUpper() == "TRUE")
            {
                e.Row.BackColor = System.Drawing.Color.LightBlue;
                e.Row.Font.Bold = true;
                e.Row.Cells[11].Text = "N/C";
                e.Row.Cells[10].Text = "CA Update Value";

                chkSelectNC.Visible = false;
                txtCAUpdateValue.Visible = false;
                e.Row.Cells[0].CssClass = "CellHeader";
                e.Row.Cells[1].CssClass = "CellHeader";
                e.Row.Cells[2].CssClass = "CellHeader";
                e.Row.Cells[3].CssClass = "CellHeader";
                e.Row.Cells[4].CssClass = "CellHeader";
                e.Row.Cells[5].CssClass = "CellHeader";
                e.Row.Cells[6].CssClass = "CellHeader";
                e.Row.Cells[7].CssClass = "CellHeader";
                e.Row.Cells[8].CssClass = "CellHeader";
                e.Row.Cells[9].CssClass = "CellHeader";
                e.Row.Cells[10].CssClass = "CellHeader";
                e.Row.Cells[11].CssClass = "CellHeader";

            }

            if (e.Row.Cells[19].Text == "True")
            {
                e.Row.BackColor = System.Drawing.Color.LightGray;

                e.Row.Cells[8].Text = e.Row.Cells[8].Text.Replace(e.Row.Cells[8].Text, string.Format("{0:$#,##0.00}", Convert.ToDecimal(decimal.Parse(e.Row.Cells[8].Text, System.Globalization.NumberStyles.Any))));
                e.Row.Cells[9].Text = e.Row.Cells[9].Text.Replace(e.Row.Cells[9].Text, string.Format("{0:$#,##0.00}", Convert.ToDecimal(decimal.Parse(e.Row.Cells[9].Text, System.Globalization.NumberStyles.Any))));
                //e.Row.Cells[9].Text = e.Row.Cells[9].Text.Replace(e.Row.Cells[9].Text, string.Format("${0:0,0.00}", Convert.ToDecimal(e.Row.Cells[9].Text)));

                chkSelectNC.Visible = false;
                txtCAUpdateValue.Visible = false;
                e.Row.Cells[9].Style.Add("text-align", "right");
                e.Row.Cells[0].CssClass = "CellHeader";
                e.Row.Cells[1].CssClass = "CellHeader";
                e.Row.Cells[2].CssClass = "CellHeader";
                e.Row.Cells[3].CssClass = "CellHeader";
                e.Row.Cells[4].CssClass = "CellHeader";
                e.Row.Cells[5].CssClass = "CellHeader";
                e.Row.Cells[6].CssClass = "CellHeader";
                e.Row.Cells[7].CssClass = "CellHeader";
                e.Row.Cells[8].CssClass = "CellHeader";
                e.Row.Cells[9].CssClass = "CellHeader";
                e.Row.Cells[10].CssClass = "CellHeader";
                e.Row.Cells[11].CssClass = "CellHeader";

            }

            if (e.Row.Cells[18].Text == "True")
            {
                e.Row.Cells[8].Text = e.Row.Cells[8].Text.Replace(e.Row.Cells[8].Text, string.Format("{0:$#,##0.00}", Convert.ToDecimal(decimal.Parse(e.Row.Cells[8].Text, System.Globalization.NumberStyles.Any))));
                e.Row.Cells[9].Text = e.Row.Cells[9].Text.Replace(e.Row.Cells[9].Text, string.Format("{0:$#,##0.00}", Convert.ToDecimal(decimal.Parse(e.Row.Cells[9].Text, System.Globalization.NumberStyles.Any))));
                //e.Row.Cells[8].Text = e.Row.Cells[8].Text.Replace(e.Row.Cells[8].Text, "$" + " " + e.Row.Cells[8].Text);
                //e.Row.Cells[9].Text = e.Row.Cells[9].Text.Replace(e.Row.Cells[9].Text, "$" + " " + e.Row.Cells[9].Text);
                txtCAUpdateValue.Visible = true;
                chkSelectNC.Visible = true;
                e.Row.Cells[9].Style.Add("text-align", "right");
                e.Row.Cells[0].CssClass = "CellHeader";
                e.Row.Cells[1].CssClass = "CellHeader";
                e.Row.Cells[2].CssClass = "CellHeader";
                e.Row.Cells[3].CssClass = "CellHeader";
                e.Row.Cells[4].CssClass = "CellHeader";
                e.Row.Cells[5].CssClass = "CellHeader";
                e.Row.Cells[6].CssClass = "CellHeader";
                e.Row.Cells[7].CssClass = "CellHeader";
                e.Row.Cells[8].CssClass = "CellHeader";
                e.Row.Cells[9].CssClass = "CellHeader";
                e.Row.Cells[10].CssClass = "CellHeader";
                e.Row.Cells[11].CssClass = "CellHeader";

                // For Auto refreshing the grid values
                //validateCAUpdateValue
                txtCAUpdateValue.Attributes.Add("OnChange", "javascript:return Refressh();");
                //txtCAUpdateValue.Attributes.Add("onkeyup", "Refressh('" + txtCAUpdateValue.ClientID + "');");

            }
            else if (e.Row.Cells[18].Text == "False")
            {
                txtCAUpdateValue.Visible = false;
                chkSelectNC.Visible = false;
            }

            if (e.Row.Cells[20].Text == "True")
            {

                e.Row.Cells[9].Text = e.Row.Cells[9].Text.Replace(e.Row.Cells[9].Text, string.Format("{0:$#,##0.00}", Convert.ToDecimal(decimal.Parse(e.Row.Cells[9].Text, System.Globalization.NumberStyles.Any))));
                //e.Row.Cells[9].Text = e.Row.Cells[9].Text.Replace(e.Row.Cells[9].Text, "$" + " " + decimal.Parse(e.Row.Cells[9].Text, System.Globalization.NumberStyles.Any));
                e.Row.Cells[9].Style.Add("text-align", "right");
                chkSelectNC.Visible = false;
                txtCAUpdateValue.Visible = false;

                e.Row.BackColor = System.Drawing.Color.LightSkyBlue;
                e.Row.Cells[8].Text = "Total Cash in Transit";
                //e.Row.Cells[8].ColumnSpan = 2;
                //e.Row.Cells.RemoveAt(7);
                e.Row.Cells[0].CssClass = "CellTotLeft";
                e.Row.Cells[10].CssClass = "CellTotRight";
                e.Row.Cells[11].CssClass = "CellTotRight";

                e.Row.Cells[1].CssClass = "CellTitle";
                e.Row.Cells[2].CssClass = "CellTitle";
                e.Row.Cells[3].CssClass = "CellTitle";
                e.Row.Cells[4].CssClass = "CellTitle";
                e.Row.Cells[5].CssClass = "CellTitle";
                e.Row.Cells[6].CssClass = "CellTitle";
                e.Row.Cells[7].CssClass = "CellTitle";
                e.Row.Cells[8].CssClass = "CellTotLeft";
                e.Row.Cells[9].CssClass = "CellTotLeft";


            }




        }

    }

    protected void txtCashinTransit_TextChanged1(object sender, EventArgs e)
    {

        GridViewRow r = (GridViewRow)((DataControlFieldCell)((System.Web.UI.WebControls.TextBox)sender).Parent).Parent;
        int rowIndex = Convert.ToInt32(r.RowIndex);

        System.Web.UI.WebControls.TextBox txt1 = (System.Web.UI.WebControls.TextBox)sender;
        //CheckBox chk = (CheckBox)sender;

        if (gvCashinTransit.Rows[rowIndex + 1].Cells[8].Text != "")
        {
            decimal OrgTotal;
            decimal CurrenValue;
            if (gvCashinTransit.Rows[rowIndex].Cells[18].Text.ToUpper() == "TRUE")
            {
                OrgTotal = Convert.ToDecimal(decimal.Parse(gvCashinTransit.Rows[rowIndex + 1].Cells[22].Text.Replace("$ ", ""), System.Globalization.NumberStyles.Any));

                if (txt1.Text == "")
                {
                    txt1.Text = "0";
                    CurrenValue = Convert.ToDecimal(txt1.Text);
                    txt1.Text = "";
                }
                else
                {
                    CurrenValue = Convert.ToDecimal(txt1.Text);
                }
                decimal FinalTotal = OrgTotal + CurrenValue;

                if (txt1.Text != "0" && txt1.Text != "")
                {
                    FinalTotal = FinalTotal - Convert.ToDecimal(decimal.Parse(gvCashinTransit.Rows[rowIndex].Cells[9].Text, System.Globalization.NumberStyles.Any));
                }
                else if (txt1.Text == "0")
                {
                    FinalTotal = FinalTotal - Convert.ToDecimal(decimal.Parse(gvCashinTransit.Rows[rowIndex].Cells[9].Text, System.Globalization.NumberStyles.Any));
                }
                gvCashinTransit.Rows[rowIndex + 1].Cells[9].Text = string.Format("${0:#,000.00}", Convert.ToString(FinalTotal));//"$" + " " + FinalTotal;

            }

        }

        //FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'");


    }


    private void HideControls()
    {
        for (int i = 0; i < gvCashinTransit.Rows.Count; i++)
        {
            string gridcell21 = gvCashinTransit.Rows[i].Cells[21].Text.Replace("&nbsp;", "");
            if (gvCashinTransit.Rows[i].Cells[17].Text.ToUpper() == "TRUE")
            {
                System.Web.UI.WebControls.TextBox txtCAUpdateValue = (System.Web.UI.WebControls.TextBox)gvCashinTransit.Rows[i].FindControl("txtCashinTransit");
                CheckBox chkSelectNC = (CheckBox)gvCashinTransit.Rows[i].FindControl("chkSelectNC");
                txtCAUpdateValue.Visible = false;
                chkSelectNC.Visible = false;
                gvCashinTransit.Rows[i].Cells[10].Text = "CA Update Value";
                gvCashinTransit.Rows[i].Cells[11].Text = "N/C";
            }

            if (gridcell21 == "")
            {
                System.Web.UI.WebControls.TextBox txtCAUpdateValue = (System.Web.UI.WebControls.TextBox)gvCashinTransit.Rows[i].FindControl("txtCashinTransit");
                CheckBox chkSelectNC = (CheckBox)gvCashinTransit.Rows[i].FindControl("chkSelectNC");
                txtCAUpdateValue.Visible = false;
                chkSelectNC.Visible = false;
            }
        }
    }



    protected void Button1_Click(object sender, EventArgs e)
    {
        GenerateReport();
    }

    protected void btnRefresh_Click(object sender, EventArgs e)
    {
        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");

        lblMessage.Text = "Updated Successfully";
    }
    protected void ddlUpdateMonth_SelectedIndexChanged(object sender, EventArgs e)
    {
        string Value = ddlUpdateMonth.SelectedValue;
        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");

        ddlUpdateMonth.SelectedValue = Value;
    }
    protected void ddlView_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");
        ddlUpdateMonth.ClearSelection();
        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
    }
    protected void ddlFrequency_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");
        ddlUpdateMonth.ClearSelection();
        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
    }


    protected void gvList_SelectedIndexChanged(object sender, EventArgs e)
    {

    }
    protected void gvpopup_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            System.Web.UI.WebControls.TextBox txtCAUpdateValue = (System.Web.UI.WebControls.TextBox)e.Row.FindControl("txtCAUpdateValue");
            e.Row.Cells[3].Text = e.Row.Cells[3].Text.Replace(e.Row.Cells[3].Text, string.Format("{0:$#,##0.00}", Convert.ToDecimal(decimal.Parse(e.Row.Cells[3].Text, System.Globalization.NumberStyles.Any))));
            e.Row.Cells[4].Text = e.Row.Cells[4].Text.Replace(e.Row.Cells[4].Text, string.Format("{0:$#,##0.00}", Convert.ToDecimal(decimal.Parse(e.Row.Cells[4].Text, System.Globalization.NumberStyles.Any))));

            if (e.Row.RowIndex > 1)
            {
                if (e.Row.Cells[1].Text == gvList.Rows[e.Row.RowIndex - 1].Cells[1].Text)
                {
                    //gvList.Rows[e.Row.RowIndex - 1].Style["border-top-color"] = "Red";
                    //gvList.Rows[e.Row.RowIndex - 1].Style["border-top"] = "solid";
                    //gvList.Rows[e.Row.RowIndex - 1].Style["border-top-width"] = "thick";
                    // e.Row.Style.Add( border-top-color:Gray; border-top:solid; border-top-width:thick;
                    e.Row.Cells[2].Text = "";
                    e.Row.Cells[3].Text = "";
                    e.Row.Cells[4].Text = "";
                }
                else
                {

                }




            }

            if (e.Row.RowIndex > 0)
                if (e.Row.Cells[1].Text != gvList.Rows[e.Row.RowIndex - 1].Cells[1].Text)
                {
                    for (int i = 2; i < gvList.Columns.Count; i++)
                    {
                        e.Row.Cells[i].Style["border-style"] = "solid";
                        e.Row.Cells[i].Style["border-top-color"] = "#D8D8D8";
                        e.Row.Cells[i].Style["border-top-width"] = "3px";
                    }
                }





            if (e.Row.Cells[0].Text == "&nbsp;")
            {
                // Response.Write("pos1:" + gvList.Rows[e.Row.RowIndex].Cells[0].Text);

                txtCAUpdateValue.Visible = false;
            }


            if (e.Row.Cells[2].Text != "")
            {
                e.Row.BackColor = System.Drawing.Color.LightGray;

            }
        }
    }
    protected void btnsubmitpopup_Click(object sender, EventArgs e)
    {
        InsertIntoCRM();
        btnRefresh_Click(sender, e);
    }

    public void InsertIntoCRM()
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;
        lblMessage.Text = "";

        try
        {
            // service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            Label1.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            Label1.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;


        string posid = Convert.ToString(ViewState["ssi_positionid"]).Replace("'", "").Replace("'", "");

        try
        {
            // if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["posid"]) != "" && Convert.ToString(Request.QueryString["posid"]) != null)

            if (posid != "" && posid != null)
            {
                // update position
                for (int i = 0; i < gvpopup.Rows.Count; i++)
                {
                    System.Web.UI.WebControls.TextBox txtCAUpdateValue2 = (System.Web.UI.WebControls.TextBox)gvpopup.Rows[i].FindControl("txtCAUpdateValue");
                    //Response.Write(txtCAUpdateValue2.Text + ",");
                    if (txtCAUpdateValue2.Text == "")
                    {
                        txtCAUpdateValue2.Text = "0";
                    }

                    CheckBox chkbxNC2 = (CheckBox)gvpopup.Rows[i].FindControl("chkbNC");

                    //ssi_position objPosition = new ssi_position();
                    Entity objPosition = new Entity("ssi_position");


                    //objPosition.ssi_positionid = new Key();
                    //objPosition.ssi_positionid.Value = new Guid(Convert.ToString(Request.QueryString["posid"]));
                    objPosition["ssi_positionid"] = new Guid(posid);
                    //objPosition.ssi_datasource = new Picklist();
                    //objPosition.ssi_datasource.Value = 6; // value of CA Update

                    //objPosition.ssi_commitment = new CrmDecimal();
                    //objPosition.ssi_commitment.Value = Convert.ToDecimal(txtCAUpdateValue2.Text.Trim());
                    objPosition["ssi_commitment"] = Convert.ToDecimal(txtCAUpdateValue2.Text.Trim());
                    service.Update(objPosition);
                    successcount++;

                    // Response.Write(successcount.ToString());
                }

                if (successcount > 0)
                {
                    //lblMessage.Text = lblMessage.Text + "<br/>" + successcount.ToString() + " Records updated successfully";// for CLIENT SPECIFIC UPDATES";
                    //lblMessage.Visible = true;
                    performancepanel.Visible = false;
                    performancepopup.Hide();



                    //Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "close", "<script type='text/javascript'>ReturnToParent('true');</script>");
                }

                //FillGridForCashinTransit(_PositionID);

            }
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            performancepanel.Visible = true;
            performancepopup.Show();
            Label1.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            performancepanel.Visible = true;
            performancepopup.Show();
            Label1.Text = strDescription;
        }

    }
    protected void gvList_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        if (e.CommandName == "linkButton1")
        {
            //Determine the RowIndex of the Row whose LinkButton was clicked.
            int rowIndex = Convert.ToInt32(e.CommandArgument);

            //Reference the GridView Row.
            GridViewRow row = gvList.Rows[rowIndex];

            string ssi_positionid = "'" + row.Cells[0].Text + "'";
            ViewState["ssi_positionid"] = ssi_positionid;
            FillGridForCashinTransitPopUP(ssi_positionid);
            Label1.Text = "";
            performancepanel.Visible = true;
            performancepopup.Show();




        }
    }
    protected void btnCancel_Click(object sender, EventArgs e)
    {
        performancepanel.Visible = false;
        performancepopup.Hide();
        FillGridForCashinTransit("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'");
        BindGridView("'" + ddlHousehold.SelectedValue + "'", "'" + ddlUpdateMonth.SelectedValue + "'", "'" + ddlFrequency.SelectedValue + "'", "'" + ddlView.SelectedItem.Value + "'");
        //  btnRefresh_Click(sender, e);
    }
}
