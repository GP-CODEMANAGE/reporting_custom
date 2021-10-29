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
using System.Data.SqlClient;
//using CrmSdk;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Threading;
using Microsoft.IdentityModel.Claims;
using System.IO;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using System.ServiceModel.Description;

public partial class PerformanceAddUpdateTool : System.Web.UI.Page
{
    Logs lg = new Logs();
    public StreamWriter sw = null;
    public string Filename = "";
    DB clsdb = new DB();
    GeneralMethods clsGM = new GeneralMethods();
    string strDescription;
    bool bProceed = true;

    //string Type = string.Empty;

    //string Name = string.Empty;

    //string CurrAsOfDate = string.Empty;

    //string UUID = string.Empty;

    //string PervAsOfDate = string.Empty;
    int totalCount = 0;
    int successcount = 0;


    protected void Page_Load(object sender, EventArgs e)
    {
        DateTime dtmain = DateTime.Now;
        string LogFileName = string.Empty;
        LogFileName = "Log-" + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        sw = new StreamWriter(LogFileName);
        sw.Close();
        HttpContext.Current.Session["Filename"] = LogFileName;
        ViewState["Filename"] = LogFileName;

        Response.Write("LOGCREATED1" + LogFileName);

        LogFileName = (string)ViewState["Filename"];

        Session["Filename"] = LogFileName;

        lg.AddinLogFile(Session["Filename"].ToString(), "Start Page Load " + dtmain);

        IOrganizationService s1 = GetCrmService();
        //Session.Remove("ddlHousehold");
        //Session.Remove("txtAsOfDate");
        //Session.Remove("ddlAssociateOps");
        //Session.Remove("lblMessage");

        if (!IsPostBack)
        {
            ISOpsTeamMember();
            BindAssociateOrOps();
            BindHousehold();
            btnSubmit.Visible = false;
            btnSumbitTop.Visible = false;
            btnCanceltop.Visible = false;
            btnCancelbottom.Visible = false;

            if (HdIsOpsTeamMember.Value == "True")
            {
                chkUnsupressAll.Style.Add("display", "inline");
                chkUnsupressAll.Style.Add("cursor", "hand");
            }
            else
                chkUnsupressAll.Style.Add("display", "none");

            //Session.Remove("ddlHousehold");

            #region postback by session
            //if (Convert.ToString(Session["ddlHousehold"]) != "")
            //{
            //    string ddlHousehold1 = Convert.ToString(Session["ddlHousehold"]);
            //    ddlHousehold.SelectedValue = ddlHousehold.Items.FindByValue(ddlHousehold1).Value;
            //    if (Convert.ToString(Session["txtAsOfDate"]) != "")
            //    {

            //        string AsOfDate = Convert.ToString(Session["txtAsOfDate"]);
            //        txtAsOfDate.Text = AsOfDate;
            //    }

            //    if (Convert.ToString(Session["ddlAssociateOps"]) != "")
            //    {

            //        string ddlAssociateOps1 = Convert.ToString(Session["ddlAssociateOps"]);
            //        ddlAssociateOps.SelectedValue = ddlAssociateOps.Items.FindByValue(ddlAssociateOps1).Value;
            //    }


            //    BindGridView("'" + ddlHousehold.SelectedValue.ToString() + "'");

            //    if (Convert.ToString(Session["lblMessage"]) != "")
            //    {
            //        string Messeage = Convert.ToString(Session["lblMessage"]);
            //        lblMessage.Text = Messeage;
            //    }

            //    Session.Remove("ddlHousehold");
            //    Session.Remove("txtAsOfDate");
            //    Session.Remove("ddlAssociateOps");
            //    Session.Remove("lblMessage");


            //}

            #endregion
            //CrmService service = new CrmService();//this line added to intialize the service on load to fast the insert and update.			
            IOrganizationService service = null;
        }
    }

    protected void btnLoadData_Click(object sender, EventArgs e)
    {
        lblMessage.Text = "";

        BindGridView("'" + ddlHousehold.SelectedValue + "'");

        //performancepanel.Visible = true;
        //performancepopup.Show();

    }

    private void ISOpsTeamMember()
    {
        string OutParam = "null";
        SqlParameter[] param = new SqlParameter[2];
        param[0] = new SqlParameter("@SystemUserId", SqlDbType.UniqueIdentifier);
        param[0].Value = GetcurrentUser() == "" ? Guid.Empty : new Guid(GetcurrentUser());

        param[1] = new SqlParameter("@Return", SqlDbType.Bit);
        param[1].Direction = ParameterDirection.Output;

        clsdb.ExecuteScalar("SP_S_GET_OPERATIONGROUP_MEMBER", "StoredProcedure", "@Return", out OutParam, param);
        HdIsOpsTeamMember.Value = OutParam;
    }

    protected void ddlAssociateOps_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindHousehold();
        //BindGridView("'" + ddlHousehold.SelectedValue + "'");
        HideGrid();
    }
    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        //BindGridView("'0'");
        HideGrid();
    }
    private void HideGrid()
    {
        gvList.DataSource = null;
        gvList.DataBind();
        btnSumbitTop.Visible = false;
        btnSubmit.Visible = false;
        btnCanceltop.Visible = false;
        btnCancelbottom.Visible = false;
    }
    private void BindAssociateOrOps()
    {
        DataSet ds = new DataSet();

        string sqlstr = "[SP_S_GET_HOUSEHOLD_SECONDARY_OWNERS] @IsOperationGrpMem=" + HdIsOpsTeamMember.Value + "";

        ds = clsdb.getDataSet(sqlstr);
        ddlAssociateOps.DataTextField = "SecondaryOwnerIDName";
        ddlAssociateOps.DataValueField = "SecondaryOwnerID";

        ddlAssociateOps.DataSource = ds;
        ddlAssociateOps.DataBind();

        ddlAssociateOps.Items.Insert(0, "select");
        ddlAssociateOps.Items[0].Value = "0";

        if (HdIsOpsTeamMember.Value == "True")
        {
            //ddlAssociateOps.Items.Insert(1, "OPS Team");
            //ddlAssociateOps.Items[1].Value = "1";
        }
        else
        {
            ddlAssociateOps.Items.Insert(1, "All");
            ddlAssociateOps.Items[1].Value = "1";
        }
        ddlAssociateOps.SelectedIndex = 0;


    }
    private void BindHousehold()
    {
        DataSet ds = new DataSet();
        object AssociateOpsId = "null";
        object OpsTeamFlg = "null";
        if (ddlAssociateOps.SelectedValue == "0" || ddlAssociateOps.SelectedValue == "1")
        {
            AssociateOpsId = "null";
            OpsTeamFlg = "null";
        }
        else if (ddlAssociateOps.SelectedItem.Text == "OPS Team")
            OpsTeamFlg = "null";//OpsTeamFlg = "1";
        else
            AssociateOpsId = "'" + ddlAssociateOps.SelectedValue + "'";
        string sqlstr = "[SP_S_GET_PERF_ENTRY_HOUSEHOLDNAME] @SecondaryOwnerId=" + AssociateOpsId + ",@OpsTeamFlg=" + OpsTeamFlg + "";

        ds = clsdb.getDataSet(sqlstr);
        ddlHousehold.DataTextField = "name";
        ddlHousehold.DataValueField = "accountid";

        ddlHousehold.DataSource = ds;
        ddlHousehold.DataBind();

        ddlHousehold.Items.Insert(0, "select");
        ddlHousehold.Items[0].Value = "0";
        if (HdIsOpsTeamMember.Value == "True")
        {
            ddlHousehold.Items.Insert(1, "All");
            ddlHousehold.Items[1].Value = "1";
        }
    }

    private void BindGridView(string HouseholdId)
    {
        DataSet ds = new DataSet();
        string greshamquery = string.Empty;

        try
        {
            object OpsTeamFlg = "null";
            object SecondryOwnerId = "null";
            string AsOfDate = string.Empty;
            if (ddlAssociateOps.SelectedItem.Text == "OPS Team")
                OpsTeamFlg = "1";
            string HouseholdAll = HouseholdId;
            if (HouseholdId == "'0'" || HouseholdId == "'1'")
                HouseholdId = "null";

            if (ddlAssociateOps.SelectedValue != "0" && ddlAssociateOps.SelectedValue != "1" && ddlAssociateOps.SelectedValue != "")
                SecondryOwnerId = "'" + ddlAssociateOps.SelectedValue + "'";

            if (txtAsOfDate.Text != "")
                AsOfDate = txtAsOfDate.Text;
            else
                AsOfDate = DateTime.Now.AddDays(-DateTime.Now.Day).ToString("MM/dd/yyyy");

            greshamquery = "exec SP_S_PERFORMANCE_ENTRY_FORM @HHID=" + HouseholdId
            + ",@AsOfdate='" + AsOfDate + "'"
            + ",@OpsTeamFlg = " + OpsTeamFlg + " "
            + ",@SystemUserId='" + GetcurrentUser() + "'"
            + ",@SecondaryOwnerId=" + SecondryOwnerId + ""
            + ",@UnsupressFlg=" + chkUnsupressAll.Checked + "";

            ds = clsdb.getDataSet(greshamquery);

        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {

            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Detail.Message;
        }
        catch (Exception exc)
        {
            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Message;
        }

        gvList.Columns[5].Visible = true;
        gvList.Columns[6].Visible = true;
        gvList.Columns[7].Visible = true;
        gvList.Columns[8].Visible = true;
        gvList.Columns[9].Visible = true;

        gvList.Columns[10].Visible = true;
        gvList.Columns[11].Visible = true;
        gvList.Columns[12].Visible = true;
        gvList.Columns[13].Visible = true;
        gvList.Columns[14].Visible = true;

        gvList.DataSource = ds;
        gvList.DataBind();

        gvList.Columns[5].Visible = false;
        gvList.Columns[6].Visible = false;
        gvList.Columns[7].Visible = false;
        gvList.Columns[8].Visible = false;
        gvList.Columns[9].Visible = false;

        gvList.Columns[10].Visible = false;
        gvList.Columns[11].Visible = false;
        gvList.Columns[12].Visible = false;
        gvList.Columns[13].Visible = false;
        gvList.Columns[14].Visible = false;

        if (gvList.Rows.Count > 0)
        {
            btnSubmit.Visible = true;
            btnSumbitTop.Visible = true;
            btnCanceltop.Visible = true;
            btnCancelbottom.Visible = true;
        }
        else
        {
            btnSubmit.Visible = false;
            btnSumbitTop.Visible = false;
            btnCanceltop.Visible = false;
            btnCancelbottom.Visible = false;
            lblMessage.Text = "No Records Found.";
        }
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        int intResult = 0;
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        //string orgName = "GreshamPartners";
        //CrmService service = null;
        IOrganizationService service = null;

        lblMessage.Text = "";

        string UserId = GetcurrentUser();

        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
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

        // service.PreAuthenticate = true;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        for (int i = 0; i < gvList.Rows.Count; i++)
        {
            string _Header = gvList.Rows[i].Cells[9].Text.Replace("&nbsp;", "");
            //string UUID = Convert.ToString(gvList.Rows[i].Cells[10].Text.Trim().Replace("_UUID", "").Replace("&nbsp;", ""));
            //string Name = Convert.ToString(gvList.Rows[i].Cells[11].Text.Trim().Replace("Name", "").Replace("&nbsp;", ""));
            //string CurrAsofDate = Convert.ToString(gvList.Rows[i].Cells[12].Text.Trim().Replace("_CurrAsofDate", "").Replace("&nbsp;", ""));
            //string PrevAsOfDate = Convert.ToString(gvList.Rows[i].Cells[13].Text.Trim().Replace("_PrevAsOfDate", "").Replace("&nbsp;", ""));
            //string Type = Convert.ToString(gvList.Rows[i].Cells[14].Text.Trim().Replace("_Type", "").Replace("&nbsp;", ""));

            string PerformanceId1 = Convert.ToString(gvList.Rows[i].Cells[6].Text.Trim().Replace("_CurrAsofDateUUID", "").Replace("&nbsp;", ""));
            string PerformanceId2 = Convert.ToString(gvList.Rows[i].Cells[8].Text.Trim().Replace("_PrevAsOfDateUUID", "").Replace("&nbsp;", ""));

            TextBox txtPerformance1 = (TextBox)gvList.Rows[i].FindControl("txtPerformance1");
            TextBox txtPerformance2 = (TextBox)gvList.Rows[i].FindControl("txtPerformance2");

            TextBox txtPerfHidden1 = (TextBox)gvList.Rows[i].FindControl("txtPerfHidden1");
            TextBox txtPerfHidden2 = (TextBox)gvList.Rows[i].FindControl("txtPerfHidden2");

            //sas_publicperformance objPerformance = new sas_publicperformance();
            Entity objPerformance = new Entity("sas_publicperformance");

            try
            {


                if (_Header == "3")
                {
                    if (txtPerformance1.Text != "")
                    {
                        if (txtPerformance1.Text != txtPerfHidden1.Text)
                        {
                            // objPerformance.sas_publicperformanceid = new Key();
                            // objPerformance.sas_publicperformanceid.Value = new Guid(Convert.ToString(PerformanceId1));

                            objPerformance["sas_publicperformanceid"] = new Guid(Convert.ToString(PerformanceId1));

                            // objPerformance.sas_performance = new CrmDecimal();
                            // objPerformance.sas_performance.Value = Convert.ToDecimal(txtPerformance1.Text);

                            objPerformance["sas_performance"] = Convert.ToDecimal(txtPerformance1.Text);

                            service.Update(objPerformance);
                            intResult++;
                        }
                    }
                    if (txtPerformance2.Text != "")
                    {
                        if (txtPerformance2.Text != txtPerfHidden2.Text)
                        {
                            // objPerformance.sas_publicperformanceid = new Key();
                            // objPerformance.sas_publicperformanceid.Value = new Guid(Convert.ToString(PerformanceId2));

                            objPerformance["sas_publicperformanceid"] = new Guid(Convert.ToString(PerformanceId2));

                            // objPerformance.sas_performance = new CrmDecimal();
                            // objPerformance.sas_performance.Value = Convert.ToDecimal(txtPerformance2.Text);

                            objPerformance["sas_performance"] = Convert.ToDecimal(txtPerformance2.Text);

                            service.Update(objPerformance);
                            intResult++;
                        }
                    }
                }
            }
            //catch (System.Web.Services.Protocols.SoapException exc)
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
            {
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
                lblMessage.Text = strDescription;
            }
            catch (Exception exc)
            {
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                lblMessage.Text = strDescription;
            }
        }
        if (intResult > 0)
            lblMessage.Text = lblMessage.Text + "<br/>Records updated successfully";// for CLIENT SPECIFIC UPDATES";
        else
            lblMessage.Text = lblMessage.Text + "<br/> No Records updated.";

        BindGridView("'" + ddlHousehold.SelectedValue + "'");
    }
    protected void Button1_Click(object sender, EventArgs e)
    {

    }
    protected void btnRefresh_Click(object sender, EventArgs e)
    {
        BindGridView("'" + ddlHousehold.SelectedValue + "'");

        lblMessage.Text = "New Performance added Successfully";
    }
    protected void gvList_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            System.Web.UI.WebControls.Label lblLockDate1 = new System.Web.UI.WebControls.Label();
            lblLockDate1 = (System.Web.UI.WebControls.Label)e.Row.FindControl("lblLockDate1");

            System.Web.UI.WebControls.Label lblLockDate2 = new System.Web.UI.WebControls.Label();
            lblLockDate2 = (System.Web.UI.WebControls.Label)e.Row.FindControl("lblLockDate2");

            //HtmlAnchor lnkAdd1 = (HtmlAnchor)e.Row.FindControl("lnkAdd1");
            //HtmlAnchor lnkAdd2 = (HtmlAnchor)e.Row.FindControl("lnkAdd2");

            //HtmlAnchor lnkAdd1 = (HtmlAnchor)e.Row.FindControl("lnkAdd1");
            //HtmlAnchor lnkAdd2 = (HtmlAnchor)e.Row.FindControl("lnkAdd2");

            LinkButton lnkAdd1 = (LinkButton)e.Row.FindControl("lnkAdd1");
            LinkButton lnkAdd2 = (LinkButton)e.Row.FindControl("lnkAdd2");

            TextBox txtPerformance1 = (TextBox)e.Row.FindControl("txtPerformance1");
            TextBox txtPerformance2 = (TextBox)e.Row.FindControl("txtPerformance2");

            TextBox txtPerfHidden1 = (TextBox)e.Row.FindControl("txtPerfHidden1");
            TextBox txtPerfHidden2 = (TextBox)e.Row.FindControl("txtPerfHidden2");

            string _Headerflg = DataBinder.Eval(e.Row.DataItem, "_Headerflg").ToString();
            string _Performance1 = DataBinder.Eval(e.Row.DataItem, "CurrAsofDate").ToString();
            string _Performance2 = DataBinder.Eval(e.Row.DataItem, "PrevAsOfDate").ToString();

            string UUID = DataBinder.Eval(e.Row.DataItem, "_UUID").ToString();
            string Name = DataBinder.Eval(e.Row.DataItem, "Name").ToString().Replace("&", "%26").Replace("'", "%27");
            string CurrAsofDate = DataBinder.Eval(e.Row.DataItem, "_CurrAsofDate").ToString();
            string PrevAsOfDate = DataBinder.Eval(e.Row.DataItem, "_PrevAsOfDate").ToString();
            string Type = DataBinder.Eval(e.Row.DataItem, "_Type").ToString();

            string _CurrentFlg = DataBinder.Eval(e.Row.DataItem, "_CurrentFlg").ToString();
            string _PreviousFlg = DataBinder.Eval(e.Row.DataItem, "_PreviousFlg").ToString();
            string strLockDate = DataBinder.Eval(e.Row.DataItem, "_PerfLockDate").ToString();

            if (_CurrentFlg == "0")
            {
                txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = true;
                txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.Gray;
            }
            if (_PreviousFlg == "0")
            {
                txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = true;
                txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.Gray;
            }

            if (_Headerflg == "1")
            {
                e.Row.Font.Bold = true;
                e.Row.ForeColor = System.Drawing.Color.Orange;
                e.Row.Cells[0].CssClass = "CellTitle";
                e.Row.Cells[1].CssClass = "CellTitle";
                e.Row.Cells[2].CssClass = "CellTitle";
                e.Row.Cells[3].CssClass = "CellTitle";
                e.Row.Cells[4].CssClass = "CellTitle";
                e.Row.Height = Unit.Pixel(30);

                lnkAdd1.Visible = false;
                lnkAdd2.Visible = false;
                txtPerformance1.Visible = false;
                txtPerformance2.Visible = false;
            }
            else if (_Headerflg == "2")
            {
                e.Row.Font.Bold = true;
                e.Row.Cells[0].CssClass = "CellHeader";
                e.Row.Cells[1].CssClass = "CellHeader";
                e.Row.Cells[2].CssClass = "CellHeader";
                e.Row.Cells[3].CssClass = "CellHeader";
                e.Row.Cells[4].CssClass = "CellHeader";

                lnkAdd1.Visible = false;
                lnkAdd2.Visible = false;
                txtPerformance1.Visible = false;
                txtPerformance2.Visible = false;

                e.Row.Cells[4].Text = _Performance1;
                e.Row.Cells[3].Text = _Performance2;
            }
            else if (_Headerflg == "3")
            {
                e.Row.Cells[0].CssClass = "CellHeader";
                e.Row.Cells[1].CssClass = "CellHeader";
                e.Row.Cells[2].CssClass = "CellHeader";
                e.Row.Cells[3].CssClass = "CellHeader";
                e.Row.Cells[4].CssClass = "CellHeader";

                if (_Performance1 == "")
                {
                    if (_CurrentFlg == "1")
                    {
                        lnkAdd1.Visible = true;
                        txtPerformance1.Visible = false;
                        // lnkAdd1.Attributes["onclick"] = "OpenChild('" + UUID + "','" + Type + "','" + Name + "','" + CurrAsofDate + "');";//commented on 10/04/2018
                    }
                    else
                    {
                        lnkAdd1.Visible = false;
                        if (strLockDate != string.Empty)
                        {
                            lblLockDate1.Text = "Locked " + strLockDate;
                            lblLockDate1.Font.Name = "Verdana";
                            lblLockDate1.Font.Size = FontUnit.XSmall;

                            txtPerformance1.Visible = false;
                            //txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                            //txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                        }
                    }
                }
                else
                {
                    lnkAdd1.Visible = false;
                    txtPerformance1.Visible = true;
                    //   txtPerformance1.ReadOnly = false;
                    //  txtPerformance1.BackColor = System.Drawing.Color.White;
                    //txtPerformance1.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance1));
                    txtPerformance1.Text = Convert.ToString(_Performance1);
                    txtPerformance1.Attributes["onkeyup"] = "ChangeColor('" + txtPerformance1.ClientID + "')";

                    //txtPerfHidden1.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance1));
                    txtPerfHidden1.Text = Convert.ToString(_Performance1);


                    if (_CurrentFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                    if (_PreviousFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                }

                if (_Performance2 == "")
                {
                    if (_PreviousFlg == "1")
                    {
                        lnkAdd2.Visible = true;
                        txtPerformance2.Visible = false;
                        // lnkAdd2.Attributes["onclick"] = "OpenChild('" + UUID + "','" + Type + "','" + Name + "','" + PrevAsOfDate + "');";//commented on 10/04/2018
                    }
                    else
                    {
                        lnkAdd2.Visible = false;
                        if (strLockDate != string.Empty)
                        {
                            lblLockDate2.Text = "Locked " + strLockDate;
                            lblLockDate2.Font.Name = "Verdana";
                            lblLockDate2.Font.Size = FontUnit.XSmall;

                            txtPerformance2.Visible = false;
                            // txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                            // txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                        }
                    }
                }
                else
                {
                    lnkAdd2.Visible = false;
                    txtPerformance2.Visible = true;
                    //  txtPerformance2.ReadOnly = false;
                    //   txtPerformance2.BackColor = System.Drawing.Color.White;
                    //txtPerformance2.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance2));
                    txtPerformance2.Text = Convert.ToString(_Performance2);
                    txtPerformance2.Attributes["onkeyup"] = "ChangeColor('" + txtPerformance2.ClientID + "')";

                    //txtPerfHidden2.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance2));
                    txtPerfHidden2.Text = Convert.ToString(_Performance2);

                    if (_CurrentFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                    if (_PreviousFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                }
            }
        }
    }


    public void GridView()
    {
        foreach (GridViewRow grrow in gvList.Rows)
        {
            //TextBox Marketvalue = (TextBox)grrow.Cells[1].FindControl("txtMarketvalue");
            //CheckBox chkcolor = (CheckBox)grrow.Cells[8].FindControl("chkcolor");
            //string cell6 = grrow.Cells[6].Text;
            //if (Marketvalue.Text != cell6)
            //{
            //    //  grow.BackColor = System.Drawing.Color.FromName("#FCFB9C");
            //    Marketvalue.BackColor = System.Drawing.Color.FromName("#FCFB9C");

            //    // chkcolor.Checked = true;
            //}


            System.Web.UI.WebControls.Label lblLockDate1 = new System.Web.UI.WebControls.Label();
            lblLockDate1 = (System.Web.UI.WebControls.Label)grrow.FindControl("lblLockDate1");

            System.Web.UI.WebControls.Label lblLockDate2 = new System.Web.UI.WebControls.Label();
            lblLockDate2 = (System.Web.UI.WebControls.Label)grrow.FindControl("lblLockDate2");

            //HtmlAnchor lnkAdd1 = (HtmlAnchor)e.Row.FindControl("lnkAdd1");
            //HtmlAnchor lnkAdd2 = (HtmlAnchor)e.Row.FindControl("lnkAdd2");

            //HtmlAnchor lnkAdd1 = (HtmlAnchor)e.Row.FindControl("lnkAdd1");
            //HtmlAnchor lnkAdd2 = (HtmlAnchor)e.Row.FindControl("lnkAdd2");

            LinkButton lnkAdd1 = (LinkButton)grrow.FindControl("lnkAdd1");
            LinkButton lnkAdd2 = (LinkButton)grrow.FindControl("lnkAdd2");

            TextBox txtPerformance1 = (TextBox)grrow.FindControl("txtPerformance1");
            TextBox txtPerformance2 = (TextBox)grrow.FindControl("txtPerformance2");

            TextBox txtPerfHidden1 = (TextBox)grrow.FindControl("txtPerfHidden1");
            TextBox txtPerfHidden2 = (TextBox)grrow.FindControl("txtPerfHidden2");

            //string _Headerflg = DataBinder.Eval(grrow.DataItem, "_Headerflg").ToString();
            //string _Performance1 = DataBinder.Eval(grrow.DataItem, "CurrAsofDate").ToString();
            //string _Performance2 = DataBinder.Eval(grrow.DataItem, "PrevAsOfDate").ToString();


            string _Headerflg = grrow.Cells[9].Text;
            string _Performance1 = grrow.Cells[5].Text;
            string _Performance2 = grrow.Cells[7].Text;

            //string UUID = DataBinder.Eval(grrow.DataItem, "_UUID").ToString();
            //string Name = DataBinder.Eval(grrow.DataItem, "Name").ToString().Replace("&", "%26").Replace("'", "%27");
            //string CurrAsofDate = DataBinder.Eval(grrow.DataItem, "_CurrAsofDate").ToString();
            //string PrevAsOfDate = DataBinder.Eval(grrow.DataItem, "_PrevAsOfDate").ToString();
            // string Type = DataBinder.Eval(grrow.DataItem, "_Type").ToString();

            string UUID = grrow.Cells[10].Text;
            string Name = grrow.Cells[0].Text.Replace("&", "%26").Replace("'", "%27");
            string CurrAsofDate = grrow.Cells[12].Text;
            string PrevAsOfDate = grrow.Cells[13].Text;
            string Type = grrow.Cells[14].Text;

            //string _CurrentFlg = DataBinder.Eval(grrow.DataItem, "_CurrentFlg").ToString();
            //string _PreviousFlg = DataBinder.Eval(grrow.DataItem, "_PreviousFlg").ToString();
            //string strLockDate = DataBinder.Eval(grrow.DataItem, "_PerfLockDate").ToString();


            string _CurrentFlg = grrow.Cells[15].Text;
            string _PreviousFlg = grrow.Cells[16].Text;
            string strLockDate = grrow.Cells[17].Text;

            if (_CurrentFlg == "0")
            {
                txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = true;
                txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.Gray;
            }
            if (_PreviousFlg == "0")
            {
                txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = true;
                txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.Gray;
            }

            if (_Headerflg == "1")
            {
                grrow.Font.Bold = true;
                grrow.ForeColor = System.Drawing.Color.Orange;
                grrow.Cells[0].CssClass = "CellTitle";
                grrow.Cells[1].CssClass = "CellTitle";
                grrow.Cells[2].CssClass = "CellTitle";
                grrow.Cells[3].CssClass = "CellTitle";
                grrow.Cells[4].CssClass = "CellTitle";
                grrow.Height = Unit.Pixel(30);

                lnkAdd1.Visible = false;
                lnkAdd2.Visible = false;
                txtPerformance1.Visible = false;
                txtPerformance2.Visible = false;
            }
            else if (_Headerflg == "2")
            {
                grrow.Font.Bold = true;
                grrow.Cells[0].CssClass = "CellHeader";
                grrow.Cells[1].CssClass = "CellHeader";
                grrow.Cells[2].CssClass = "CellHeader";
                grrow.Cells[3].CssClass = "CellHeader";
                grrow.Cells[4].CssClass = "CellHeader";

                lnkAdd1.Visible = false;
                lnkAdd2.Visible = false;
                txtPerformance1.Visible = false;
                txtPerformance2.Visible = false;

                grrow.Cells[4].Text = _Performance1;
                grrow.Cells[3].Text = _Performance2;
            }
            else if (_Headerflg == "3")
            {
                grrow.Cells[0].CssClass = "CellHeader";
                grrow.Cells[1].CssClass = "CellHeader";
                grrow.Cells[2].CssClass = "CellHeader";
                grrow.Cells[3].CssClass = "CellHeader";
                grrow.Cells[4].CssClass = "CellHeader";

                if (_Performance1 == "")
                {
                    if (_CurrentFlg == "1")
                    {
                        lnkAdd1.Visible = true;
                        txtPerformance1.Visible = false;
                        // lnkAdd1.Attributes["onclick"] = "OpenChild('" + UUID + "','" + Type + "','" + Name + "','" + CurrAsofDate + "');";//commented on 10/04/2018
                    }
                    else
                    {
                        lnkAdd1.Visible = false;
                        if (strLockDate != string.Empty)
                        {
                            lblLockDate1.Text = "Locked " + strLockDate;
                            lblLockDate1.Font.Name = "Verdana";
                            lblLockDate1.Font.Size = FontUnit.XSmall;

                            txtPerformance1.Visible = false;
                            //txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                            //txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                        }
                    }
                }
                else
                {
                    lnkAdd1.Visible = false;
                    txtPerformance1.Visible = true;
                    //   txtPerformance1.ReadOnly = false;
                    //  txtPerformance1.BackColor = System.Drawing.Color.White;
                    //txtPerformance1.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance1));
                    txtPerformance1.Text = Convert.ToString(_Performance1);
                    txtPerformance1.Attributes["onkeyup"] = "ChangeColor('" + txtPerformance1.ClientID + "')";

                    //txtPerfHidden1.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance1));
                    txtPerfHidden1.Text = Convert.ToString(_Performance1);


                    if (_CurrentFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                    if (_PreviousFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                }

                if (_Performance2 == "")
                {
                    if (_PreviousFlg == "1")
                    {
                        lnkAdd2.Visible = true;
                        txtPerformance2.Visible = false;
                        // lnkAdd2.Attributes["onclick"] = "OpenChild('" + UUID + "','" + Type + "','" + Name + "','" + PrevAsOfDate + "');";//commented on 10/04/2018
                    }
                    else
                    {
                        lnkAdd2.Visible = false;
                        if (strLockDate != string.Empty)
                        {
                            lblLockDate2.Text = "Locked " + strLockDate;
                            lblLockDate2.Font.Name = "Verdana";
                            lblLockDate2.Font.Size = FontUnit.XSmall;

                            txtPerformance2.Visible = false;
                            // txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                            // txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                        }
                    }
                }
                else
                {
                    lnkAdd2.Visible = false;
                    txtPerformance2.Visible = true;
                    //  txtPerformance2.ReadOnly = false;
                    //   txtPerformance2.BackColor = System.Drawing.Color.White;
                    //txtPerformance2.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance2));
                    txtPerformance2.Text = Convert.ToString(_Performance2);
                    txtPerformance2.Attributes["onkeyup"] = "ChangeColor('" + txtPerformance2.ClientID + "')";

                    //txtPerfHidden2.Text = string.Format("{0:#,###0.00;-#,###0.00}", Convert.ToDecimal(_Performance2));
                    txtPerfHidden2.Text = Convert.ToString(_Performance2);

                    if (_CurrentFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                    if (_PreviousFlg == "1")
                    {
                        txtPerformance1.ReadOnly = txtPerformance2.ReadOnly = false;
                        txtPerformance1.BackColor = txtPerformance2.BackColor = System.Drawing.Color.White;
                    }
                }
            }

        }
    }

    // /// <summary>
    // /// Set up the CRM Service.
    // /// </summary>
    // /// <param name="organizationName">My Organization</param>
    // /// <returns>CrmService configured with AD Authentication</returns>
    // public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    // {
    // // Get the CRM Users appointments
    // // Setup the Authentication Token
    // CrmAuthenticationToken token = new CrmAuthenticationToken();
    // token.AuthenticationType = 0; // Use Active Directory authentication.
    // token.OrganizationName = organizationName;
    // // string username = WindowsIdentity.GetCurrent().Name;

    // if (CallerId != "")
    // token.CallerId = new Guid(CallerId);

    // CrmService service = new CrmService();

    // if (crmServerUrl != null &&
    // crmServerUrl.Length > 0)
    // {
    // UriBuilder builder = new UriBuilder(crmServerUrl);
    // builder.Path = "//MSCRMServices//2007//CrmService.asmx";
    // service.Url = builder.Uri.ToString();
    // }

    // service.CrmAuthenticationTokenValue = token;
    // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

    // return service;
    // }
    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        //  string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;

        //Changed Windows to - ADFS Claims Login 8_9_2019
      //  IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
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
        }
        else
        {
            return UserID = "568C5730-8228-E411-8DFC-0002A5443D86"; //"E3259341-8303-DE11-A38C-001D09665E8F";
        }
    }
    protected void btnCanceltop_Click(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        BindGridView("'" + ddlHousehold.SelectedValue + "'");
    }
    protected void gvList_RowCommand(object sender, GridViewCommandEventArgs e)
    {
        if (e.CommandName == "linkButton2")
        {
            //Determine the RowIndex of the Row whose LinkButton was clicked.
            int rowIndex = Convert.ToInt32(e.CommandArgument);

            //Reference the GridView Row.
            GridViewRow row = gvList.Rows[rowIndex];

            string Type = row.Cells[14].Text;
            ViewState["Type"] = Type;
            if (Type != "")
                lblPerfType.Text = Type;


            string Name = row.Cells[0].Text;
            ViewState["Name"] = Name;
            if (Name != "")
                lblName.Text = Name;



            string PervAsOfDate = row.Cells[13].Text;

            DateTime Date = Convert.ToDateTime(PervAsOfDate);

            var Date1 = Date.ToShortDateString();
            //Convert.ToDateTime(Request.QueryString["asofdate"]).ToString("MM/dd/yyyy");
            //DateTime Date = DateTime.ParseExact(PervAsOfDate, "M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture);
            lblDate.Text = Date1;

            ViewState["PervAsOfDate"] = lblDate.Text;
            string UUID = row.Cells[10].Text;
            ViewState["UUID"] = UUID;
            Label1.Text = "";
            txtPerformance.Text = "";
            performancepanel.Visible = true;
            performancepopup.Show();


        }

        else if (e.CommandName == "linkButton1")
        {
            //Determine the RowIndex of the Row whose LinkButton was clicked.
            int rowIndex = Convert.ToInt32(e.CommandArgument);

            //Reference the GridView Row.
            GridViewRow row = gvList.Rows[rowIndex];



            string Type = row.Cells[14].Text;
            ViewState["Type"] = Type;
            if (Type != "")
                lblPerfType.Text = Type;

            string Name = row.Cells[0].Text;
            ViewState["Name"] = Name;

            if (Name != "")
                lblName.Text = Name;

            string CurrAsOfDate = row.Cells[12].Text;

            DateTime Date = Convert.ToDateTime(CurrAsOfDate);

            var Date1 = Date.ToShortDateString();

            lblDate.Text = Date1;

            ViewState["CurrAsOfDate"] = lblDate.Text;

            string UUID = row.Cells[10].Text;

            ViewState["UUID"] = UUID;
            Label1.Text = "";
            txtPerformance.Text = "";
            performancepanel.Visible = true;
            performancepopup.Show();
        }
    }


    public void InsertIntoCRM()
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        Type tp = this.GetType();

        // string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        // string orgName = "GreshamPartners";
        //CrmService service = null;

        IOrganizationService service = null;

        lblMessage.Text = "";

        string UserId = GetcurrentUser();
        lg.AddinLogFile(Session["Filename"].ToString(), "UserId" + UserId);
        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = GetCrmService();
            strDescription = "Crm Service starts successfully";
            lg.AddinLogFile(Session["Filename"].ToString(), strDescription);
        }
        // catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
            lg.AddinLogFile(Session["Filename"].ToString(), "Error" + strDescription);
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
            lg.AddinLogFile(Session["Filename"].ToString(), "Error1" + strDescription);
        }


        string UUID = Convert.ToString(ViewState["UUID"]);


        string Type = Convert.ToString(ViewState["Type"]);


        string Name = Convert.ToString(ViewState["Name"]);


        // string Type = Convert.ToString(ViewState["Type"]);
        lg.AddinLogFile(Session["Filename"].ToString(), "UUID" + UUID);
        lg.AddinLogFile(Session["Filename"].ToString(), "Type" + Type);
        lg.AddinLogFile(Session["Filename"].ToString(), "Name" + Name);

        // service.PreAuthenticate = true;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        if (UUID != "" && UUID != null)

        //if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["uuid"]) != null)
        {

            if (txtPerformance.Text == "")
            {
                //sb.Append("\n<script type=text/javascript>\n");
                //sb.Append("\n alert('Please enter value in performance.');");
                //sb.Append("</script>");
                //ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                performancepanel.Visible = true;
                performancepopup.Show();
                Label1.Visible = true;
                Label1.Text = "Please enter value in performance";
                return;
            }

            if (!IsDuplicateExists())
            {
                lg.AddinLogFile(Session["Filename"].ToString(), "Inside IF"  );

                //sas_publicperformance objPerformance = new sas_publicperformance();
                Entity objPerformance = new Entity("sas_publicperformance");

                // objPerformance.sas_performance = new CrmDecimal();
                // objPerformance.sas_performance.Value = Convert.ToDecimal(txtPerformance.Text.Trim());

                objPerformance["sas_performance"] = Convert.ToDecimal(txtPerformance.Text.Trim());

                // if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["type"]) == "FUND")
                if (UUID != "" && Type.ToUpper() == "FUND")
                {
                    // objPerformance.ssi_fundid = new Lookup();
                    // objPerformance.ssi_fundid.type = EntityName.ssi_fund.ToString();
                    // objPerformance.ssi_fundid.Value = new Guid(Convert.ToString(Request.QueryString["uuid"]));

                    objPerformance["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(UUID));
                }

                //if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["type"]) == "ACCOUNT")

                if (UUID != "" && Type.ToUpper() == "ACCOUNT")
                {
                    // objPerformance.ssi_clientaccountid = new Lookup();
                    // objPerformance.ssi_clientaccountid.type = EntityName.ssi_account.ToString();
                    // objPerformance.ssi_clientaccountid.Value = new Guid(Convert.ToString(Request.QueryString["uuid"]));

                    objPerformance["ssi_clientaccountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(UUID));
                }

                // if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["type"]) == "BENCHMARK")
                if (UUID != "" && Type.ToUpper() == "BENCHMARK")
                {
                    // objPerformance.sas_performanceid = new Lookup();
                    // objPerformance.sas_performanceid.type = EntityName.sas_benchmark.ToString();
                    // objPerformance.sas_performanceid.Value = new Guid(Convert.ToString(Request.QueryString["uuid"]));

                    objPerformance["sas_performanceid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(UUID));
                }

                // objPerformance.sas_enddate = new CrmDateTime();
                // objPerformance.sas_enddate.Value = lblDate.Text.Trim();

                objPerformance["sas_enddate"] = Convert.ToDateTime(lblDate.Text.Trim());

                // objPerformance.sas_startdate = new CrmDateTime();
                // objPerformance.sas_startdate.Value = lblDate.Text.Split('/')[0] + "/01/" + lblDate.Text.Split('/')[2];

                objPerformance["sas_startdate"] = Convert.ToDateTime(lblDate.Text.Split('/')[0] + "/01/" + lblDate.Text.Split('/')[2]);

                try
                {
                    lg.AddinLogFile(Session["Filename"].ToString(), "TRY Create" );
                    service.Create(objPerformance);
                    lg.AddinLogFile(Session["Filename"].ToString(), "Created");
                }
                catch (Exception ex)
                {
                    lg.AddinLogFile(Session["Filename"].ToString(), "Error" + ex.Message.ToString());
                    performancepanel.Visible = true;
                    performancepopup.Show();
                    Label1.Visible = true;
                    Label1.Text = ex.ToString();
                    return;
                }
                successcount++;
                lg.AddinLogFile(Session["Filename"].ToString(), "successcount" + successcount);
                // Response.Write(successcount.ToString());


                if (successcount > 0)
                {
                    //Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "close", "<script type='text/javascript'>ReturnToParent('true');</script>");

                    /// BindGridView("'" + ddlHousehold.SelectedValue + "'");
                    performancepanel.Visible = false;
                    performancepopup.Hide();
                    BindGridView("'" + ddlHousehold.SelectedValue + "'");
                    lblMessage.Text = "New Performance added Successfully";

                    #region postback for Session 
                    //Session["ddlHousehold"] = ddlHousehold.SelectedValue.ToString();
                    //Session["ddlAssociateOps"] = ddlAssociateOps.SelectedValue.ToString();
                    //Session["txtAsOfDate"] = txtAsOfDate.Text;
                    //Session["lblMessage"] = lblMessage.Text;
                    //// Session["txtAsOfDate"] = "";

                    //Response.Redirect(Request.RawUrl);

                    #endregion
                    //lblMessage.Text = "New Performance added Successfully";
                }
            }
            else
            {
                lg.AddinLogFile(Session["Filename"].ToString(), "Inside Ielse"  );
                lblMessage.Visible = true;
                if (lblMessage.Text == "")
                {
                    performancepanel.Visible = false;
                    performancepopup.Hide();
                    //lblMessage.Text = "Record already exists.";
                     Label1.Text = "Record already exists.";
                }
            }
        }
    }


    private bool IsDuplicateExists()
    {
        bool status = false;
        try
        {
            object UUid = "null";

            string UUID = Convert.ToString(ViewState["UUID"]);


            string Type = Convert.ToString(ViewState["Type"]);


            string Name = Convert.ToString(ViewState["Name"]);


            // if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "")
            if (UUID != "")
            {
               // UUid = new Guid(Convert.ToString(Request.QueryString["uuid"]));
                UUid = new Guid(UUID);
            }

            string EndDate = lblDate.Text.Trim();

            string StartDate = lblDate.Text.Split('/')[0] + "/01/" + lblDate.Text.Split('/')[2];

            string query = "exec SP_S_PERFORMANCE_CHECK @UUID='" + UUid + "'"
                                                         + ",@StartDT='" + StartDate + "'"
                                                         + ",@EndDt = '" + EndDate + "'";

            DataSet ds = clsdb.getDataSet(query);

            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    status = Convert.ToBoolean(ds.Tables[0].Rows[0]["ReturnNmb"]);
                }
            }
        }
        catch (Exception ex)
        {
            performancepanel.Visible = false;
            performancepopup.Hide();
            Label1.Text = "Error occured while checking duplicate - " + ex.Message;
            status = true;
        }
        return status;
    }

    protected void btnCancel_Click(object sender, EventArgs e)
    {
        performancepopup.Hide();
        performancepanel.Visible = false;
        //Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "close", "<script type='text/javascript'>Postback();</script>");

        ////  Session["Assosiate"] = "";
        //Session["ddlHousehold"] = ddlHousehold.SelectedValue.ToString();
        //Session["ddlAssociateOps"] = ddlAssociateOps.SelectedValue.ToString();
        //Session["txtAsOfDate"] = txtAsOfDate.Text;
        //// Session["txtAsOfDate"] = "";

       // Response.Redirect(Request.RawUrl);

        BindGridView("'" + ddlHousehold.SelectedValue + "'");
        //Response.Redirect("~");
        // GridView();
    }
    protected void txtPerformance_TextChanged(object sender, EventArgs e)
    {
        Label1.Text = "";
    }
    protected void btnsubmitpopup_Click(object sender, EventArgs e)
    {
        
        InsertIntoCRM();


    }
    public IOrganizationService GetCrmService()
    {
        Microsoft.Xrm.Sdk.IOrganizationService service;
        try
        {


            ClientCredentials Credentials = new ClientCredentials();
            // Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

            //  Credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;

            //string str = ((WindowsIdentity)HttpContext.Current.User.Identity).Name;
            string str = string.Empty;
            if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
            {
                str = "corp\\gbhagia";
            }
            else
            {
                IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
                str = claimsIdentity.Name;

            }
            lg.AddinLogFile(Session["Filename"].ToString(), "User " + str);
            string UserID = string.Empty;
            string sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + str + "'";
            DB clsDB = new DB();
            DataSet lodataset = clsDB.getDataSet(sqlstr);
            if (lodataset.Tables[0].Rows.Count > 0)
            {
                UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);

                // Response.Write("UserID:" + UserID);
                //return UserID = "DFCE21B1-B81E-E211-A2B7-0002A5443D86";
            }
            lg.AddinLogFile(Session["Filename"].ToString(), "UserID from DB  " + UserID);

            // *** for specific user credential *********  //
            Credentials.UserName.UserName = "corp\\crmadmin";
            Credentials.UserName.Password = "51ngl3malt_51ngl3malt";




            // Credentials.UserName.UserName = "corp\\crmadmin";
            // Credentials.UserName.Password = "W!gmxF26ggw]";

            //This URL needs to be updated to match the servername and Organization for the environment.
            //  Uri OrganizationUri = new Uri("http://gp-crm2016/GreshamPartners/XRMServices/2011/Organization.svc");
            Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
            Uri HomeRealmUri = null;



            //OrganizationServiceProxy serviceProxy; 
            System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy serviceProxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null);

            // This statement is required to enable early-bound type support.
            serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

            service = (Microsoft.Xrm.Sdk.IOrganizationService)serviceProxy;
            Guid _UserID = new Guid(UserID);
            lg.AddinLogFile(Session["Filename"].ToString(), "Connected to CRM  " + service.ToString());
            //serviceProxy.CallerId = _UserID;
            lg.AddinLogFile(Session["Filename"].ToString(), "CallerID set  " + _UserID);
            return service;

        }
        catch (Exception ex)
        {
            service = null;
        }
        return service;
    }
}
