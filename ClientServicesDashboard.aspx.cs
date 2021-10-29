using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class ClientServicesDashboard : System.Web.UI.Page
{
    DataSet ds = new DataSet();
    public string ConnString = "";
    string query = string.Empty;
    public String _dbErrorMsg;
    DB cls = new DB();
    //string HouseHoldID = "800a713d-6a15-de11-8391-001d09665e8f";
    //string HouseholdName = "Baird Family";

    string HouseHoldID = string.Empty;
    string HouseholdName = string.Empty;
    // string HouseHoldID = string.Empty;
    string ddltext = string.Empty;
    // string AsOFDate = string.Empty;

    string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            HouseHoldID = Convert.ToString(Request.QueryString["Id"]);
            HouseHoldID = "440B713D-6A15-DE11-8391-001D09665E8F"; ////Response.Write(HouseHoldID);
            string sqlstr = "select top 1 Name from Account where accountid = '" + HouseHoldID + "'";
            DB clsDB = new DB();
            DataSet lodataset = clsDB.getDataSet(sqlstr);
            if (lodataset.Tables[0].Rows.Count > 0)
            {
                HouseholdName = Convert.ToString(lodataset.Tables[0].Rows[0]["Name"]);
            }


            // HouseholdName = Convert.ToString(Request.QueryString["Name"]);
            ViewState["HouseholdName"] = HouseholdName;

            //Response.Write(HouseholdName);
            if (HouseHoldID != null)
            {
                if (HouseholdName != "")
                    this.Title = HouseholdName;
                BindGridView_LE();
                BindGridView_Recommendation();
                BindGridView_CallReports();
                BindGridView_Task();
                BindGridView_Email();
                #region for AsOFDate
                string AsOFDate = AsOfDate(HouseHoldID);
                //string AsOFDate = "31-DEC-2017";
                ViewState["AsOFDate"] = AsOFDate;
                if (AsOFDate != "")
                    lblAsofdate.Text = "<b>As Of:</b>" + AsOFDate;
                #endregion
                Bindddl();
                DataTable dt = (DataTable)ViewState["ddlBindAllGrp"];
                if (dt.Rows.Count > 0)
                    ddltext = ddlList.SelectedItem.Text.ToString();
                else
                    ddltext = "";
                BindGridView_Allocation(ddltext, HouseholdName, AsOFDate);
                BindGridView_SPMM();
                lblHoushold.Text = HouseholdName + " Dashboard";
                //lblAsOfDate.Text = "03/31/2018";
            }
            else
            {
                lblerror.Text = "HouseHoldID Not Found";
            }
        }
    }

    public SortDirection dir
    {

        get
        {

            if (ViewState["dirState"] == null)
            {
                ViewState["dirState"] = SortDirection.Ascending;
            }

            return (SortDirection)ViewState["dirState"];
        }

        set
        {
            ViewState["dirState"] = value;
        }

    }

    public void BindGridView_Recommendation()
    {
        query = "EXEC SP_S_RECOMMENDATIONS_DashBoard @HouseHoldID='" + HouseHoldID + "'";
        ds = cls.getDataSet(query);
        //ds.Tables[0].DefaultView.Sort = "Performance Date desc";
        ViewState["dtRecomm"] = ds.Tables[0];

        if (ds.Tables[0].Rows.Count > 0)
        {
            gvRecommendation.DataSource = ds.Tables[0];
            gvRecommendation.DataBind();

            lblRecommTimeStamp.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();
        }
        else
        {
            gvRecommendation.DataSource = null;
            gvRecommendation.DataBind();
            lblalertRecommendation.Text = "No Data Found";
            lblRecommTimeStamp.Text = "";

            //lbExporttoExcel.Visible = false;

        }
    }

    public void BindGridView_Allocation(string selectedItem, string Houshholdname, string asofdate)
    {
        // string AsOFDate = "31-DEC-2017";

        query = "exec GreshamPartners_MSCRM.dbo.SP_R_INVESTMENT_OBJECTIVE_CHART_EXCEL_SMA_NEW_GA_BASEDATA @HouseholdName = '" + Houshholdname + "',@AsofDate = '" + asofdate + "', @GreshamAdvisedFlagTxt = 'TIA' ,@AllocGroupName = '" + selectedItem + "',@IncludeTotFlg = 1";
        SqlParameter[] param2 = new SqlParameter[5];
        param2[0] = new SqlParameter("@HouseholdName", SqlDbType.VarChar);
        param2[0].Value = Houshholdname;

        param2[1] = new SqlParameter("@AsofDate", SqlDbType.VarChar);
        param2[1].Value = asofdate;

        param2[2] = new SqlParameter("@GreshamAdvisedFlagTxt", SqlDbType.VarChar);
        param2[2].Value = "TIA";

        param2[3] = new SqlParameter("@AllocGroupName", SqlDbType.VarChar);
        param2[3].Value = selectedItem;

        param2[4] = new SqlParameter("@IncludeTotFlg", SqlDbType.VarChar);
        param2[4].Value = "1";


        // query = "exec GreshamPartners_MSCRM.dbo.SP_R_INVESTMENT_OBJECTIVE_CHART_EXCEL_SMA_NEW_GA_BASEDATA @HouseholdName='" + HouseholdName
        //+ "',@Start_DT=" + startdte + ""
        //+ ",@End_DT = " + enddte + " ";

        ds = cls.ExecuteSPOutParameter("GreshamPartners_MSCRM.dbo.SP_R_INVESTMENT_OBJECTIVE_CHART_EXCEL_SMA_NEW_GA_BASEDATA ", 0, param2);
        //   ds = cls.getDataSet(query);
        ViewState["dtAllocation"] = ds.Tables[0];
        if (ds.Tables[0].Rows.Count > 0)
        {
            gvAllocation.DataSource = ds.Tables[0];
            gvAllocation.DataBind();
            UpdatePanelAllocation.Update();
            //   UpdatePanelAllocation.Update();
            //   lblTagetTimeFrame.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();

        }
        else
        {
            gvAllocation.DataSource = null;
            gvAllocation.DataBind();
            //LinkbtnAllocation.Visible = false;
            lblAlertAllocation.Text = "No Data Found";

            //    lblTagetTimeFrame.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();
        }

    }

    public void BindGridView_LE()
    {
        query = "EXEC SP_S_LEGALENTITY_DashBoard @HouseHoldID='" + HouseHoldID + "'";
        ds = cls.getDataSet(query);
        //ds.Tables[0].DefaultView.Sort = "Performance Date desc";
        ViewState["dtLE"] = ds.Tables[0];
        if (ds.Tables[0].Rows.Count > 0)
        {
            // gvLE.Columns[0].Visible = true;
            gvLE.DataSource = ds.Tables[0];
            gvLE.DataBind();
            // gvLE.Columns[0].Visible = false;
            lblLETimeStamp.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();
        }
        else
        {
            gvLE.DataSource = null;
            gvLE.DataBind();
            lblAlertLE.Text = "No Data Found";
            lblLETimeStamp.Text = "";
            //lbLE.Visible = false;
        }
    }

    public void BindGridView_SPMM()//New Function 
    {
        query = "EXEC SP_S_Transactions_DashBoard @HouseHoldID='" + HouseHoldID + "'";
        ds = cls.getDataSet(query);
        //ds.Tables[0].DefaultView.Sort = "Performance Date desc";
        ViewState["dtSales&Purchase"] = ds.Tables[0];
        ViewState["dtMoneyMarket"] = ds.Tables[1];
        if (ds.Tables[0].Rows.Count > 0)
        {
            // gvLE.Columns[0].Visible = true;
            gvSalesPurchase.DataSource = ds.Tables[0];
            gvSalesPurchase.DataBind();
            // gvLE.Columns[0].Visible = false;

            lblSales.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();

        }
        else
        {
            gvSalesPurchase.DataSource = null;
            gvSalesPurchase.DataBind();
            lblSP.Text = "No Data Found";
            lblSales.Text = "";
            //lbLE.Visible = false;
        }

        if (ds.Tables[1].Rows.Count > 0)
        {
            // gvLE.Columns[0].Visible = true;
            gvMoney.DataSource = ds.Tables[1];
            gvMoney.DataBind();

            lblMMTimeStamp.Text = ds.Tables[1].Rows[0]["TimeFrame"].ToString();
            // gvLE.Columns[0].Visible = false;
        }
        else
        {
            gvMoney.DataSource = null;
            gvMoney.DataBind();
            lblMM.Text = "No Data Found";
            //   lblMMTimeStamp.Text = "";
            //lbLE.Visible = false;
        }

    }

    public void BindGridView_Task()
    {
        query = "EXEC SP_S_Task_DashBoard @HouseHoldID='" + HouseHoldID + "'";
        ds = cls.getDataSet(query);
        //ds.Tables[0].DefaultView.Sort = "Performance Date desc";
        ViewState["dtTask"] = ds.Tables[0];

        if (ds.Tables[0].Rows.Count > 0)
        {
            gvTask.DataSource = ds.Tables[0];
            gvTask.DataBind();
            lblTaskTimeFrame.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();
        }
        else
        {
            gvTask.DataSource = null;
            gvTask.DataBind();
            lblAlertTask.Text = "No Data Found";
            //Linkbtntask.Visible = false;
            lblTaskTimeFrame.Text = "";
        }
    }
    public void BindGridView_Email()
    {
        query = "EXEC SP_S_Email_DashBoard @HouseHoldID='" + HouseHoldID + "'";
        ds = cls.getDataSet(query);
        //ds.Tables[0].DefaultView.Sort = "Performance Date desc";
        ViewState["dtEmail"] = ds.Tables[0];

        if (ds.Tables[0].Rows.Count > 0)
        {

            //gvEmail.Columns[5].Visible = true;
            gvEmail.DataSource = ds.Tables[0];
            gvEmail.DataBind();
            //gvEmail.Columns[5].Visible = true;

            lblEmailTimeFrame.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();
        }
        else
        {
            gvEmail.DataSource = null;
            gvEmail.DataBind();
            lblAlertTask.Text = "No Data Found";
            //Linkbtntask.Visible = false;
            lblEmailTimeFrame.Text = "";
        }
    }

    public void BindGridView_CallReports()
    {
        query = "EXEC SP_S_Activity_DashBoard @HouseHoldID='" + HouseHoldID + "'";
        ds = cls.getDataSet(query);
        ViewState["dtCallReports"] = ds.Tables[0];


        if (ds.Tables[0].Rows.Count > 0)
        {
            gvcallreports.DataSource = ds.Tables[0];
            gvcallreports.DataBind();
            lblActivityTimeFrame.Text = ds.Tables[0].Rows[0]["TimeFrame"].ToString();

        }
        else
        {
            gvcallreports.DataSource = null;
            gvcallreports.DataBind();
            lblAlertCallReport.Text = "No Data Found";
            lblActivityTimeFrame.Text = "";

            //LinkbtnCallReport.Visible = false;
        }
    }


    public void Bindddl()
    {
        // string ddlvalue = ddlList.SelectedValue.ToString();
        //  string AsOfDate = string.Empty;
        string AsOfDate = ViewState["AsOFDate"].ToString();
        //query = "EXEC sp_s_Get_OtherGroupName_Only @HouseHoldNameTxt='" + HouseholdName + "',@AllocationGroupFlg  = 1,@AsOfDate='" + AsOfDate + "'";
        //ds = cls.getDataSet(query);
        //ViewState["ddlBindAllGrp"] = ds.Tables[0];
        //if (ds.Tables[0].Rows.Count > 0)
        //{
        //    ddlList.DataTextField = "other2name";
        //    ddlList.DataValueField = "other2name";
        //    ddlList.DataSource = ds.Tables[0];
        //    ddlList.DataBind();

        //}

      //  query = "EXEC sp_s_Get_OtherGroupName_Only @HouseHoldNameTxt='" + HouseholdName + "',@AllocationGroupFlg  = 1,@AsOfDate='" + AsOfDate + "'";
        SqlParameter[] param2 = new SqlParameter[3];
        param2[0] = new SqlParameter("@HouseHoldNameTxt", SqlDbType.VarChar);
        param2[0].Value = HouseholdName;

        param2[1] = new SqlParameter("@AllocationGroupFlg", SqlDbType.VarChar);
        param2[1].Value = "1";

        param2[2] = new SqlParameter("@AsofDate", SqlDbType.VarChar);
        param2[2].Value = AsOfDate;

       

        // query = "exec GreshamPartners_MSCRM.dbo.SP_R_INVESTMENT_OBJECTIVE_CHART_EXCEL_SMA_NEW_GA_BASEDATA @HouseholdName='" + HouseholdName
        //+ "',@Start_DT=" + startdte + ""
        //+ ",@End_DT = " + enddte + " ";

        ds = cls.ExecuteSPOutParameter("sp_s_Get_OtherGroupName_Only", 0, param2);
        ViewState["ddlBindAllGrp"] = ds.Tables[0];
        if (ds.Tables[0].Rows.Count > 0)
        {
            ddlList.DataTextField = "other2name";
            ddlList.DataValueField = "other2name";
            ddlList.DataSource = ds.Tables[0];
            ddlList.DataBind();

        }

    }


    public string AsOfDate(string Housholdid)
    {
        string AsOfDate = string.Empty;
        query = "EXEC SP_S_HouseHoldList_DashBoard @HouseholdUUID='" + HouseHoldID + "'";

        ds = cls.getDataSet(query);
        if (ds.Tables[0].Rows.Count > 0)
        {
            AsOfDate = ds.Tables[0].Rows[0]["AsOfDate"].ToString();
        }
        return AsOfDate;
    }
    protected void gvRecommendation_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtRecomm"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvRecommendation.DataSource = sortedView;
        //gvRecommendation.Columns[9].Visible = true;
        gvRecommendation.DataBind();


        // UpdatePanel1.Update();
    }
    protected void gvcallreports_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtCallReports"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvcallreports.DataSource = sortedView;
        //gvRecommendation.Columns[9].Visible = true;
        gvcallreports.DataBind();
    }
    protected void gvLE_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtLE"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvLE.DataSource = sortedView;
        //gvRecommendation.Columns[9].Visible = true;
        gvLE.DataBind();
    }
    protected void gvAllocation_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtAllocation"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvAllocation.DataSource = sortedView;
        //gvRecommendation.Columns[9].Visible = true;
        gvAllocation.DataBind();
    }

    protected void gvEmail_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtEmail"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvEmail.DataSource = sortedView;        
        gvEmail.DataBind();
    }
    protected void gvTask_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtTask"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvTask.DataSource = sortedView;
        //gvRecommendation.Columns[9].Visible = true;
        gvTask.DataBind();
    }





    protected void lbExporttoExcel_Click(object sender, EventArgs e)//for Recommendation
    {

    }
    protected void lbLE_Click(object sender, EventArgs e)//for LE
    {

    }
    protected void LinkbtnCallReport_Click(object sender, EventArgs e)//for call reports
    {

    }
    protected void Linkbtntask_Click(object sender, EventArgs e)//for task
    {

    }
    protected void LinkbtnAllocation_Click(object sender, EventArgs e)//for Allocation
    {

    }
    protected void ddlList_SelectedIndexChanged(object sender, EventArgs e)
    {
        string ddlItem = ddlList.SelectedItem.Text.ToString();
        string AsOfDate = ViewState["AsOFDate"].ToString();
        string HouseholdName = ViewState["HouseholdName"].ToString();
        //ddlvalue = "";
        BindGridView_Allocation(ddlItem, HouseholdName, AsOfDate);
    }
    protected void gvSalesPurchase_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtSales&Purchase"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvTask.DataSource = sortedView;
        //gvRecommendation.Columns[9].Visible = true;
        gvTask.DataBind();
    }
    protected void gvMoney_Sorting(object sender, GridViewSortEventArgs e)
    {
        string sortingDirection = string.Empty;

        if (dir == SortDirection.Ascending)
        {

            dir = SortDirection.Descending;

            sortingDirection = "Desc";

        }

        else
        {

            dir = SortDirection.Ascending;

            sortingDirection = "Asc";

        }


        DataView sortedView = new DataView((DataTable)ViewState["dtMoneyMarket"]);

        sortedView.Sort = e.SortExpression + " " + sortingDirection;

        gvTask.DataSource = sortedView;
        //gvRecommendation.Columns[9].Visible = true;
        gvTask.DataBind();
    }
}