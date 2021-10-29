using Microsoft.IdentityModel.Claims;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class CleanRRG : System.Web.UI.Page
{
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            BindHoushold();
            BindRRG();
        }
    }
    public void BindHoushold()
    {
        ddlHH.Items.Clear();
        string sqlstr = @"SP_S_HouseHoldName @IncludeClassB = 1";
        clsGM.getListForBindDDL(ddlHH, sqlstr, "Name", "AccountId");

        ddlHH.Items.Insert(0, "All");
        ddlHH.Items[0].Value = "0";
        ddlHH.SelectedIndex = 0;
    }
    public void BindRRG()
    {

        object HHId = ddlHH.SelectedValue.ToString() == "0" || ddlHH.SelectedValue.ToString() == "" ? "null" : "'" + ddlHH.SelectedValue.ToString() + "'";
        ddlRRG.Items.Clear();
        string sqlstr = @"select
Distinct RRG.Sas_name,rrg.Sas_reportrollupgroupId
from sas_reportrollupgroup RRG
join (
Select RRG.Sas_reportrollupgroupId, Count(*) as CountNmb
from
( Select Distinct RRG.Sas_reportrollupgroupId, RRG.Ssi_LegalEntityId , A.Ssi_LegalEntityId as AccountLEId
from Sas_reportrollupgroup RRG
join Ssi_lookthroughaccountgroup LK on LK.Ssi_ReportGroupId = RRG.Sas_reportrollupgroupId
join ssi_Account A on A.ssi_AccountId = LK.Ssi_AccountId
Where RRG.statecode = 0
and RRG.statuscode = 1
and LK.statecode = 0
and LK.statuscode = 1
and A.statecode = 0
and A.statuscode = 1
) RRG

 

group by RRG.Sas_reportrollupgroupId
having Count(*) > 1
) RRG2 on RRG.Sas_reportrollupgroupId = RRG2.Sas_reportrollupgroupId
where RRg.statuscode = 1
and Sas_HouseholdId = " + HHId + " ORDER by RRG.Sas_name";
        clsGM.getListForBindDDL(ddlRRG, sqlstr, "Sas_name", "Sas_reportrollupgroupId");

        ddlRRG.Items.Insert(0, "All");
        ddlRRG.Items[0].Value = "0";
        ddlRRG.SelectedIndex = 0;
    }

    protected void RRG_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Visible = false;
        lblError.Text = "";
        lblMessage.Visible = false;
        lblMessage.Text = "";
        // string HouseholdTxt = ddlHouseHold.SelectedItem.ToString();
        object RRGId = ddlRRG.SelectedValue.ToString() == "0" || ddlRRG.SelectedValue.ToString() == "" ? "null" : "'" + ddlRRG.SelectedValue.ToString() + "'";

        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet(@"select 
        Distinct Acc.Ssi_LegalEntityIdName
                , Acc.Ssi_LegalEntityId
from ssi_Account Acc
Inner Join  Ssi_lookthroughaccountgroup LAG on lag.Ssi_AccountId = acc.ssi_AccountId
where
         Ssi_ReportGroupId = " + RRGId);
        lstLegalEntity.Items.Clear();
        //   ddlGAGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            //  string val = loDataset.Tables[0].Rows[liCounter][1].ToString();

            //  ddlGAGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
            //  lstLegalEntity.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(liCounter)));
            lstLegalEntity.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][1])));
        }
    }

    protected void btnSumbit_Click(object sender, EventArgs e)
    {
        try
        {
            ViewState["newRRGId"] = null;
            string strDescription = string.Empty;
            IOrganizationService service = null;
            bool bProceed = false;
            try
            {

                service = clsGM.GetCrmService();
                bProceed = true;

            }

            catch (Exception Exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + Exc.Message.ToString();
                lblError.Visible = true;
                lblError.Text = strDescription;
            }
            if (bProceed)
            {
                int CreateCount = 0;
                int RecordCount = 0;
                string greshamquery = string.Empty;
                object RRGId = ddlRRG.SelectedValue.ToString() == "0" || ddlRRG.SelectedValue.ToString() == "" ? "null" : "'" + ddlRRG.SelectedValue.ToString() + "'";
                object LegalEntity = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";

                greshamquery = "SP_S_CleanUpRRG @ReportGroupID =" + RRGId + ",@LegalEntityIdList = " + LegalEntity + ",@NewReportGroupID =null";



                DataSet loDataset = clsDB.getDataSet(greshamquery);
                DataTable dtRRG = loDataset.Tables[0];
                RecordCount = loDataset.Tables[0].Rows.Count;
                if (loDataset.Tables[0].Rows.Count > 0)
                {
                    //bool UpdatedOldRRG = false;
                    DataTable dt = new DataTable();
                    dt.Columns.Add("NewRRGId");
                    dt.Columns.Add("LegalEntityId");
                    for (int i = 0; i < loDataset.Tables[0].Rows.Count; i++)
                    {
                        Entity objRRG = new Entity("sas_reportrollupgroup");
                        if (Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]) != null)
                        {
                            objRRG["sas_name"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);

                        }

                        if (Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_HouseholdId"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_HouseholdId"]) != null)
                        {
                            objRRG["sas_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_HouseholdId"])));

                        }
                        if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_LegalEntityId"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_LegalEntityId"]) != null)
                        {
                            objRRG["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_LegalEntityId"])));

                        }
                        if (Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_FamilyId"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_FamilyId"]) != null)
                        {
                            objRRG["sas_familyid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_greshamfamily", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_FamilyId"])));

                        }
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_LegalEntityId"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_LegalEntityId"]) != null)
                        //{
                        //    objRRG["ssi_ownerlegalentity"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_LegalEntityId"])));

                        //}

                        objRRG["ssi_cleanupstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000000);
                        string UserId = GetcurrentUser();

                        if (UserId != "")
                        {
                            objRRG["ssi_reviwedby"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                        }
                        objRRG["ssi_reviewstartdt"] = Convert.ToDateTime(DateTime.Now.ToString());


                        ////if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroup"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroup"]) != null)
                        ////{
                        ////    objRRG["ssi_allocationgroup"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroup"])); ;

                        ////}
                        //if (Convert.ToBoolean(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroup"]) == true)
                        //{
                        //    objRRG["ssi_allocationgroup"] = true;

                        //}
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroupTitle1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroupTitle1"]) != null)
                        //{
                        //    objRRG["ssi_allocationgrouptitle1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroupTitle1"]);

                        //}
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther2"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther2"]) != null)
                        //{
                        //    objRRG["ssi_typeother2"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther2"]);

                        //}
                        ////if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Report2"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Report2"]) != null)
                        ////{
                        ////    objRRG["ssi_report2"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_Report2"]));

                        ////}
                        //if (Convert.ToBoolean(loDataset.Tables[0].Rows[i]["Ssi_Report2"]) == true)
                        //{
                        //    objRRG["ssi_report2"] = true;

                        //}
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnName"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnName"]) != null)
                        //{
                        //    objRRG["ssi_meetingbookreporttab2columnname"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnName"]);

                        //}
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnOrder"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnOrder"]) != null)
                        //{
                        //    objRRG["ssi_meetingbookreporttab2columnorder"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnOrder"]);

                        //}

                        ////if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Performance"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Performance"]) != null)
                        ////{
                        ////    objRRG["ssi_performance"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_Performance"]));
                        ////}
                        ////if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Report"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Report"]) != null)
                        ////{
                        ////    objRRG["ssi_report"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_Report"])); ;

                        ////}
                        //if (Convert.ToBoolean(loDataset.Tables[0].Rows[i]["Ssi_Performance"]) == true)
                        //{
                        //    objRRG["ssi_performance"] = true;
                        //}
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Report"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Report"]) != null)
                        //{
                        //    if (Convert.ToBoolean(loDataset.Tables[0].Rows[i]["Ssi_Report"]) == true )
                        //    {
                        //        objRRG["ssi_report"] = true;

                        //    }
                        //}
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther1"]) != null)
                        //{
                        //    objRRG["ssi_typeother1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther1"]);

                        //}
                        //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Other1Order"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Other1Order"]) != null)
                        //{
                        //    objRRG["ssi_other1order"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_Other1Order"]);

                        //}

                        Guid newRRGId = service.Create(objRRG);

                        // Guid newRRGId = new Guid("");
                        CreateCount++;

                        DataRow dr = dt.NewRow();
                        dr["NewRRGId"] = newRRGId;
                        dr["LegalEntityId"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_LegalEntityId"]);
                        dt.Rows.Add(dr);



                        /* #region Update old RRG with Z
                           //  Entity oldRRG = new Entity();
                           if (!UpdatedOldRRG)
                           {
                               Entity oldRRG = new Entity("sas_reportrollupgroup");
                               oldRRG["sas_reportrollupgroupid"] = new Guid(ddlRRG.SelectedValue.ToString());
                               if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther2"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther2"]) != null)
                               {
                                   oldRRG["ssi_typeother2"] = "Z" + Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther2"]);
                               }
                               if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroupTitle1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroupTitle1"]) != null)
                               {
                                   oldRRG["ssi_allocationgrouptitle1"] = "Z" + Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AllocationGroupTitle1"]);
                               }
                               if (Convert.ToBoolean(loDataset.Tables[0].Rows[i]["Ssi_Report"]) == true)//HH Report is checked
                               {
                                   oldRRG["ssi_report"] = false; // HH Report
                                   if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther1"]) != null)
                                   {
                                       oldRRG["ssi_typeother1"] = "Z" + Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_TypeOther1"]);//HH Report Col Name

                                   }

                                   objRRG["ssi_hhrrgcleanupflg"] = true;//  Removed from report during RRG Cleanup
                               }

                               if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnName"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnName"]) != null)
                               {
                                   oldRRG["ssi_meetingbookreporttab2columnname"] ="Z" + Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MeetingBookReportTab2ColumnName"]);
                               }


                               service.Update(oldRRG);
                               UpdatedOldRRG = true;
                           }


                           #endregion */

                    }

                    if (ViewState["newRRGId"] == null)
                    {
                        ViewState["newRRGId"] = dt;
                    }

                    lblMessage.Visible = true;
                    lblMessage.Text = CreateCount + " Report Rollup Group Created out of " + RecordCount + " Successfully " + "<br />" + "Please run your BEFORE comparison reports";
                    btnLookthroughAccount.Enabled = true;
                    btnLookthroughAccount.Visible = true;
                }
                else
                {

                }
                //if (CreateCount == RecordCount)
                //  {

                //  }
            }
        }
        catch (Exception ex)
        {
            lblError.Visible = true;
            lblError.Text = "Error Occured: " + ex.Message.ToString();
        }
    }

    protected void btnLookthroughAccount_Click(object sender, EventArgs e)
    {
        try


        {

            lblMessage.Text = "";
            lblMessage.Visible = false;
            string strDescription = string.Empty;
            IOrganizationService service = null;
            bool bProceed = false;
            try
            {

                service = clsGM.GetCrmService();
                bProceed = true;

            }

            catch (Exception Exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + Exc.Message.ToString();
                lblError.Visible = true;
                lblError.Text = strDescription;
            }
            if (bProceed)
            {
                int CreateCount = 0;
                int RecordCount = 0;
                string greshamquery = string.Empty;
                object RRGId = ddlRRG.SelectedValue.ToString() == "0" || ddlRRG.SelectedValue.ToString() == "" ? "null" : "'" + ddlRRG.SelectedValue.ToString() + "'";
                // object LegalEntity = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";
                // object newRRGId = ViewState["newRRGId"].ToString();
                //  greshamquery = "SP_S_CleanUpRRG @ReportGroupID =" + RRGId + ",@LegalEntityIdList = " + LegalEntity + ",@NewReportGroupID ='" + newRRGId + "',@updateflg = 1";
                foreach (ListItem item in lstLegalEntity.Items)
                {
                    bool isSelected = item.Selected;
                    string LEid = item.Value;
                    string id3 = item.Text;

                    if (isSelected)
                    {
                        bool UpdatedOldRRG = false;
                        bool UpdatedNewRRG = false;
                        if (ViewState["newRRGId"] != null)
                        {
                            DataTable dtNewRRG = (DataTable)ViewState["newRRGId"];
                            for (int j = 0; j < dtNewRRG.Rows.Count; j++)
                            {
                                string NewRRGId = dtNewRRG.Rows[j]["NewRRGId"].ToString();
                                string LegalEntityID = dtNewRRG.Rows[j]["LegalEntityId"].ToString();
                                if (LegalEntityID == LEid)
                                {
                                    greshamquery = "SP_S_CleanUpRRG @ReportGroupID =" + RRGId + ",@LegalEntityIdList = '" + LegalEntityID + "',@NewReportGroupID ='" + NewRRGId + "',@updateflg = 1";
                                    DataSet loDataset = clsDB.getDataSet(greshamquery);
                                    DataTable dtLTA = loDataset.Tables[2];
                                    RecordCount = loDataset.Tables[2].Rows.Count;
                                    if (loDataset.Tables[2].Rows.Count > 0)
                                    {
                                        for (int i = 0; i < loDataset.Tables[2].Rows.Count; i++)
                                        {
                                            Entity objLookthroughaccountgroup = new Entity("ssi_lookthroughaccountgroup");


                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_AccountId"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_AccountId"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_AccountId"])));

                                            }
                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_ReportGroupId"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_ReportGroupId"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_reportgroupid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_reportrollupgroup", new Guid(Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_ReportGroupId"])));

                                            }
                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["LookThroughName"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["LookThroughName"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_name"] = Convert.ToString(loDataset.Tables[2].Rows[i]["LookThroughName"]);
                                            }

                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_Ownership"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_Ownership"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_ownership"] = Convert.ToDecimal(loDataset.Tables[2].Rows[i]["Ssi_Ownership"]);
                                            }
                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_StartDate"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_StartDate"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_startdate"] = Convert.ToDateTime(loDataset.Tables[2].Rows[i]["Ssi_StartDate"]);
                                            }
                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_EndDate"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["Ssi_EndDate"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_enddate"] = Convert.ToDateTime(loDataset.Tables[2].Rows[i]["Ssi_EndDate"]);
                                            }

                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["Other1"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["Other1"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_typeother1"] = Convert.ToString(loDataset.Tables[2].Rows[i]["Other1"]);
                                            }

                                            //if (Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingException"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingException"]) != null)
                                            //{
                                            //    objLookthroughaccountgroup["ssi_billingexception"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[2].Rows[i]["ssi_BillingException"]));
                                            //}
                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingException"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingException"]) != null)
                                            {
                                                if (Convert.ToBoolean(loDataset.Tables[2].Rows[i]["ssi_BillingException"]) == true)
                                                {
                                                    objLookthroughaccountgroup["ssi_billingexception"] = true;
                                                }
                                            }


                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingOwnership"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingOwnership"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_billingownership"] = Convert.ToDecimal(loDataset.Tables[2].Rows[i]["ssi_BillingOwnership"]);
                                            }
                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingStartDate"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingStartDate"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_billingstartdate"] = Convert.ToDateTime(loDataset.Tables[2].Rows[i]["ssi_BillingStartDate"]);
                                            }
                                            if (Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingEndDate"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_BillingEndDate"]) != null)
                                            {
                                                objLookthroughaccountgroup["ssi_billingenddate"] = Convert.ToDateTime(loDataset.Tables[2].Rows[i]["ssi_BillingEndDate"]);
                                            }

                                            service.Create(objLookthroughaccountgroup);
                                            CreateCount++;


                                            #region Update old RRG with Z
                                            //  Entity oldRRG = new Entity();
                                            if (j == dtNewRRG.Rows.Count - 1)
                                            {


                                                if (!UpdatedOldRRG)
                                                {
                                                    Entity oldRRG = new Entity("sas_reportrollupgroup");
                                                    oldRRG["sas_reportrollupgroupid"] = new Guid(ddlRRG.SelectedValue.ToString());
                                                    if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther2"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther2"]) != null)
                                                    {
                                                        oldRRG["ssi_typeother2"] = "Z" + Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther2"]);
                                                    }
                                                    if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroupTitle1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroupTitle1"]) != null)
                                                    {
                                                        oldRRG["ssi_allocationgrouptitle1"] = "Z" + Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroupTitle1"]);
                                                    }

                                                    if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Report"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Report"]) != null)
                                                    {
                                                        if (Convert.ToBoolean(loDataset.Tables[0].Rows[0]["Ssi_Report"]) == true)//HH Report is checked
                                                        {
                                                            oldRRG["ssi_report"] = false; // HH Report
                                                            if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther1"]) != null)
                                                            {
                                                                oldRRG["ssi_typeother1"] = "Z" + Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther1"]);//HH Report Col Name

                                                            }

                                                            oldRRG["ssi_hhrrgcleanupflg"] = true;//  Removed from report during RRG Cleanup
                                                        }
                                                    }

                                                    if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnName"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnName"]) != null)
                                                    {
                                                        oldRRG["ssi_meetingbookreporttab2columnname"] = "Z" + Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnName"]);
                                                    }


                                                    service.Update(oldRRG);
                                                    UpdatedOldRRG = true;
                                                }
                                            }
                                            #endregion
                                            #region Update new RRg 
                                            if (!UpdatedNewRRG)
                                            {
                                                Entity objRRG = new Entity("sas_reportrollupgroup");
                                                objRRG["sas_reportrollupgroupid"] = new Guid(NewRRGId);

                                                if (Convert.ToString(loDataset.Tables[2].Rows[i]["Name"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["Name"]) != null)
                                                {
                                                    objRRG["sas_name"] = Convert.ToString(loDataset.Tables[2].Rows[i]["Name"]);

                                                }


                                                if (Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_OwnerLegalEntityID"]) != "" && Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_OwnerLegalEntityID"]) != null)
                                                {
                                                    objRRG["ssi_ownerlegalentity"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[2].Rows[i]["ssi_OwnerLegalEntityID"])));

                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroup"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroup"]) != null)
                                                {
                                                    if (Convert.ToBoolean(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroup"]) == true)
                                                    {
                                                        objRRG["ssi_allocationgroup"] = true;

                                                    }
                                                }

                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroupTitle1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroupTitle1"]) != null)
                                                {
                                                    objRRG["ssi_allocationgrouptitle1"] = Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_AllocationGroupTitle1"]);

                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther2"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther2"]) != null)
                                                {
                                                    objRRG["ssi_typeother2"] = Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther2"]);

                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Report2"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Report2"]) != null)
                                                {
                                                    if (Convert.ToBoolean(loDataset.Tables[0].Rows[0]["Ssi_Report2"]) == true)
                                                    {
                                                        objRRG["ssi_report2"] = true;

                                                    }
                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnName"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnName"]) != null)
                                                {
                                                    objRRG["ssi_meetingbookreporttab2columnname"] = Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnName"]);

                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnOrder"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnOrder"]) != null)
                                                {
                                                    objRRG["ssi_meetingbookreporttab2columnorder"] = Convert.ToInt32(loDataset.Tables[0].Rows[0]["Ssi_MeetingBookReportTab2ColumnOrder"]);

                                                }

                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Performance"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Performance"]) != null)
                                                {
                                                    if (Convert.ToBoolean(loDataset.Tables[0].Rows[0]["Ssi_Performance"]) == true)
                                                    {
                                                        objRRG["ssi_performance"] = true;
                                                    }
                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Report"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Report"]) != null)
                                                {
                                                    if (Convert.ToBoolean(loDataset.Tables[0].Rows[0]["Ssi_Report"]) == true)
                                                    {
                                                        objRRG["ssi_report"] = true;

                                                    }
                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther1"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther1"]) != null)
                                                {
                                                    objRRG["ssi_typeother1"] = Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_TypeOther1"]);

                                                }
                                                if (Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Other1Order"]) != "" && Convert.ToString(loDataset.Tables[0].Rows[0]["Ssi_Other1Order"]) != null)
                                                {
                                                    objRRG["ssi_other1order"] = Convert.ToInt32(loDataset.Tables[0].Rows[0]["Ssi_Other1Order"]);

                                                }
                                                service.Update(objRRG);
                                                #endregion
                                            }
                                        }





                                        //if (CreateCount == RecordCount)
                                        //{
                                        lblMessage.Visible = true;
                                        //lblMessage.Text = CreateCount + " Look through Account Group Created out of " + RecordCount + " Successfully" + "<br />" + " Please run your AFTER comparison reports";
                                        lblMessage.Text = CreateCount + " Look through Account Group Created " + "<br />" + " Please run your AFTER comparison reports";
                                        //  btnLookthroughAccount.Enabled = false;
                                        //btnLookthroughAccount.Visible = true;
                                        //  }
                                    }
                                }
                            }
                        }
                        else

                        {
                            lblMessage.Visible = true;
                            lblMessage.Text = "No Records Found";
                        }
                    }
                }




            }
        }
        catch (Exception exe)
        {
            lblError.Visible = true;
            lblError.Text = "Error Occured: " + exe.Message.ToString();
        }

    }

    protected void ddlHH_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindRRG();
    }
    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;

        string strName = string.Empty;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            strName = "corp\\gbhagia";
        }
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            strName = claimsIdentity.Name;

        }

        string sqlstr = string.Empty;
        sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        //Response.Write(strName + "<br/><br/>");
        //Response.Write(Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]));
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
        }
        else
        {
            return UserID = "";
        }
    }

}