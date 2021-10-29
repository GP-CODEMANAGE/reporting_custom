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
using System.Security.Principal;
using System.IO;
using System.Text;
using Spire.Xls;
using System.Xml;
//using CrmSdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class HistoricalAccountCopy : System.Web.UI.Page
{
    string sqlstr = string.Empty;
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    public StreamWriter sw = null;
    public string execType = string.Empty;
    string strDescription = string.Empty;

    public const string Position = "Summary";
    public const string Transaction = "Transaction";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            fillHousehold();
            lnkTransactionErrorFile.Style.Add("display", "none");
            lnkPositionErrorFile.Style.Add("display", "none");
            ViewState["TrxnErrorDT"]=null;
            ViewState["PostnErrorDT"] = null;
            chkCopyAllTranPostn.Checked = true;
        }
    }

    public void fillHousehold()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("SP_S_GET_HOUSEHOLDNAME_HISTORICAL");
        ddlHouseHold.Items.Clear();
        ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }
    }

    private string GetSqlString(string Type)
    {
         string StartDate =string.Empty;
         string EndDate = string.Empty;
         if (chkCopyAllTranPostn.Checked)
         {
             StartDate = "null";
             EndDate = "null";
         }
         else
         {
             StartDate = "'" + txtStartDate.Text + "'";
             EndDate = "'" + txtEndDate.Text + "'";
         }
        if (Type == Transaction)
        {
            sqlstr = "SP_S_HISTORICAL_TRXN_COPY @StrtDt=" + StartDate + "" +
                                 ",@EndDt=" + EndDate + "" +
                                 ",@HHUUID='" + ddlHouseHold.SelectedValue + "'";

        }
        else if (Type == Position)
        {
            sqlstr = "SP_S_HISTORICAL_POSITION_COPY @StrtDt=" + StartDate + "" +
                                 ",@EndDt=" + EndDate + "" +
                                 ",@HHUUID='" + ddlHouseHold.SelectedValue + "'";
        }

        return sqlstr;
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        bool status = LoadData();

        if (status == true)
        {
            if (chkCopyAllTranPostn.Checked == true)
                lblError.Text = "All positions and transactions have been copied.";
            else
                lblError.Text = "All positions and transactions in the date range indicated have been copied.";
        }
        else
        {
            lblError.Text = "No records found.";
            lnkTransactionErrorFile.Style.Add("display", "none");
            lnkPositionErrorFile.Style.Add("display", "none");
        }
        trSubmit.Style.Add("display", "inline");
    }

    private bool LoadData()
    {
        #region Declaration
        string LogFileName = "LogFile " + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        sw = new StreamWriter(Request.PhysicalApplicationPath + "\\Log\\" + LogFileName + ".txt", true);
        //sw = new StreamWriter(LogFileName + ".txt", true);

        // string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
        //string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);

        // string CRM_constring = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";

        ////string crmServerUrl = "http://Crm01/";
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        ////string crmServerURL = "http://server:5555/";

        //string orgName = "GreshamPartners";
        //string orgName = "Webdev";

        string Userid = GetcurrentUser();
        bool bProceed = true;
        string strDescription;
       // CrmService service = null;
        IOrganizationService service = null;
        try
        {
           // service = GetCrmService(crmServerUrl, orgName, Userid);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            LogMessage(sw, service, strDescription, 62, "GeneralError");
            sw.WriteLine("step 1 ");
        }
       // catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {

            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            LogMessage(sw, service, strDescription, 62, "GeneralError");
        }

        //service.PreAuthenticate = true;

        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);

        SqlConnection CRM_con;
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery;
        int totalCount = 0;
        int successCount = 0;
        int failiureCount = 0;
        int sleepTime = 0;

        int Pload = 0;
        int Tload = 0;

        int totalCountTransaction = 0;
        int totalCountTransactionAccount = 0;
        int totalCountTransactionError = 0;
        int totalCountPosition = 0;
        int totalCountPositionAccount = 0;
        int totalCountPositionError = 0;


        #endregion

        if (bProceed == true)
        {
            // using data load from file
            #region Transaction Position Load


            successCount = 0;
            failiureCount = 0;
            // bProceed = true;
            bool positionSuccess = true;
            bool transactionSuccess = true;
            execType = "B";

            if (bProceed == true)
            {
                if (execType == "P" || execType == "B")
                {
                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    #region Transaction
                    try
                    {
                        sw.WriteLine("---------------------------- New Transaction Insert Starts -------------------");
                        greshamquery = GetSqlString(Transaction);
                        dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        ds_gresham = new DataSet();
                        dagersham.SelectCommand.CommandTimeout = 1800;
                        dagersham.Fill(ds_gresham);
                        totalCount = ds_gresham.Tables[0].Rows.Count;
                        totalCountTransaction = totalCount;

                        sw.WriteLine("Insert Starts for New Transaction on: " + DateTime.Now.ToString());
                    }
                    catch (System.Web.Services.Protocols.SoapException exc)
                    {
                        bProceed = true;
                        totalCount = 0;
                        transactionSuccess = false;
                        strDescription = "Transaction Insert failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                        LogMessage(sw, service, strDescription, 62, "HistoricalAccountCopy");
                    }
                    catch (Exception exc)
                    {
                        bProceed = true;
                        totalCount = 0;
                        transactionSuccess = false;
                        strDescription = "Transaction Insert failed, please contact administrator. Error Detail: " + exc.Message;
                        LogMessage(sw, service, strDescription, 62, "HistoricalAccountCopy");
                    }
                    #region Transaction Insert
                    if (bProceed == true)
                        for (int i = 0; i < totalCount; i++)
                        {
                            try
                            {
                                if (bProceed == true)
                                {
                                 //   ssi_transactionlog objTransaction = new ssi_transactionlog();
                                    Entity objTransaction = new Entity("ssi_transactionlog");
                                   
                                    //objTransaction.ssi_transactionlogid = new Key();
                                    //objTransaction.ssi_transactionlogid.Value = Guid.NewGuid();

                                    //account
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) != "")
                                    {
                                        //objTransaction.ssi_accountid = new Lookup();
                                        //objTransaction.ssi_accountid.type = EntityName.ssi_account.ToString();
                                        //objTransaction.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));
                                        objTransaction["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"])));
                                    }

                                    //Security
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]) != "")
                                    {
                                        //objTransaction.ssi_securityid = new Lookup();
                                        //objTransaction.ssi_securityid.type = EntityName.ssi_security.ToString();
                                        //objTransaction.ssi_securityid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]));
                                        objTransaction["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"])));
                                    }

                                    //Assetclassid 
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_assetclassid"]) != "")
                                    {
                                        //objTransaction.ssi_assetclassid = new Lookup();
                                        //objTransaction.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                        //objTransaction.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_assetclassid"]));
                                        objTransaction["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_assetclassid"])));
                                    }

                                    //SectorId 
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]) != "")
                                    {
                                        //objTransaction.ssi_sectorid = new Lookup();
                                        //objTransaction.ssi_sectorid.type = EntityName.ssi_sector.ToString();
                                        //objTransaction.ssi_sectorid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]));
                                        objTransaction["ssi_sectorid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_sector", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"])));
                                    }

                                    //quantity
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Quantity"]) != "" && Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Quantity"]) != "0")
                                    {
                                        //objTransaction.ssi_quantity = new CrmDecimal();
                                        //objTransaction.ssi_quantity.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Quantity"]);
                                        objTransaction["ssi_quantity"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Quantity"]);

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_quantity = new CrmDecimal();
                                        //objTransaction.ssi_quantity.IsNull = true;
                                        //objTransaction.ssi_quantity.IsNullSpecified = true;
                                        objTransaction["ssi_quantity"] = null;
                                    }

                                    //Trade date
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TradeDate"]) != "")
                                    {
                                        //objTransaction.ssi_tradedate = new CrmDateTime();
                                        //objTransaction.ssi_tradedate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TradeDate"]);
                                        objTransaction["ssi_tradedate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_TradeDate"]);

                                    }

                                    //Trade Amt
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Value"]) != "")
                                    {
                                        //objTransaction.ssi_value = new CrmMoney();
                                        //objTransaction.ssi_value.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Value"]);
                                        objTransaction["ssi_value"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Value"]));

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_value = new CrmMoney();
                                        //objTransaction.ssi_value.IsNull = true;
                                        //objTransaction.ssi_value.IsNullSpecified = true;
                                        objTransaction["ssi_value"] = null;
                                    }

                                    //transactioncodeid
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TransactionCodeId"]) != "")
                                    {
                                        //objTransaction.ssi_transactioncodeid = new Lookup();
                                        //objTransaction.ssi_transactioncodeid.type = EntityName.ssi_transactionlog.ToString();
                                        //objTransaction.ssi_transactioncodeid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TransactionCodeId"]));
                                        objTransaction["ssi_transactioncodeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_transactionlog", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TransactionCodeId"])));
                                    }

                                    //ssi_lock
                                    if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Ssi_Lock"]) == true)
                                    {
                                        //objTransaction.ssi_lock = new CrmBoolean();
                                        //objTransaction.ssi_lock.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Ssi_Lock"]);
                                        objTransaction["ssi_lock"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Lock"]).ToLower());

                                    }

                                    // Txn Num
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionNum"]) != "")
                                    {
                                        //objTransaction.ssi_tnrtransactionnum = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionNum"]);
                                        objTransaction["ssi_tnrtransactionnum"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionNum"]);
                                    }

                                    // Transaction Type
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionType"]) != "")
                                    {
                                        //objTransaction.ssi_tnrtransactiontype = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionType"]);
                                        objTransaction["ssi_tnrtransactiontype"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionType"]);
                                    }

                                    // Transaction Name
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionName"]) != "")
                                    {
                                       // objTransaction.ssi_tnrtransactionname = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionName"]);
                                        objTransaction["ssi_tnrtransactionname"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionName"]);
                                    }

                                    // LastLockDate
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_LastLockDate"]) != "")
                                    {
                                        //objTransaction.ssi_lastlockdate = new CrmDateTime();
                                        //objTransaction.ssi_lastlockdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_LastLockDate"]);
                                        objTransaction["ssi_lastlockdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_LastLockDate"]);

                                    }

                                    if (Userid != "")
                                    {
                                        //objTransaction.createdby = new Lookup();
                                        //objTransaction.createdby.type = EntityName.systemuser.ToString();
                                        //objTransaction.createdby.Value = new Guid(Userid);
                                        objTransaction["createdby"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Convert.ToString(Userid)));
                                    }

                                    //SourceCode (TNR)
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_DataSource"]) != "")
                                    {
                                        //objTransaction.ssi_datasource = new Picklist();
                                        //objTransaction.ssi_datasource.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["Ssi_DataSource"]);
                                        objTransaction["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["Ssi_DataSource"]));

                                    }
                                    //***********New Fields **********************//
                                    // Comment
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Comment"]) != "")
                                    {
                                        //objTransaction.ssi_comment = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Comment"]);
                                        objTransaction["ssi_comment"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Comment"]);
                                    }

                                    //Unrealized
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Unrealized"]) != "")
                                    {
                                        //objTransaction.ssi_unrealized = new CrmDecimal();
                                        //objTransaction.ssi_unrealized.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Unrealized"]);
                                        objTransaction["ssi_unrealized"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Unrealized"]);

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_unrealized = new CrmDecimal();
                                        //objTransaction.ssi_unrealized.IsNull = true;
                                        //objTransaction.ssi_unrealized.IsNullSpecified = true;
                                        objTransaction["ssi_unrealized"] = null;

                                    }

                                    //Valuation
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Valuation"]) != "")
                                    {
                                        //objTransaction.ssi_valuation = new CrmDecimal();
                                        //objTransaction.ssi_valuation.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Valuation"]);
                                        objTransaction["ssi_valuation"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Valuation"]);
                                    }
                                    else
                                    {
                                        //objTransaction.ssi_valuation = new CrmDecimal();
                                        //objTransaction.ssi_valuation.IsNull = true;
                                        //objTransaction.ssi_valuation.IsNullSpecified = true;
                                        objTransaction["ssi_valuation"] = null;
                                    }

                                    // Date Received
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_DateReceived"]) != "")
                                    {
                                        //objTransaction.ssi_datereceived = new CrmDateTime();
                                        //objTransaction.ssi_datereceived.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_DateReceived"]);
                                        objTransaction["ssi_datereceived"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_DateReceived"]);

                                    }

                                    //Amount Received
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AmountReceived"]) != "")
                                    {
                                        //objTransaction.ssi_amountreceived = new CrmMoney();
                                        //objTransaction.ssi_amountreceived.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_AmountReceived"]);
                                        objTransaction["ssi_amountreceived"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_AmountReceived"]));

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_amountreceived = new CrmMoney();
                                        //objTransaction.ssi_amountreceived.IsNull = true;
                                        //objTransaction.ssi_amountreceived.IsNullSpecified = true;
                                        objTransaction["ssi_amountreceived"] = null;
                                    }

                                    //Balance
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Balance"]) != "")
                                    {
                                        //objTransaction.ssi_balance = new CrmDecimal();
                                        //objTransaction.ssi_balance.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Balance"]);
                                        objTransaction["ssi_balance"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Balance"]);

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_balance = new CrmDecimal();
                                        //objTransaction.ssi_balance.IsNull = true;
                                        //objTransaction.ssi_balance.IsNullSpecified = true;
                                        objTransaction["ssi_balance"] = null;
                                    }

                                    //Cost
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Cost"]) != "")
                                    {
                                        //objTransaction.ssi_cost = new CrmMoney();
                                        //objTransaction.ssi_cost.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Cost"]);
                                        objTransaction["ssi_cost"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Cost"]));
                                    }
                                    else
                                    {
                                        //objTransaction.ssi_cost = new CrmMoney();
                                        //objTransaction.ssi_cost.IsNull = true;
                                        //objTransaction.ssi_cost.IsNullSpecified = true;
                                        objTransaction["ssi_cost"] = null;
                                    }

                                    //Gain Loss
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_GainLoss"]) != "")
                                    {
                                        //objTransaction.ssi_gainloss = new CrmDecimal();
                                        //objTransaction.ssi_gainloss.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_GainLoss"]);
                                        objTransaction["ssi_gainloss"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_GainLoss"]);

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_gainloss = new CrmDecimal();
                                        //objTransaction.ssi_gainloss.IsNull = true;
                                        //objTransaction.ssi_gainloss.IsNullSpecified = true;
                                        objTransaction["ssi_gainloss"] = null;
                                    }

                                    //Withhold
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Withhold"]) != "")
                                    {
                                        //objTransaction.ssi_withhold = new CrmDecimal();
                                        //objTransaction.ssi_withhold.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Withhold"]);
                                        objTransaction["ssi_withhold"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Withhold"]);
                                    }
                                    else
                                    {
                                        //objTransaction.ssi_withhold = new CrmDecimal();
                                        //objTransaction.ssi_withhold.IsNull = true;
                                        //objTransaction.ssi_withhold.IsNullSpecified = true;
                                        objTransaction["ssi_withhold"] = null;
                                    }

                                    //Distribution Amount
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_DistributionAmount"]) != "")
                                    {
                                        //objTransaction.ssi_distributionamount = new CrmMoney();
                                        //objTransaction.ssi_distributionamount.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_DistributionAmount"]);
                                        objTransaction["ssi_distributionamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_DistributionAmount"]));

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_distributionamount = new CrmMoney();
                                        //objTransaction.ssi_distributionamount.IsNull = true;
                                        //objTransaction.ssi_distributionamount.IsNullSpecified = true;
                                        objTransaction["ssi_distributionamount"] = null;
                                    }

                                    //Distribution Withholding Percent
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_DistributionWithholdingPercent"]) != "")
                                    {
                                        //objTransaction.ssi_distributionwithholdingpercent = new CrmDecimal();
                                        //objTransaction.ssi_distributionwithholdingpercent.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_DistributionWithholdingPercent"]);
                                        objTransaction["ssi_distributionwithholdingpercent"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_DistributionWithholdingPercent"]);

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_distributionwithholdingpercent = new CrmDecimal();
                                        //objTransaction.ssi_distributionwithholdingpercent.IsNull = true;
                                        //objTransaction.ssi_distributionwithholdingpercent.IsNullSpecified = true;
                                        objTransaction["ssi_distributionwithholdingpercent"] = null;
                                    }

                                    //1099 Amount
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_1099Amount"]) != "")
                                    {
                                        //objTransaction.ssi_1099amount = new CrmMoney();
                                        //objTransaction.ssi_1099amount.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_1099Amount"]);
                                        objTransaction["ssi_1099amount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_1099Amount"]));

                                    }
                                    else
                                    {
                                        //objTransaction.ssi_1099amount = new CrmMoney();
                                        //objTransaction.ssi_1099amount.IsNull = true;
                                        //objTransaction.ssi_1099amount.IsNullSpecified = true;
                                        objTransaction["ssi_1099amount"] = null;
                                    }

                                    //Opt Out
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_OptOut"]) != "")
                                    {
                                        //objTransaction.ssi_optout = new CrmBoolean();
                                        //objTransaction.ssi_optout.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Ssi_OptOut"]);
                                        objTransaction["ssi_optout"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_OptOut"]).ToLower());

                                    }

                                    //greshamadvised (sectorflg)
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_GrehamAdvised"]) != "")
                                    {
                                        //objTransaction.ssi_grehamadvised = new CrmBoolean();
                                        //objTransaction.ssi_grehamadvised.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_GrehamAdvised"]).ToLower());
                                        objTransaction["ssi_grehamadvised"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_GrehamAdvised"]).ToLower());
                                    }

                                    //Settle date
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SettleDate"]) != "")
                                    {
                                        //objTransaction.ssi_settledate = new CrmDateTime();
                                        //objTransaction.ssi_settledate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SettleDate"]);
                                        objTransaction["ssi_settledate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_SettleDate"]);

                                    }

                                    //Currency (default to USD)
                                    //objTransaction.transactioncurrencyid = new Lookup();
                                    //objTransaction.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                    //objTransaction.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                                    objTransaction["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));

                                    //subassetclassId
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SubAssetClassId"]) != "")
                                    {
                                        //objTransaction.ssi_subassetclassid = new Lookup();
                                        //objTransaction.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                                        //objTransaction.ssi_subassetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SubAssetClassId"]));
                                        objTransaction["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SubAssetClassId"])));
                                    }

                                    //BenchmarkSubAssetClassId
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkId"]) != "")
                                    {
                                        //objTransaction.ssi_benchmarkid = new Lookup();
                                        //objTransaction.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                                        //objTransaction.ssi_benchmarkid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkId"])); ;
                                        objTransaction["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkId"])));
                                    }

                                    service.Create(objTransaction);
                                    //Thread.Sleep(sleepTime);

                                    successCount = successCount + 1;
                                }
                                else
                                    break;
                            }
                            //catch (System.Web.Services.Protocols.SoapException exc)
                            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                            {
                                bProceed = false;
                                failiureCount = failiureCount + 1;
                                transactionSuccess = false;
                                string failiureText = "Transaction Insert - Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", Transaction Name: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionName"]);
                                strDescription = failiureText + " Error Detail:" + exc.Detail.Message;
                                LogMessage(sw, service, strDescription, 26, "HistoricalAccountCopy");
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                transactionSuccess = false;
                                failiureCount = failiureCount + 1;
                                string failiureText = "Transaction Insert - Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", Transaction Name: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TNRTransactionName"]);
                                strDescription = failiureText + " Error Detail:" + exc.Message;
                                LogMessage(sw, service, strDescription, 26, "HistoricalAccountCopy");
                            }
                        }

                    sw.WriteLine("Insert Ends for Transaction on: " + DateTime.Now.ToString());

                    strDescription = "Total Transaction failed to insert: " + failiureCount;
                    LogMessage(sw, service, strDescription, 40, "HistoricalAccountCopy");

                    strDescription = "Total Transaction inserted: " + successCount;
                    LogMessage(sw, service, strDescription, 29, "HistoricalAccountCopy");
                    sw.WriteLine("---------------------------- New Transaction Insert Ends -------------------");
                    sw.WriteLine();
                    #endregion

                    #region AccountUpdate

                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    sw.WriteLine("---------------------------- Account Update Starts -------------------");
                    sw.WriteLine("Update Starts for Accounts on: " + DateTime.Now.ToString());
                    totalCount = ds_gresham.Tables[1].Rows.Count;
                    totalCountTransactionAccount = totalCount;

                    if (bProceed == true)
                        for (int i = 0; i < totalCount; i++)
                        {
                            try
                            {
                                if (bProceed == true)
                                {
                                    Guid AccountID = new Guid(Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_accountid"]));

                                    //ssi_account objAccount = new ssi_account();
                                    Entity objAccount = new Entity("ssi_account");
                                    //objAccount.ssi_accountid = new Key();
                                    //objAccount.ssi_accountid.Value = AccountID;
                                    objAccount["ssi_accountid"] = new Guid(Convert.ToString(AccountID));


                                  //  objAccount.ssi_name = Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]);
                                    objAccount["ssi_name"] = Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]);
                                    
                                    service.Update(objAccount);
                                    //Thread.Sleep(sleepTime);

                                    successCount = successCount + 1;
                                }
                                else
                                    break;
                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = false;
                                failiureCount = failiureCount + 1;
                                strDescription = "Update failed for Account " + Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]) + " Error Detail: " + exc.Message + " " + exc.Detail.InnerText;
                                LogMessage(sw, service, strDescription, 5, "HistoricalAccountCopy");
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                failiureCount = failiureCount + 1;
                                strDescription = "Update failed for Account : " + Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]) + " Error Detail: " + exc.Message;
                                LogMessage(sw, service, strDescription, 5, "HistoricalAccountCopy");
                            }
                        }
                    sw.WriteLine("Update Ends for Acoount on: " + DateTime.Now.ToString());
                    strDescription = "Total Accounts Failed to Update: " + failiureCount;
                    LogMessage(sw, service, strDescription, 31, "HistoricalAccountCopy");

                    strDescription = "Total Accounts Updated: " + successCount;
                    LogMessage(sw, service, strDescription, 4, "HistoricalAccountCopy");
                    sw.WriteLine("---------------------------- Account Update Ends  -------------------");
                    #endregion

                    #region ShowErrorFile

                    if (ds_gresham.Tables[2].Rows.Count > 0)
                    {
                        totalCountTransactionError = ds_gresham.Tables[2].Rows.Count;
                        ViewState["TrxnErrorDT"] = ds_gresham.Tables[2];
                        lnkTransactionErrorFile.Style.Add("display", "");
                    }
                    else
                    {
                        ViewState["TrxnErrorDT"] = null;
                        lnkTransactionErrorFile.Style.Add("display", "none");
                    }

                    ds_gresham.Dispose();
                    dagersham.Dispose();
                    #endregion

                    #endregion

                    ///////////////////////////////// Position ///////////////////////////////////////////

                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        ////////////////////////////// Insert Position Starts ////////////////////////////////////////////
                        #region Position
                        try
                        {
                            greshamquery = GetSqlString(Position);
                            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            ds_gresham = new DataSet();
                            dagersham.SelectCommand.CommandTimeout = 1800;
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            totalCountPosition = totalCount;
                            sw.WriteLine("---------------------------- New Position Insert Starts -------------------");
                            sw.WriteLine("Insert Starts for New Position on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            positionSuccess = false;
                            strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 62, "HistoricalAccountCopy");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            positionSuccess = false;
                            strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Message;
                            LogMessage(sw, service, strDescription, 62, "HistoricalAccountCopy");
                        }
                        #region Insert Position
                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                      //  ssi_position objPosition = new ssi_position();
                                        Entity objPosition = new Entity("ssi_position");
                                        //primary key ssi_positionid
                                        //objPosition.ssi_positionid = new Key();
                                        //objPosition.ssi_positionid.Value = Guid.NewGuid();

                                        //accountid
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountId"]) != "")
                                        {
                                            //objPosition.ssi_accountid = new Lookup();
                                            //objPosition.ssi_accountid.type = EntityName.ssi_account.ToString();
                                            //objPosition.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountId"]));
                                            objPosition["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountId"])));
                                        }

                                        //SecurityId
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SecurityId"]) != "")
                                        {
                                            //objPosition.ssi_securityid = new Lookup();
                                            //objPosition.ssi_securityid.type = EntityName.ssi_security.ToString();
                                            //objPosition.ssi_securityid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SecurityId"]));
                                            objPosition["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SecurityId"])));
                                        }

                                        //sas_assetclassid
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AssetClassId"]) != "")
                                        {
                                            //objPosition.ssi_assetclassid = new Lookup();
                                            //objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                            //objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AssetClassId"]));
                                            objPosition["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AssetClassId"])));
                                        }

                                        //SectorId
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SectorId"]) != "")
                                        {
                                            //objPosition.ssi_sectorid = new Lookup();
                                            //objPosition.ssi_sectorid.type = EntityName.ssi_sector.ToString();
                                            //objPosition.ssi_sectorid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SectorId"]));
                                            objPosition["ssi_sectorid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_sector", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SectorId"])));
                                        }

                                        //FundId
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]) != "")
                                        {
                                            //objPosition.ssi_fundid = new Lookup();
                                            //objPosition.ssi_fundid.type = EntityName.ssi_fund.ToString();
                                            //objPosition.ssi_fundid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]));
                                            objPosition["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"])));
                                        }

                                        //Quantity
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Quantity"]) != "")
                                        {
                                            //objPosition.ssi_quantity = new CrmDecimal();
                                            //objPosition.ssi_quantity.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Quantity"]);
                                            objPosition["ssi_quantity"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Quantity"]);

                                        }
                                        else
                                        {
                                            //objPosition.ssi_quantity = new CrmDecimal();
                                            //objPosition.ssi_quantity.IsNull = true;
                                            //objPosition.ssi_quantity.IsNullSpecified = true;
                                            objPosition["ssi_quantity"] = null;

                                        }

                                        // AS OF DATE / PriceDate
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AsofDate"]) != "")
                                        {
                                            //objPosition.ssi_asofdate = new CrmDateTime();
                                            //objPosition.ssi_asofdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AsofDate"]);
                                            objPosition["ssi_asofdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_AsofDate"]);

                                        }

                                        //Market Value
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_MarketValue"]) != "")
                                        {
                                            //objPosition.ssi_marketvalue = new CrmMoney();
                                            //objPosition.ssi_marketvalue.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["ssi_MarketValue"]);
                                            objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["ssi_MarketValue"]));

                                        }
                                        else
                                        {
                                            //objPosition.ssi_marketvalue = new CrmMoney();
                                            //objPosition.ssi_marketvalue.IsNull = true;
                                            //objPosition.ssi_marketvalue.IsNullSpecified = true;
                                            objPosition["ssi_marketvalue"] = null;
                                        }

                                        //CommitmentSummaryFlg,  
                                        if (Convert.ToString((ds_gresham.Tables[0].Rows[i]["Ssi_CommitmentSummaryFlg"])) != "")
                                        {
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Ssi_CommitmentSummaryFlg"]) == true)
                                            {
                                                //objPosition.ssi_commitmentsummaryflg = new CrmBoolean();
                                                //objPosition.ssi_commitmentsummaryflg.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Ssi_CommitmentSummaryFlg"]);
                                                objPosition["ssi_commitmentsummaryflg"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_CommitmentSummaryFlg"]).ToLower());

                                            }
                                        }

                                        //ssi_lock
                                        if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Ssi_Lock"]) == true)
                                        {
                                            //objPosition.ssi_lock = new CrmBoolean();
                                            //objPosition.ssi_lock.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Ssi_Lock"]);
                                            objPosition["ssi_lock"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Lock"]).ToLower());
                                        }

                                        //ssi_committedcapital
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_CommittedCapital"]) != "")
                                        {
                                            //objPosition.ssi_committedcapital = new CrmDecimal();
                                            //objPosition.ssi_committedcapital.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_CommittedCapital"]);
                                            objPosition["ssi_committedcapital"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_CommittedCapital"]);

                                        }
                                        else
                                        {
                                            //objPosition.ssi_committedcapital = new CrmDecimal();
                                            //objPosition.ssi_committedcapital.IsNull = true;
                                            //objPosition.ssi_committedcapital.IsNullSpecified = true;
                                            objPosition["ssi_committedcapital"] = null;

                                        }

                                        //Remaining Capital
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_RemainingCapital"]) != "")
                                        {
                                            //objPosition.ssi_remainingcapital = new CrmDecimal();
                                            //objPosition.ssi_remainingcapital.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_RemainingCapital"]);
                                            objPosition["ssi_remainingcapital"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_RemainingCapital"]);
                                        }
                                        else
                                        {
                                            //objPosition.ssi_remainingcapital = new CrmDecimal();
                                            //objPosition.ssi_remainingcapital.IsNull = true;
                                            //objPosition.ssi_remainingcapital.IsNullSpecified = true;
                                            objPosition["ssi_remainingcapital"] = null;
                                        }

                                        //Current Month Paid In
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_CurrentMonthPaidIn"]) != "")
                                        {
                                            //objPosition.ssi_currentmonthpaidin = new CrmDecimal();
                                            //objPosition.ssi_currentmonthpaidin.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_CurrentMonthPaidIn"]);
                                            objPosition["ssi_currentmonthpaidin"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_CurrentMonthPaidIn"]);
                                        }
                                        else
                                        {
                                            //objPosition.ssi_currentmonthpaidin = new CrmDecimal();
                                            //objPosition.ssi_currentmonthpaidin.IsNull = true;
                                            //objPosition.ssi_currentmonthpaidin.IsNullSpecified = true;
                                            objPosition["ssi_currentmonthpaidin"] = null;
                                        }

                                        //Current Month Distribution
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_CurrentMonthDistribution"]) != "")
                                        {
                                            //objPosition.ssi_currentmonthdistribution = new CrmDecimal();
                                            //objPosition.ssi_currentmonthdistribution.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_CurrentMonthDistribution"]);
                                            objPosition["ssi_currentmonthdistribution"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_CurrentMonthDistribution"]);
                                        }
                                        else
                                        {
                                            //objPosition.ssi_currentmonthdistribution = new CrmDecimal();
                                            //objPosition.ssi_currentmonthdistribution.IsNull = true;
                                            //objPosition.ssi_currentmonthdistribution.IsNullSpecified = true;
                                            objPosition["ssi_currentmonthdistribution"] = null;

                                        }

                                        // LastLockDate
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_LastLockDate"]) != "")
                                        {
                                            //objPosition.ssi_lastlockdate = new CrmDateTime();
                                            //objPosition.ssi_lastlockdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_LastLockDate"]);
                                            objPosition["ssi_lastlockdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_LastLockDate"]);

                                        }

                                        if (Userid != "")
                                        {
                                            //objPosition.createdby = new Lookup();
                                            //objPosition.createdby.type = EntityName.systemuser.ToString();
                                            //objPosition.createdby.Value = new Guid(Userid);
                                            objPosition["createdby"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Convert.ToString(Userid)));
                                        }

                                        //SourceCode
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_DataSource"]) != "")
                                        {
                                            //objPosition.ssi_datasource = new Picklist();
                                            //objPosition.ssi_datasource.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_DataSource"]);
                                            objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_DataSource"]));

                                        }

                                        //CustodianId
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_CustodianId"]) != "")
                                        {
                                            //objPosition.ssi_custodianid = new Lookup();
                                            //objPosition.ssi_custodianid.type = EntityName.account.ToString();
                                            //objPosition.ssi_custodianid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_CustodianId"]));
                                            objPosition["ssi_custodianid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_CustodianId"])));

                                        }

                                        // Managed flag
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_ManagedCode"]) != "")
                                        {
                                            //objPosition.ssi_managedcode = new CrmBoolean();
                                            //objPosition.ssi_managedcode.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_ManagedCode"]).ToLower());
                                            objPosition["ssi_managedcode"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_ManagedCode"]).ToLower());

                                        }

                                        // ReportCode
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AdventReportCode"]) != "")
                                        {
                                           // objPosition.ssi_adventreportcode = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AdventReportCode"]);
                                            objPosition["ssi_adventreportcode"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AdventReportCode"]);
                                        }

                                        //price
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Price"]) != "")
                                        {
                                            //objPosition.ssi_price = new CrmMoney();
                                            //objPosition.ssi_price.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Price"]);
                                            objPosition["ssi_price"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_Price"]));

                                        }
                                        else
                                        {
                                            //objPosition.ssi_price = new CrmMoney();
                                            //objPosition.ssi_price.IsNull = true;
                                            //objPosition.ssi_price.IsNullSpecified = true;
                                            objPosition["ssi_price"] = null;
                                        }

                                        //Unreal G L
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_UnrealizedGL"]) != "")
                                        {
                                            //objPosition.ssi_unrealizedgl = new CrmMoney();
                                            //objPosition.ssi_unrealizedgl.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["ssi_UnrealizedGL"]);
                                            objPosition["ssi_unrealizedgl"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["ssi_UnrealizedGL"]));
                                        }
                                        else
                                        {
                                            //objPosition.ssi_unrealizedgl = new CrmMoney();
                                            //objPosition.ssi_unrealizedgl.IsNull = true;
                                            //objPosition.ssi_unrealizedgl.IsNullSpecified = true;
                                            objPosition["ssi_unrealizedgl"] = null;
                                        }

                                        //greshamadvised (sectorflg)
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_GreshamAdvised"]).ToLower() == "true")
                                        {
                                            //objPosition.ssi_greshamadvised = new CrmBoolean();
                                            //objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_GreshamAdvised"]).ToLower());
                                            objPosition["ssi_greshamadvised"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_GreshamAdvised"]).ToLower());

                                        }

                                        //Unit Adj Cost
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AdjustedUnitCost"]) != "")
                                        {
                                            //objPosition.ssi_adjustedunitcost = new CrmMoney();
                                            //objPosition.ssi_adjustedunitcost.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_AdjustedUnitCost"]);
                                            objPosition["ssi_adjustedunitcost"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_AdjustedUnitCost"]));
                                        }
                                        else
                                        {
                                            //objPosition.ssi_adjustedunitcost = new CrmMoney();
                                            //objPosition.ssi_adjustedunitcost.IsNull = true;
                                            //objPosition.ssi_adjustedunitcost.IsNullSpecified = true;
                                            objPosition["ssi_adjustedunitcost"] = null;
                                        }

                                        //Adjusted Cost
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TotalAdjustedCost"]) != "")
                                        {
                                            //objPosition.ssi_totaladjustedcost = new CrmMoney();
                                            //objPosition.ssi_totaladjustedcost.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_TotalAdjustedCost"]);
                                            objPosition["ssi_totaladjustedcost"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Ssi_TotalAdjustedCost"]));
                                        }
                                        else
                                        {
                                            //objPosition.ssi_totaladjustedcost = new CrmMoney();
                                            //objPosition.ssi_totaladjustedcost.IsNull = true;
                                            //objPosition.ssi_totaladjustedcost.IsNullSpecified = true;
                                            objPosition["ssi_totaladjustedcost"] = null;
                                        }

                                        // Saving flag
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Saving"]) != "")
                                        {
                                            //objPosition.ssi_saving = new CrmBoolean();
                                            //objPosition.ssi_saving.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Saving"]).ToLower());
                                            objPosition["ssi_saving"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_Saving"]).ToLower());

                                        }

                                        // Start date
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_StartDate"]) != "")
                                        {
                                            //objPosition.ssi_startdate = new CrmDateTime();
                                            //objPosition.ssi_startdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_StartDate"]);
                                            objPosition["ssi_startdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_StartDate"]);

                                        }

                                        // End date
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_EndDate"]) != "")
                                        {
                                            //objPosition.ssi_enddate = new CrmDateTime();
                                            //objPosition.ssi_enddate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_EndDate"]);
                                            objPosition["ssi_enddate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Ssi_EndDate"]);
                                        }

                                        // MTDPERF
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["SSI_PerfMTD"]) != "")
                                        {
                                            //objPosition.ssi_perfmtd = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SSI_PerfMTD"]);
                                            objPosition["ssi_perfmtd"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SSI_PerfMTD"]);
                                        }

                                        // Currency ( default to USD )
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["TransactionCurrencyId"]) != "")
                                        {
                                            //objPosition.transactioncurrencyid = new Lookup();
                                            //objPosition.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                            //objPosition.transactioncurrencyid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["TransactionCurrencyId"]));
                                            objPosition["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["TransactionCurrencyId"])));
                                        }

                                        //SubAssetclassid 
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SubAssetClassId"]) != "")
                                        {
                                            //objPosition.ssi_subassetclassid = new Lookup();
                                            //objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                                            //objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]));
                                            objPosition["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"])));
                                        }

                                        //SubAssetclassBenchmarkid 
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkId"]) != "")
                                        {
                                            //objPosition.ssi_benchmarkid = new Lookup();
                                            //objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                                            //objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkId"]));
                                            objPosition["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkId"])));
                                        }

                                        service.Create(objPosition);
                                        //Thread.Sleep(sleepTime);
                                        successCount = successCount + 1;
                                    }
                                    else
                                        break;
                                }
                               // catch (System.Web.Services.Protocols.SoapException exc)
                                    catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                                {
                                    bProceed = false;
                                    positionSuccess = false;
                                    failiureCount = failiureCount + 1;
                                    string failiureText = "Position Insert - Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountId"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AsofDate"]);
                                    strDescription = failiureText + " Error Detail:" + exc.Detail.Message;  //"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position " + failiureText + " Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 27, "HistoricalAccountCopy");
                                }
                                catch (Exception exc)
                                {
                                    bProceed = false;
                                    positionSuccess = false;
                                    failiureCount = failiureCount + 1;
                                    string failiureText = "Position Insert - Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountId"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AsofDate"]);
                                    strDescription = failiureText + " Error Detail: " + exc.Message;//"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position" + failiureText + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 27, "HistoricalAccountCopy");
                                }
                            }

                        sw.WriteLine("Insert Ends for Position on: " + DateTime.Now.ToString());

                        strDescription = "Total Position failed to insert: " + failiureCount;
                        LogMessage(sw, service, strDescription, 43, "HistoricalAccountCopy");

                        strDescription = "Total Position inserted: " + successCount;
                        LogMessage(sw, service, strDescription, 28, "HistoricalAccountCopy");
                        sw.WriteLine("---------------------------- New Position Insert Ends -------------------");
                        sw.WriteLine();
                        #endregion
            
                        #region AccountUpdate

                        successCount = 0;
                        failiureCount = 0;
                        totalCount = 0;

                        sw.WriteLine("---------------------------- Account Update Starts -------------------");
                        sw.WriteLine("Update Starts for Accounts on: " + DateTime.Now.ToString());
                        totalCount = ds_gresham.Tables[1].Rows.Count;
                        totalCountPositionAccount = totalCount;

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        Guid AccountID = new Guid(Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_accountid"]));

                                       // ssi_account objAccount = new ssi_account();
                                        Entity objAccount = new Entity("ssi_account");
                                        //objAccount.ssi_accountid = new Key();
                                        //objAccount.ssi_accountid.Value = AccountID;
                                        objAccount["ssi_accountid"] = new Guid(Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_accountid"]));


                                        //objAccount.ssi_name = Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]);
                                        objAccount["ssi_name"] = Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]);

                                        service.Update(objAccount);
                                        //Thread.Sleep(sleepTime);

                                        successCount = successCount + 1;
                                    }
                                    else
                                        break;
                                }
                                //catch (System.Web.Services.Protocols.SoapException exc)
                                catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)

                                {
                                    bProceed = false;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Account " + Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]) + " Error Detail: " + exc.Message + " " + exc.Detail.Message;
                                    LogMessage(sw, service, strDescription, 5, "HistoricalAccountCopy");
                                }
                                catch (Exception exc)
                                {
                                    bProceed = false;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Account : " + Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_name"]) + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 5, "HistoricalAccountCopy");
                                }
                            }
                        sw.WriteLine("Update Ends for Acoount on: " + DateTime.Now.ToString());
                        strDescription = "Total Accounts Failed to Update: " + failiureCount;
                        LogMessage(sw, service, strDescription, 31, "HistoricalAccountCopy");

                        strDescription = "Total Accounts Updated: " + successCount;
                        LogMessage(sw, service, strDescription, 4, "HistoricalAccountCopy");
                        sw.WriteLine("---------------------------- Account Update Ends  -------------------");
                        #endregion

                        #region ShowErrorFile

                        if (ds_gresham.Tables[2].Rows.Count > 0)
                        {
                            totalCountPositionError = ds_gresham.Tables[2].Rows.Count;
                            ViewState["PostnErrorDT"] = ds_gresham.Tables[2];
                            //lnkPositionErrorFile.Style.Add("display", "");
                            lnkTransactionErrorFile.Style.Add("display", "");
                        }
                        else
                        {
                            ViewState["PostnErrorDT"] = null;
                            //lnkPositionErrorFile.Style.Add("display", "none");
                        }

                        ds_gresham.Dispose();
                        dagersham.Dispose();
                        #endregion

                        #endregion
                        ////////////////////////////////// Insert Position  Ends//////////////////////////////////////////
                    }

                    ////////////////////////////////// Postion End ///////////////////////////////////////
                }
            }

            sw.Flush();
            sw.Close();

            #endregion
        }

        if (totalCountTransaction == 0 && totalCountTransactionAccount == 0 && totalCountTransactionError == 0 
            && totalCountPosition == 0 && totalCountPositionAccount == 0 && totalCountPositionError == 0)
            return false;
        else
            return true;
    }

    private static void LogMessage(StreamWriter sw, IOrganizationService service, string strDescription, int intIssueType, string strFileLoading)
    {
        try
        {
            sw.WriteLine(strDescription);

            //ssi_loadlog objLoadLog = new ssi_loadlog();
            Entity objLoadLog = new Entity("ssi_loadlog");
            //objLoadLog.ssi_name =Convert.ToString(DateTime.Today);

            //objLoadLog.ssi_date = new CrmDateTime();
            //objLoadLog.ssi_date.Value = DateTime.Now.ToString();
            objLoadLog["ssi_date"] = DateTime.Now;


           // objLoadLog.ssi_fileloading = strFileLoading;
            objLoadLog["ssi_fileloading"] = strFileLoading;

           // objLoadLog.ssi_descriptionofissue = strDescription;
            objLoadLog["ssi_descriptionofissue"] = strDescription;

            //objLoadLog.ssi_typeofissue = new Picklist();
            //objLoadLog.ssi_typeofissue.Value = intIssueType;
            objLoadLog["ssi_typeofissue"] = new Microsoft.Xrm.Sdk.OptionSetValue(intIssueType);



            service.Create(objLoadLog);
        }
        catch (Exception exc)
        {
            //HttpContext.Current.Response.Write(exc.Message.ToString());
            sw.WriteLine(exc.Message);
            //sw.Flush();
            //sw.Close();
            //throw;
        }
    }

    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        //string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        string strName = claimsIdentity.Name;


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

#region Commented OLD CODE CRM2016 UPGRADE
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
    //    //string username = WindowsIdentity.GetCurrent().Name;

    //    //if (username == "CORP\\gbhagia")
    //    //{
    //    //    // Use the global user ID of the system user that is to be impersonated.
    //    //    token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
    //    //    //token.CallerId = new Guid("C42C7E05-8303-DE11-A38C-001D09665E8F");//gary                
    //    //}
    //    if (CallerId != "")
    //        token.CallerId = new Guid(CallerId);
    //    CrmService service = new CrmService();

    //    service.Timeout = 3600000;
    //    if (crmServerUrl != null &&
    //        crmServerUrl.Length > 0)
    //    {
    //        UriBuilder builder = new UriBuilder(crmServerUrl);
    //        builder.Path = "//MSCRMServices//2007//CrmService.asmx";
    //        service.Url = builder.Uri.ToString();
    //    }

    //    service.CrmAuthenticationTokenValue = token;
    //    service.Credentials = System.Net.CredentialCache.DefaultCredentials;

    //    return service;
    //}
#endregion
    protected void lnkTransactionErrorFile_Click(object sender, EventArgs e)
    {
        #region Spire License Code
        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion
        try
        {
            String lsFileNamforFinalXls = "Historical_Acct_Load_Summary_" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);

            string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\Historical_Acct_Load_Summary.xlsx");

            string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls.Replace("xlsx", "xml"));
            string FilePath = (Server.MapPath("") + @"\ExcelTemplate\Historical_Acct_Load_Summary.xlsx");

            FileInfo loFile = new FileInfo(strDirectory1);
            loFile.CopyTo(strDirectory, true);

            DataSet ds = new DataSet();
            DataTable Transaction = new DataTable();
            DataTable Position = new DataTable();
            String sqlstr = string.Empty;

            if (ViewState["TrxnErrorDT"] != null && ViewState["PostnErrorDT"] != null)
            {
                Transaction = (DataTable)ViewState["TrxnErrorDT"];
                Position = (DataTable)ViewState["PostnErrorDT"];
                ds.Tables.Add(Transaction.Copy());
                ds.Tables[0].TableName = "Transaction";
                ds.AcceptChanges();
                ds.Tables.Add(Position.Copy());
                ds.Tables[1].TableName = "Position";
                ds.AcceptChanges();
            }
            else if (ViewState["TrxnErrorDT"] != null)
            {
                Transaction = (DataTable)ViewState["TrxnErrorDT"];
                ds.Tables.Add(Transaction.Copy());
                ds.Tables[0].TableName = "Transaction";
                ds.AcceptChanges();
            }
            else if (ViewState["PostnErrorDT"] != null)
            {
                Position = (DataTable)ViewState["PostnErrorDT"];
                ds.Tables.Add(Position.Copy());
                ds.Tables[0].TableName = "Position";
                ds.AcceptChanges();
            }
# region commented  13-04-2018 sasmit
        //    //export datatable to excel
        //    Workbook workbook = new Workbook();
        //    workbook.LoadFromFile(strDirectory, ExcelVersion.Version97to2003);
        //    for (int i = 0; i < ds.Tables.Count; i++)
        //    {
        //        Worksheet sheet = workbook.Worksheets[i];
        //        workbook.Version = ExcelVersion.Version97to2003;
        //        sheet.InsertDataTable(ds.Tables[i], true, 1, 1, -1, -1);
        //        sheet.Name = ds.Tables[i].TableName;

        //        sheet.AllocatedRange.AutoFitColumns();
        //        sheet.AllocatedRange.AutoFitRows();
        //        //sheet.Rows[0].RowHeight = 20;
        //    }

        //    workbook.SaveAsXml(strDirectory2);
        //    workbook = null;
        //    XmlDocument xmlDoc = new XmlDocument();
        //    xmlDoc.Load(strDirectory2);
        //    XmlElement businessEntities = xmlDoc.DocumentElement;
        //    XmlNode loNode = businessEntities.LastChild;
        //    XmlNode loNode1 = businessEntities.FirstChild;
        //    businessEntities.RemoveChild(loNode);


        //    xmlDoc.Save(strDirectory2);
        //    xmlDoc = null;
        //    loFile = null;
        //    loFile = new FileInfo(strDirectory);
        //    loFile.Delete();
        //    loFile = new FileInfo(strDirectory2);
        //    loFile.CopyTo(strDirectory, true);
        //    loFile = null;
        //    //loFile = new FileInfo(strDirectory2);
        //    //loFile.Delete();

        //    #region New xls to xlsx code
        //    Workbook workbook1 = new Workbook();
        //    workbook1.LoadFromXml(strDirectory2);
        //    workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);

        //   // workbook1 = new Workbook();
        //   //workbook1.LoadFromFile(strDirectory);
        //   // Worksheet sheet1 = workbook1.Worksheets[0];
        //   // sheet1.Range[6, 1, 6, 5].Style.Color = System.Drawing.Color.FromArgb(183, 221, 232);
        //   // sheet1.Range[7, 1, 500, 5].Style.Font.FontName = "Arial Unicode MS";

        //   // //workbook1.SaveToFile(strDirectory.Replace("xls", "xlsx"), ExcelVersion.Version2010);
        //   // workbook1.SaveToFile(strDirectory, ExcelVersion.Version2013);


        //    loFile = new FileInfo(strDirectory2);
        //    loFile.Delete();
        //    loFile = null;
        //    //  loFile = new FileInfo(strDirectory);
        //    ///   loFile.Delete();
        ////    lsFileNamforFinalXls = "/ExcelTemplate/TempFolder/" + lsFileNamforFinalXls;
            //    #endregion

#endregion
            int SheetNo = 0;
            Workbook book = new Workbook();
            book.CreateEmptySheets(2);
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                string SheetNme = ds.Tables[i].TableName.ToString();
                //string SheetNme = ds.Tables[i].Rows[0][0].ToString();
                //string GroupName = ds.Tables[i].Rows[0][1].ToString();
                //i++;
                //ds.Tables[i].Columns.Add("GroupName");
                //for (int k = 0; k < ds.Tables[i].Rows.Count; k++)
                //{
                //    ds.Tables[i].Rows[k]["GroupName"] = GroupName;
                //}


                Worksheet sheet = book.Worksheets[SheetNo];
                sheet.Name = SheetNme;
                if (ds.Tables[i].Rows.Count > 0)
                {
                    sheet.Range[1, 1, 1, ds.Tables[i].Columns.Count].Style.Font.IsBold = true;

                    sheet.InsertDataTable(ds.Tables[i], true, 1, 1);
                    sheet.Range[1, 1, ds.Tables[i].Rows.Count+1, ds.Tables[i].Columns.Count].AutoFitColumns();
                    sheet.Range[1, 1, ds.Tables[i].Rows.Count+1, ds.Tables[i].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;
                }
                SheetNo++;
            }

            book.SaveToFile(strDirectory, ExcelVersion.Version2016);

            #region PageSetup
            Workbook workbook1 = new Workbook();
            workbook1.LoadFromFile(strDirectory);
            Worksheet sheet1 = workbook1.Worksheets[0];
            workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);

            var setup = sheet1.PageSetup;
            setup.FitToPagesWide = 1;
            //setup.FitToPagesTall = 1;
            setup.IsFitToPage = true;
            setup.PaperSize = PaperSizeType.PaperA4;
            setup.Orientation = PageOrientationType.Landscape;
            setup.FitToPagesWide = 1;
            setup.FitToPagesTall = 0;
            setup.CenterHorizontally = true;
            setup.CenterVertically = false;
            workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);
            #endregion
            //lsFileNamforFinalXls = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strDirectory);
            //Response.ContentType = "application/octet-stream";
            //Response.AddHeader("Content-Disposition", "attachment;filename=Historical Acct Load Summary.xlsx");
            //Response.TransmitFile(strDirectory);
            //Response.End();

            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment;filename=" + Path.GetFileName(strDirectory) + "");
            Response.TransmitFile(strDirectory);
            Response.End();
        }
        catch (Exception ex)
        {
            strDescription = "Failed to throw Error file, Error Detail: " + ex.Message;
            lblError.Text = strDescription;
        }
    }
    protected void lnkPositionErrorFile_Click(object sender, EventArgs e)
    {
        try
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            String sqlstr = string.Empty;
            string filename = string.Empty;

            filename = "PositionExceptionFile";
            dt = (DataTable)ViewState["PostnErrorDT"];


            if (dt.Rows.Count > 0)
            {
                string attachment = "attachment; filename=" + filename + ".xls";
                Response.Clear();
                Response.ClearHeaders();
                Response.ClearContent();
                Response.AddHeader("content-disposition", attachment);
                Response.ContentType = "application/vnd.ms-excel";
                Response.ContentEncoding = System.Text.Encoding.Unicode;// GetEncoding("UTF-8");
                Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());
                string tab = "";
                foreach (DataColumn dc in dt.Columns)
                {
                    Response.Write(tab + dc.ColumnName);
                    tab = "\t";
                }
                Response.Write("\n");

                int i;

                foreach (DataRow dr in dt.Rows)
                {
                    tab = "";
                    for (i = 0; i < dt.Columns.Count; i++)
                    {
                        Response.Write(tab + dr[i].ToString());
                        tab = "\t";
                    }
                    Response.Write("\n");
                }
                Response.End();
            }
            else
            {
                lblError.Text = "No Records Found.";
            }

        }
        catch (Exception ex)
        {
            strDescription = "Failed to throw Error file, Error Detail: " + ex.Message;
            lblError.Text = strDescription;
        }
    }
}
