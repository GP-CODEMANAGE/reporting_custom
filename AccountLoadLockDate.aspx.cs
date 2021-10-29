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
using System.Configuration;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;

using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class AccountLoadLockDate : System.Web.UI.Page
{
    string sqlstr = string.Empty;
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    public StreamWriter sw = null;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            FillSource();
            trDownLoad.Style.Add("display", "none");
        }
    }
    private void FillSource()
    {
        string sqlstr = "EXEC SP_S_ACCOUNT_SOURCE_LKUP";
        clsGM.getListForBindListBox(lstSource, sqlstr, "ssi_name", "ssi_source");
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        DateTime value;
        bool IsProperDate = true;
        if (txtAsofdate.Text != "")
        {
            if (!DateTime.TryParse(txtAsofdate.Text, out value))
                IsProperDate = false;
        }
       if (!DateTime.TryParse(txtLockDate.Text, out value))
           IsProperDate = false;
       if (IsProperDate == false)
       {
           System.Text.StringBuilder sb = new System.Text.StringBuilder();
           Type tp = this.GetType();
           sb.Append("\n<script type=text/javascript>\n");
           sb.Append("\n alert('Please provide proper date.');");
           sb.Append("</script>");
           ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
           return;
       }
       else
       {
           LockData();
       }
    }

    private void LockData()
    {
        #region Declaration
        string LogFileName = "AccountLockLogFile " + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        sw = new StreamWriter(Request.PhysicalApplicationPath + "\\Log\\" + LogFileName + ".txt", true);

        // string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
        //string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);

        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        string orgName = "GreshamPartners";
        bool bProceed = true;
        string strDescription;
        string Userid = GetcurrentUser();
        IOrganizationService service = null;

        try
        {
            service =clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            LogMessage(sw, service, strDescription, 62, "GeneralError");
            sw.WriteLine("step 1 ");
        }
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

       // service.PreAuthenticate = true;
       //
       // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);

        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();

        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery;
        int totalCount = 0;
        int successCount = 0;
        int failiureCount = 0;

        #endregion

        if (bProceed == true)
        {
            // using data load from file
            #region Account Lock Data


            successCount = 0;
            failiureCount = 0;
            // bProceed = true;

            if (bProceed == true)
            {
                successCount = 0;
                failiureCount = 0;
                totalCount = 0;

                #region Lock Account
                try
                {
                    sw.WriteLine("---------------------------- Account lock Starts -------------------");
                    string strSource = clsGM.GetMultipleSelectedItemsFromListBox(lstSource);
                    string AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
                    string LockDate = txtLockDate.Text;
                    greshamquery = "SP_U_ACCOUNT_LOCKDATE " +
                                   "@AccountSourceIdNmb='" + strSource + "'" +
                                   ",@AsOfdate=" + AsOfDate + "" +
                                   ",@LockDt='" + LockDate + "'";
                    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                    ds_gresham = new DataSet();
                    dagersham.SelectCommand.CommandTimeout = 1800;
                    dagersham.Fill(ds_gresham);
                    totalCount = ds_gresham.Tables[0].Rows.Count;
                    sw.WriteLine("Account lock Starts on: " + DateTime.Now.ToString());
                }
                catch (System.Web.Services.Protocols.SoapException exc)
                {
                    bProceed = true;
                    totalCount = 0;
                    strDescription = "Account lock failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                    LogMessage(sw, service, strDescription, 62, "Account lock");
                }
                catch (Exception exc)
                {
                    bProceed = true;
                    totalCount = 0;
                    strDescription = "Account lock failed, please contact administrator. Error Detail: " + exc.Message;
                    LogMessage(sw, service, strDescription, 62, "Account lock");
                }

                if (bProceed == true)
                    for (int i = 0; i < totalCount; i++)
                    {
                        try
                        {
                            if (bProceed == true)
                            {
                              //  ssi_account objAccount = new ssi_account();

                                Microsoft.Xrm.Sdk.Entity objAccount = new Microsoft.Xrm.Sdk.Entity("ssi_account");
                                
                                Guid AccountId = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_Accountid"]));

                              //  Guid accountid = new Guid(HouseholdId.Replace("'", ""));

                                ///objAccount["accountid"] = accountid;

                                //objAccount.ssi_accountid = new Key();
                                objAccount["ssi_accountid"] = AccountId;

                                //Data Load Lock date
                               // objAccount.ssi_loadlockdt = new CrmDateTime();
                               // objAccount.ssi_loadlockdt.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_LoadLockDT"]);

                                objAccount["ssi_loadlockdt"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["ssi_LoadLockDT"]);

                                service.Update(objAccount);
                                //Thread.Sleep(sleepTime);

                                successCount = successCount + 1;
                            }
                            else
                                break;
                        }
                        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                        {
                            bProceed = false;
                            failiureCount = failiureCount + 1;
                            string failiureText = "AccountId:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_Accountid"]);
                            strDescription = failiureText + " Error Detail:" + exc.Detail.Message;
                            LogMessage(sw, service, strDescription, 5, "Account lock");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            failiureCount = failiureCount + 1;
                            string failiureText = "AccountId:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_Accountid"]);
                            strDescription = failiureText + " Error Detail:" + exc.Message;
                            LogMessage(sw, service, strDescription, 5, "Account lock");
                        }
                    }

                sw.WriteLine("Account Lock Ends on: " + DateTime.Now.ToString());

                strDescription = "Total Account failed to Lock: " + failiureCount;
                LogMessage(sw, service, strDescription, 31, "Account lock");

                strDescription = "Total Account Locked: " + successCount;
                LogMessage(sw, service, strDescription, 4, "Account lock");
                sw.WriteLine("---------------------------- Account lock Ends -------------------");
                sw.WriteLine();

                if (totalCount > 0 && bProceed == true)
                {
                    //Session["AccDS"] = ds_gresham;
                    trDownLoad.Style.Add("display", "inline");
                }
                else
                {
                    //Session["AccDS"] = null;
                    trDownLoad.Style.Add("display", "none");
                }


                ds_gresham.Dispose();
                dagersham.Dispose();

                lblError.Text = "Total Account failed to Update: " + failiureCount + 
                                "</br>Total Account Updated: " + successCount;

                #endregion

            }

            sw.Flush();
            sw.Close();

            #endregion

        }

    }

    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;

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

    /// <summary>
    /// Set up the CRM Service.
    /// </summary>
    /// <param name="organizationName">My Organization</param>
    /// <returns>CrmService configured with AD Authentication</returns>
    /// 
    #region Old Code 
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

    private static void LogMessage(StreamWriter sw,  IOrganizationService  service, string strDescription, int intIssueType, string strFileLoading)
    {
        try
        {
            sw.WriteLine(strDescription);

          //  ssi_loadlog objLoadLog = new ssi_loadlog();

            Microsoft.Xrm.Sdk.Entity objLoadLog = new Microsoft.Xrm.Sdk.Entity("ssi_loadlog");

            //objLoadLog.ssi_name =Convert.ToString(DateTime.Today);

           // objLoadLog.ssi_date = new CrmDateTime();
           // objLoadLog.ssi_date.Value = DateTime.Now.ToString();

            objLoadLog["ssi_date"] = DateTime.Now;

            objLoadLog["ssi_fileloading"] = strFileLoading;

            objLoadLog["ssi_descriptionofissue"] = strDescription;

           //objLoadLog.ssi_typeofissue = new Picklist();
           //objLoadLog.ssi_typeofissue.Value = intIssueType;

            objLoadLog["ssi_typeofissue"] = new Microsoft.Xrm.Sdk.OptionSetValue(intIssueType);

            service.Create(objLoadLog);
        }
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            //HttpContext.Current.Response.Write(exc.Message.ToString());
            sw.WriteLine(exc.Message);
            sw.Flush();
            sw.Close();
            throw;
        }

    }
    protected void lnkDownLoad_Click(object sender, EventArgs e)
    {
        String lsFileNamforFinalXls = "LockedAccount_" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xls";
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls);

        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\NewAccountAndSecurity.xls");

        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls.Replace("xls", "xml"));
        string FilePath = (Server.MapPath("") + @"\ExcelTemplate\NewAccountAndSecurity.xls");

        FileInfo loFile = new FileInfo(strDirectory1);
        loFile.CopyTo(strDirectory, true);

        DataSet ds = new DataSet();

        string strSource = clsGM.GetMultipleSelectedItemsFromListBox(lstSource);
        string AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
        string LockDate = txtLockDate.Text;
        string query = "SP_U_ACCOUNT_LOCKDATE " +
                       "@AccountSourceIdNmb='" + strSource + "'" +
                       ",@AsOfdate=" + AsOfDate + "" +
                       ",@LockDt='" + LockDate + "'";
        ds = clsDB.getDataSet(query);
        //ds = (DataSet)Session["AccDS"];
        if (ds.Tables.Count > 0)
        {
            ds.Tables[0].Columns[0].Caption = "Account";
            ds.Tables[0].Columns[2].Caption = "Data Load Lock Date";
            ds.Tables[0].Columns.Remove("ssi_Accountid");
            ds.AcceptChanges();
            //export datatable to excel
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(strDirectory);
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                workbook.Version = ExcelVersion.Version97to2003;
                sheet.InsertDataTable(ds.Tables[i], true, 1, 1, -1, -1);
                sheet.Name = ds.Tables[i].TableName;

                sheet.AllocatedRange.AutoFitColumns();
                sheet.AllocatedRange.AutoFitRows();
                //sheet.Rows[0].RowHeight = 20;
            }

            workbook.SaveAsXml(strDirectory2);
            workbook = null;
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(strDirectory2);
            XmlElement businessEntities = xmlDoc.DocumentElement;
            XmlNode loNode = businessEntities.LastChild;
            XmlNode loNode1 = businessEntities.FirstChild;
            businessEntities.RemoveChild(loNode);


            xmlDoc.Save(strDirectory2);
            xmlDoc = null;
            loFile = null;
            loFile = new FileInfo(strDirectory);
            loFile.Delete();
            loFile = new FileInfo(strDirectory2);
            loFile.CopyTo(strDirectory, true);
            loFile = null;
            loFile = new FileInfo(strDirectory2);
            loFile.Delete();

            lsFileNamforFinalXls = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls);
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment;filename=LockedAccount.xls");
            Response.TransmitFile(lsFileNamforFinalXls);
            Response.End();
        }
        else
        {
            lblError.Text = "No records found.";
        }
    }
}
