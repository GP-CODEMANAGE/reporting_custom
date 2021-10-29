using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.Configuration;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using System.Security.Principal;
using Microsoft.IdentityModel.Claims;
using System.Threading;
//using System.Security.Claims;

/// <summary>
/// Summary description for GeneralMethods
/// </summary>
/// 

public class GeneralMethods
{
    public String _ERRORMESSAGE;
    protected String _TransactReturn;
    protected String sqlstr;
    SqlConnection objcon;


    protected DB clsdb = new DB();

    public GeneralMethods()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    public bool IsInt(object MyObject)
    {
        try
        {
            int i = Convert.ToInt32(MyObject.ToString());
            return true;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    protected DB dbConn = new DB();

    //added 6/19/2017 CRM UPGRADE 2016 
    public IOrganizationService GetCrmService1()
    {
        ClientCredentials Credentials = new ClientCredentials();
        // Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

      //  Credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;

        Credentials.UserName.UserName = "corp\\crmadmin";
        Credentials.UserName.Password = "W!gmxF26ggw]";
        // string str = ((WindowsIdentity)HttpContext.Current.User.Identity).Name;
        IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        string str = claimsIdentity.Name;
        //This URL needs to be updated to match the servername and Organization for the environment.
        //  Uri OrganizationUri = new Uri("http://gp-crm2016/GreshamPartners/XRMServices/2011/Organization.svc");
        Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
        Uri HomeRealmUri = null;

        //OrganizationServiceProxy serviceProxy; 
        Microsoft.Xrm.Sdk.IOrganizationService service;
        Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy serviceProxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null);

        // This statement is required to enable early-bound type support.
        serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

        service = (Microsoft.Xrm.Sdk.IOrganizationService)serviceProxy;
        return service;
    }
    //public IOrganizationService GetCrmService()
    //{
    //    ClientCredentials Credentials = new ClientCredentials();
    //    // Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

    //    //  Credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;

    //     string str = ((WindowsIdentity)HttpContext.Current.User.Identity).Name;
    //  //  IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
    //   // string str = claimsIdentity.Name;
    //    string UserID = string.Empty;
    //    string sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + str + "'";
    //    DB clsDB = new DB();
    //    DataSet lodataset = clsDB.getDataSet(sqlstr);
    //    if (lodataset.Tables[0].Rows.Count > 0)
    //    {
    //        UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);

    //        // Response.Write("UserID:" + UserID);
    //        //return UserID = "DFCE21B1-B81E-E211-A2B7-0002A5443D86";
    //    }


    //    // *** for specific user credential *********  //
    //    Credentials.UserName.UserName = "corp\\crmadmin";
    //    Credentials.UserName.Password = "W!gmxF26ggw]";




    //    // Credentials.UserName.UserName = "corp\\crmadmin";
    //    // Credentials.UserName.Password = "W!gmxF26ggw]";

    //    //This URL needs to be updated to match the servername and Organization for the environment.
    //    //  Uri OrganizationUri = new Uri("http://gp-crm2016/GreshamPartners/XRMServices/2011/Organization.svc");
    //    Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
    //    Uri HomeRealmUri = null;



    //    //OrganizationServiceProxy serviceProxy; 
    //    Microsoft.Xrm.Sdk.IOrganizationService service;
    //    Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy serviceProxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null);

    //    // This statement is required to enable early-bound type support.
    //    serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

    //    service = (Microsoft.Xrm.Sdk.IOrganizationService)serviceProxy;
    //    Guid _UserID = new Guid(UserID);

    //    serviceProxy.CallerId = _UserID;

    //    return service;
    //}
    public IOrganizationService GetCrmService()
    {
        ClientCredentials Credentials = new ClientCredentials();
        // Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

        //  Credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;

        //string str = ((WindowsIdentity)HttpContext.Current.User.Identity).Name;
        string str = string.Empty;
        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            str = "corp\\skane";
        }
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            str = claimsIdentity.Name;

        }

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
        Microsoft.Xrm.Sdk.IOrganizationService service;
        Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy serviceProxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null);

        // This statement is required to enable early-bound type support.
        serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

        service = (Microsoft.Xrm.Sdk.IOrganizationService)serviceProxy;
        Guid _UserID = new Guid(UserID);

        serviceProxy.CallerId = _UserID;

        return service;
    }
    public IOrganizationService ConnectToMSCRMOnline(string UserName, string Password, string SoapOrgServiceUri)
    {
        try
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            ClientCredentials credentials = new ClientCredentials();
            credentials.UserName.UserName = UserName;
            credentials.UserName.Password = Password;
            Uri serviceUri = new Uri(SoapOrgServiceUri);
            OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
            proxy.EnableProxyTypes();
            IOrganizationService service = (IOrganizationService)proxy;
            return service;

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error while connecting to CRM " + ex.Message);
            Console.ReadKey();
            return null;
        }
    }
    // added by abhijeet on July 23rd
    public void getListForBindListBox(ListBox lstControlName, string sqlstr, string DataTextField, string DataValField)
    {
        try
        {
            DataSet DS;
            objcon = clsdb.gOpenConnection();
            DS = clsdb.getDataSet(sqlstr);
            if (DS.Tables[0].Rows.Count > 0)
            {
                DataTable dt = new DataTable();
                dt = DS.Tables[0];
                lstControlName.DataSource = dt;
                lstControlName.DataTextField = DataTextField;
                lstControlName.DataValueField = DataValField;
                lstControlName.DataBind();
            }
            else
                lstControlName.Items.Clear();

            
        }
        finally
        {
            if (objcon != null)
                clsdb.gCloseConnection(objcon);
        }
    }

    /// <summary>
    /// Execute Storeprocedure and bind the generated DataTable to DropdownList
    /// with respect to provided DataTextField and DataValField with First value Selected Always.
    /// </summary>
    /// <param name="ddlControlName">DropDownList</param>
    /// <param name="sqlstr">string</param>
    /// <param name="DataTextField">string</param>
    /// <param name="DataValField">string</param>
    public void getListForBindDDL(DropDownList ddlControlName, string sqlstr, string DataTextField, string DataValField)
    {
        try
        {
            ddlControlName.Items.Clear();
            DataSet DS;
            objcon = clsdb.gOpenConnection();
            DS = clsdb.getDataSet(sqlstr);
            if (DS.Tables[0].Rows.Count > 0)
            {
                DataTable dt = new DataTable();
                dt = DS.Tables[0];
                ddlControlName.DataSource = dt;
                ddlControlName.DataTextField = DataTextField;
                ddlControlName.DataValueField = DataValField;
                ddlControlName.DataBind();
            }
            else
            {
                ddlControlName.SelectedIndex = ddlControlName.SelectedIndex - 1;
                ddlControlName.Items.Insert(0, "No Record Found");
                ddlControlName.Items[0].Value = "0";
                ddlControlName.SelectedIndex = 0;
            }
        }
        catch (Exception ex)
        {
        }
        finally
        {
            if (objcon != null)
                clsdb.gCloseConnection(objcon);
        }
    }

    /// <summary>
    /// Execute Storeprocedure and bind the generated DataTable to DropdownList
    /// with respect to provided DataTextField and DataValField with Select Option.
    /// </summary>
    /// <param name="ddlControlName">DropDownList</param>
    /// <param name="sqlstr">string</param>
    /// <param name="DataTextField">string</param>
    /// <param name="DataValField">string</param>
    public void getBindDDL(DropDownList ddlControlName, string sqlstr, string DataTextField, string DataValField)
    {
        DataSet DS;
        objcon = clsdb.gOpenConnection();
        try
        {
            DS = clsdb.getDataSet(sqlstr);
            if (DS.Tables[0].Rows.Count > 0)
            {
                DataTable dt = new DataTable();
                dt = DS.Tables[0];
                ddlControlName.DataSource = dt;
                ddlControlName.DataTextField = DataTextField;
                ddlControlName.DataValueField = DataValField;
                ddlControlName.DataBind();
                ddlControlName.Items.Insert(0, "Select");
                ddlControlName.Items[0].Value = "0";
                ddlControlName.SelectedIndex = 0;
            }
            else
            {
                ddlControlName.SelectedIndex = ddlControlName.SelectedIndex - 1;
                ddlControlName.Items.Insert(0, "Not Specified");
                ddlControlName.Items[0].Value = "0";
                ddlControlName.SelectedIndex = 0;
            }
        }
        finally
        {
            if (objcon != null)
                clsdb.gCloseConnection(objcon);
        }
    }

    public void getStateList(DropDownList drpState)
    {
        DataSet DS;
        objcon = clsdb.gOpenConnection();
        try
        {
            sqlstr = " SP_S_state_lkup ";
            DS = clsdb.getDataSet(sqlstr);
            if (DS.Tables[0].Rows.Count > 0)
            {
                DataTable dt = new DataTable();
                dt = DS.Tables[0];
                drpState.DataSource = dt;
                drpState.DataTextField = "NameTxt";
                drpState.DataValueField = "IdNmb";
                drpState.DataBind();
                drpState.Items.Insert(0, "Select");
                drpState.SelectedIndex = 0;
            }
            else
                drpState.SelectedIndex = drpState.Items.Count - 1;
        }
        finally
        {
            if (objcon != null)
                clsdb.gCloseConnection(objcon);
        }
    }

    // addded by abhijeet on July 25th
    // for common listing on any sql select string
    public DataTable getList(string sqlstr)
    {
        DataSet DS;
        DataTable DT = new DataTable();

        DS = clsdb.getDataSet(sqlstr);
        if (DS.Tables[0].Rows.Count > 0)
        {
            DT = DS.Tables[0];
        }
        return DT;
    }

    // Code bind Dropdown list of page 
    public void DataBoundList(GridViewRow gvrPager, GridView gvList)
    {
        {
            if (gvrPager == null) return;

            // get your controls from the gridview
            DropDownList ddlPages = (DropDownList)gvrPager.Cells[0].FindControl("ddlPages");
            System.Web.UI.WebControls.Label lblPageCount = (System.Web.UI.WebControls.Label)gvrPager.Cells[0].FindControl("lblPageCount");

            if (ddlPages != null)
            {
                // populate pager
                for (int i = 0; i < gvList.PageCount; i++)
                {

                    int intPageNumber = i + 1;
                    ListItem lstItem = new ListItem((intPageNumber.ToString() + " of " + gvList.PageCount.ToString()), intPageNumber.ToString());

                    if (i == gvList.PageIndex)
                        lstItem.Selected = true;

                    ddlPages.Items.Add(lstItem);
                }
            }
        }
    }

    #region to sort gridview

    public void SortGridView(GridView gvSort, string sqlstrRpt, string sortExpression, string Direction)
    {
        int index = 0;
        DataTable dtTbl = getList(sqlstrRpt);
        DataView dv = new DataView(dtTbl);
        if (sortExpression != "" && Direction != "" && dtTbl.Rows.Count > 0)
        {
            dv.Sort = sortExpression + " " + Direction;
            foreach (DataControlField field in gvSort.Columns)
            {
                gvSort.Columns[gvSort.Columns.IndexOf(field)].HeaderStyle.CssClass = "unsorted";
                if (field.SortExpression == sortExpression)
                    index = gvSort.Columns.IndexOf(field);
            }
        }
        gvSort.DataSource = dv;
        gvSort.DataBind();
        gvSort.Columns[index].HeaderStyle.CssClass = Direction;


    }
    #endregion
    #region Move Up and down

    public void MoveUp(ListBox lstBox)
    {
        int iIndex, iCount, iOffset, iInsertAt, iIndexSelectedMarker = -1;
        string lItemData, lItemval;
        // Get the count of items in the list control
        iCount = lstBox.Items.Count;

        // Set the base loop index and the increment/decrement value based
        // on the direction the item are being moved (up or down).
        iIndex = 0;
        iOffset = -1;
        // Loop through all of the items in the list.
        while (iIndex < iCount)
        {
            // Check if this item is selected.
            if (lstBox.SelectedIndex > 0)
            {
                // Get the item data for this item
                lItemval = lstBox.SelectedItem.Value.ToString();
                lItemData = lstBox.SelectedItem.Text.ToString();
                iIndexSelectedMarker = lstBox.SelectedIndex;

                // Don't move selected items past other selected items
                if (-1 != iIndexSelectedMarker)
                {
                    for (int iIndex2 = 0; iIndex2 < iCount; ++iIndex2)
                    {
                        // Find the index of this item in enabled list
                        if (lItemval == lstBox.Items[iIndex2].Value.ToString())
                        {

                            // Remove the item from its current position
                            lstBox.Items.RemoveAt(iIndex2);

                            // Reinsert the item in the array one space higher 
                            // than its previous position
                            iInsertAt = (iIndex2 + iOffset) < 0 ? 0 : iIndex2 + iOffset;
                            ListItem li = new ListItem(lItemData, lItemval);
                            lstBox.Items.Insert(iInsertAt, li);
                            break;
                        }
                    }
                }
            }

                       // If this item wasn't selected save the index so we can check
            // it later so we don't move past the any selected items.
            else if (-1 == iIndexSelectedMarker)
            {
                iIndexSelectedMarker = iIndex;
                break;
            }
            iIndex = iIndex + 1;
        }
        if (iIndexSelectedMarker == 0)
            lstBox.SelectedIndex = iIndexSelectedMarker;
        else
            lstBox.SelectedIndex = iIndexSelectedMarker - 1;

    }

    public void MoveDown(ListBox lstBox)
    {
        int iIndex, iCount, iOffset, iInsertAt, iIndexSelectedMarker = -1;
        string lItemData;
        string lItemval;

        // Get the count of items in the list control
        iCount = lstBox.Items.Count;

        // Set the base loop index and the increment/decrement value based on 
        // the direction the item are being moved (up or down).
        iIndex = iCount - 1;
        iOffset = 1;

        // Loop through all of the items in the list.
        while (iIndex >= 0)
        {

            // Check if this item is selected.
            if (lstBox.SelectedIndex >= 0)
            {

                // Get the item data for this item
                lItemData = lstBox.SelectedItem.Text.ToString();
                lItemval = lstBox.SelectedItem.Value.ToString();
                iIndexSelectedMarker = lstBox.SelectedIndex;

                // Don't move selected items past other selected items
                if (-1 != iIndexSelectedMarker)
                {
                    for (int iIndex2 = 0; iIndex2 < iCount - 1; ++iIndex2)
                    {
                        // Find the index of this item in enabled list
                        if (lItemval == lstBox.Items[iIndex2].Value.ToString())
                        {
                            // Remove the item from its current position
                            lstBox.Items.RemoveAt(iIndex2);
                            // Reinsert the item in the array one space lower 
                            // than its previous position
                            iInsertAt = (iIndex2 + iOffset) < 0 ? 0 : iIndex2 + iOffset;
                            ListItem li = new ListItem(lItemData, lItemval);
                            lstBox.Items.Insert(iInsertAt, li);
                            break;
                        }
                    }
                }
            }
            iIndex = iIndex - 1;
        }
        if (iIndexSelectedMarker == lstBox.Items.Count - 1)
            lstBox.SelectedIndex = iIndexSelectedMarker;
        else
            lstBox.SelectedIndex = iIndexSelectedMarker + 1;

    }

    #endregion

    public string getVersion()
    {
        DataSet DS = new DataSet();
        sqlstr = "sp_s_APP_VERSION_LOG_top";
        DS = clsdb.getDataSet(sqlstr);
        return (DS.Tables[0].Rows[0]["VersionTxt"].ToString());
    }

    // added by Abhijeet Khake for Move Listbox items
    public void MoveListBoxItems(ListBox lstBoxSource, ListBox lstBoxTarget)
    {
        int itmCount = lstBoxSource.Items.Count;
        for (int i = 0; i < itmCount; i++)
        {
            if (lstBoxSource.Items[i].Selected)
            {
                lstBoxTarget.Items.Add(lstBoxSource.Items[i]);
                lstBoxSource.Items.Remove(lstBoxSource.Items[i]);
                itmCount--;
                i--;
            }
        }
    }

    public void getType(DropDownList drpComp, DataTable GlobalDT, String IdNmb, String NameTxt)
    {
        drpComp.DataSource = GlobalDT;
        drpComp.DataTextField = NameTxt;//"EmpNameTxt";
        drpComp.DataValueField = IdNmb; //"EmpIdNmb";
        drpComp.DataBind();
        drpComp.Items.Insert(0, "Select");
        drpComp.Items[0].Value = "Select";
        drpComp.SelectedIndex = 0;
    }
    public string GetNote(string ClientIdNmb, string ScheduleNmb, string EndDT)
    {
        DataTable dt = new DataTable();
        string Note = string.Empty;
        EndDT = Convert.ToDateTime(EndDT).ToString("MM/dd/yyyy");
        sqlstr = "SP_R_FOOTNOTE_ETY @ClientIdNmb=" + ClientIdNmb + ", @ScheduleNmb  = " + ScheduleNmb + ", @EndDT  = '" + EndDT + "'";
        dt = getList(sqlstr);
        if (dt.Rows.Count > 0)
            for (int i = 0; i < dt.Rows.Count; i++)
                Note = Note + "Note: " + Convert.ToString(dt.Rows[i]["FootNoteTxt"]) + "<br/>";
        return Note;
    }


    //Added By Rohit Pawar
    //Purpose : To Get Multiple Selected Item value from listbox
    public string GetMultipleSelectedItemsFromListBox(ListBox lstBox)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 0; i < lstBox.Items.Count; i++)
            {
                if (lstBox.Items[i].Selected)
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

    //Added By Harshit on 25th Feb 2016
    //Purpose : To Get All Item's VALUE from listbox
    public string GetALLValuesFromListBox(ListBox lstBox)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 1; i < lstBox.Items.Count; i++)
            {
                lstselecteditems = lstselecteditems + ", " + lstBox.Items[i].Value;
                //insert command
            }
            if (lstselecteditems != "")
            {
                lstselecteditems = lstselecteditems.Substring(1);
            }
        }


        return lstselecteditems;
    }

    //Added By Harshit on 11th June 2015
    //Purpose : To Get the Text of the Multiple Selected Items of listbox
    public string GetMultipleSelectedItemsTEXTFromListBox(ListBox lstBox)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 0; i < lstBox.Items.Count; i++)
            {
                if (lstBox.Items[i].Selected)
                {
                    lstselecteditems = lstselecteditems + ", " + lstBox.Items[i].Text;
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
    //Added By Harshit on 11th June 2015
    //Purpose : To Get the Text of the All the Items of listbox
    public string GetALLItemsTEXTFromListBox(ListBox lstBox)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 0; i < lstBox.Items.Count; i++)
            {
                lstselecteditems = lstselecteditems + ", " + lstBox.Items[i].Text;
                //insert command
            }
            if (lstselecteditems != "")
            {
                lstselecteditems = lstselecteditems.Substring(1);
            }
        }


        return lstselecteditems;
    }

    static public string RemoveSpecialCharacters(string str)
    {

        //StringBuilder sb = new StringBuilder();
        //foreach (char c in str) 
        //{
        //   if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') | c == '.' || c == '_') 
        //   {
        //      sb.Append(c);
        //   }
        //}

        // return sb.ToString();

        System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex("[;\\/:*?\"<>|&']");
        string outputString = re.Replace(str, " ");

        outputString = outputString.Replace("(Q", " - Q").Replace("(M", " - M").Replace("(MTGBK", " - MTGBK").Replace("Q)", "Q").Replace("M)", "M"); ; //added 2_1_2019 Non Marketable (DYNAMO)
        //outputString = outputString.Replace("(Q", " - Q").Replace("(M", " - M").Replace("(MTGBK", " - MTGBK").Replace(")", "");
        return outputString;
    }

    public string CreateRandomNumber(int PasswordLength)
    {
        string _allowedChars = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNOPQRSTUVWXYZ0123456789";
        Random randNum = new Random();
        char[] chars = new char[PasswordLength];
        int allowedCharCount = _allowedChars.Length;

        for (int i = 0; i < PasswordLength; i++)
        {
            chars[i] = _allowedChars[(int)((_allowedChars.Length) * randNum.NextDouble())];
        }

        return new string(chars);
    }

}
