using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Data.SqlClient;
using System.Data;
//using System.Reflection.PropertyInfo;
using System.Reflection;
using System.Configuration;
using System.Collections;
using System.Globalization;
using System.Text;
using System.Net.Mail;
using Microsoft.SharePoint.Client.Taxonomy;


public partial class frmAnzanioFilecopy : System.Web.UI.Page
{

    DataTable dt = new DataTable();
    DataTable dtclients = new DataTable();
    DataTable dt2 = new DataTable();
    DataRow rowRootFolder = null;

    DataRow rowRootFolder1 = null;//1st grid
    DataRow rowRootFolder2 = null;//2nd grid
    DataTable dt3 = new DataTable();
    DataTable fileCopyList = new DataTable();
    DataSet ds = new DataSet();
    DB clsDB = new DB();
    sharepoint sh = new sharepoint();
    DataSet dstaxonomy = new DataSet();
    DataTable FolderData = new DataTable();
    DataTable dtMail = new DataTable();

    DataTable dtMail1 = new DataTable();
    DataTable dtSiteClientList;
    DataTable year = new DataTable();

    List<string> anzianoIds = new List<string>();

    DataTable FolderPath = new DataTable();
    int count = 0;
    int Success = 0;

    int anz;
    string FileNAME = string.Empty;
    ClientContext context;

    private StringBuilder emailBody = new StringBuilder("<html>");
    string text = string.Empty;

    string vSourcePath = string.Empty;
    //string ListEmail = AppLogic.GetParam(AppLogic.ConfigParam.ToEmailIDs);
    // string ListEmail = "DBailey@greshampartners.com;JMasa@greshampartners.com;JScalise@greshampartners.com";
   
        //commented on 12_4_2018 Jscalise nolonger in process
   // string ListEmail = "JScalise@greshampartners.com";
   // string ListEmail = "skane@infograte.com";
    string AnziID = string.Empty;
    DataTable dtSharepint = new DataTable();
    DataTable dtSharepint1 = new DataTable();
    DataTable dtSharepint2 = new DataTable();

    DataTable dtCLIENT = new DataTable();




    public void newFolderstructure()
    {
       // string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
        context = new ClientContext(siteUrl);

        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;
        List list = context.Web.Lists.GetByTitle("Fund Admin");

        context.Load(list);
        context.Load(list.RootFolder);
        context.Load(list.RootFolder.Folders);
        context.Load(list.RootFolder.Files);
        context.ExecuteQuery();
        FolderCollection fcol = list.RootFolder.Folders;
        //  List<string> lstFile = new List<string>()
        List<string> lstFile = new List<string>();

        dt.Columns.Add("BatchName");
        // DataTable dt1 = lstFile.tablecopy.ToDataTable();
        foreach (Folder f in fcol)
        {
            //Response.Write("<br/>," + f.Name.ToString());
            dt.Rows.Add(f.Name.ToString());


            // ddlFolderName.Items.Add(f.Name.ToString());


            //if (f.Name == "Gerst")
            //{
            //    cxt.Load(f.Files);
            //    cxt.ExecuteQuery();
            //    FileCollection fileCol = f.Files;
            //    foreach (File file in fileCol)
            //    {
            //        lstFile.Add(file.Name);
            //    }
            //}
        }

        dt.DefaultView.Sort = "BatchName ASC";

        ddlFolderName.DataSource = dt;
        ddlFolderName.DataTextField = "BatchName";

        ddlFolderName.DataValueField = "BatchName";
        ddlFolderName.DataBind();

        ddlFolderName.Items.Insert(0, "Select a Folder");
        ddlFolderName.Items[0].Value = "0";
    }


    #region getdata from sharepoint and bind


    //public void newFolderstructuretemp()
    //{
    //    string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
    //    ClientContext context = new ClientContext(siteUrl);
    //    SecureString passWord = new SecureString();
    //    foreach (var c in "w!ldWind36") passWord.AppendChar(c);
    //    context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
    //    Web site = context.Web;

    //    Folder subFoldercol = site.GetFolderByServerRelativeUrl("Documents/Adams");
    //    ListCollection collList = site.Lists;

    //    FolderCollection fcolection = subFoldercol.Folders;
    //    context.Load(fcolection);
    //    context.Load(collList);
    //    context.ExecuteQuery();

    //    // DataTable dtSharepint = new DataTable();
    //    dtSharepint.Columns.Add("Rootfolder");
    //    dtSharepint.Columns.Add("SubFolder");
    //    dtSharepint.Columns.Add("Path");

    //    foreach (Folder f in fcolection)
    //    {
    //        //Response.Write("<br/>");
    //        //Response.Write("<br/>," + f.Name.ToString());
    //        DataRow rowRootFolder = dtSharepint.NewRow();
    //        rowRootFolder["Rootfolder"] = f.Name.ToString();
    //        rowRootFolder["Path"] = "/" + f.Name.ToString();
    //        dtSharepint.Rows.Add(rowRootFolder);

    //        string foldername = "Documents/Adams/" + f.Name;

    //        List<string> folderlist = getFolderPath(site, context, foldername, dtSharepint);

    //        foreach (string FolderName in folderlist)
    //        {
    //            DataRow rowRootFolder1 = dtSharepint.NewRow();
    //            rowRootFolder1["SubFolder"] = FolderName;
    //            //dt.Rows.Add(rowRootFolder1);


    //        }
    //        //    Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);
    //        ////    ListCollection collList = site.Lists;

    //        //    FolderCollection fcolection1 = subFolder1.Folders;
    //        //    context.Load(fcolection1);
    //        //   // context.Load(collList);
    //        //    context.ExecuteQuery();

    //        //    foreach (Folder subfolder in fcolection1)
    //        //    {
    //        //        //Response.Write("<br/>," + subfolder.Name.ToString());
    //        //        DataRow rowRootFolder1 = dt.NewRow();
    //        //        rowRootFolder1["SubFolder"] = subfolder.Name.ToString();
    //        //        dt.Rows.Add(rowRootFolder1);
    //        //    }

    //        //  Folder subFoldercol = site.GetFolderByServerRelativeUrl("Documents/Adams/Gresham Statements");

    //    }

    //    //GridView1.DataSource = dtSharepint;
    //    //GridView1.DataBind();
    //    dtSharepint.DefaultView.Sort = "path ASC";
    //    //ListBox1.DataSource = dtSharepint;
    //    //ListBox1.DataTextField = "path";
    //    //ListBox1.DataValueField = "path";
    //    //ListBox1.DataBind();
    //    //foreach (DataRow row in dtSharepint.Rows)
    //    //{
    //    //    ListBox1.Items.Add("row[Rootfolder]");
    //    //}


    //}
    //public List<string> getFolderPath(Web site, ClientContext context, string foldername, DataTable dtSharepint)
    //{

    //    List<string> folderlist = new List<string>();

    //    List<string> subfolderlist = new List<string>();


    //    // foldername = "Documents/Adams/" + f.Name;
    //    Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);
    //    //    ListCollection collList = site.Lists;

    //    FolderCollection fcolection1 = subFolder1.Folders;
    //    context.Load(fcolection1);
    //    //context.Load(collList);
    //    context.ExecuteQuery();

    //    foreach (Folder subfolder in fcolection1)
    //    {

    //        folderlist.Add(subfolder.Name.ToString());
    //        DataRow rowRootFolder = dtSharepint.NewRow();
    //        rowRootFolder["Subfolder"] = subfolder.Name.ToString();
    //        string newfoldername = foldername.Replace("Documents/Adams/", "");
    //        rowRootFolder["Path"] = "/" + newfoldername + "/" + subfolder.Name.ToString();
    //        ListBox1.Items.Add(newfoldername + "/" + subfolder.Name.ToString());
    //        dtSharepint.Rows.Add(rowRootFolder);

    //        subfolderlist = getSubFolderPath(site, context, foldername + "/" + subfolder.Name, dtSharepint);

    //    }

    //    return folderlist;
    //}
    //public List<string> getSubFolderPath(Web site, ClientContext context, string foldername, DataTable dtSharepint)
    //{

    //    List<string> folderlist = new List<string>();



    //    Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);

    //    FolderCollection fcolection1 = subFolder1.Folders;
    //    context.Load(fcolection1);
    //    //context.Load(collList);
    //    context.ExecuteQuery();

    //    foreach (Folder subfolder in fcolection1)
    //    {
    //        folderlist.Add(subfolder.Name.ToString());

    //        if (dtSharepint.Columns.Count != 4)
    //            dtSharepint.Columns.Add("subSubfolder");

    //        DataRow rowRootFolder = dtSharepint.NewRow();
    //        rowRootFolder["subSubfolder"] = subfolder.Name.ToString();
    //        string newfoldername = foldername.Replace("Documents/Adams/", "");
    //        rowRootFolder["Path"] = "/" + newfoldername + "/" + subfolder.Name.ToString();


    //        dtSharepint.Rows.Add(rowRootFolder);



    //    }

    //    return folderlist;
    //}

    #endregion


    public void newFolderstructuretempAnzi()//fetch all files from Anziano->Subfolder: ddlFolderName.SelectedItem.Text;
    {
        ViewState["AnzianoTable"] = null;
      //  string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
        context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);
        Web site = context.Web;
        text = ddlFolderName.SelectedItem.Text;
        Folder subFoldercol = site.GetFolderByServerRelativeUrl("Anziano" + "/" + text);

        //Microsoft.SharePoint.Client.File subfile = site.GetFileByServerRelativeUrl("Anziano" + "/" + text);
        ListCollection collList = site.Lists;

        // FolderCollection fcolection = subFoldercol.Folders;
        Microsoft.SharePoint.Client.FileCollection fcolection = subFoldercol.Files;
        context.Load(fcolection);
        context.Load(collList);
        context.ExecuteQuery();
        int fcount = fcolection.Count();

        if (fcount != 0)
        {
            // DataTable dtSharepint = new DataTable();


            dtSharepint.Columns.Add("FileGuid");
            dtSharepint.Columns.Add("FileName");
            dtSharepint.Columns.Add("PathUrl");
            dtSharepint.Columns.Add("AnzianoID");
            dtSharepint.Columns.Add("FileNameMatchBool", typeof(Boolean));
            dtSharepint.Columns.Add("Verified", typeof(Boolean));
            dtSharepint.Columns.Add("LegalEntity");
            dtSharepint.Columns.Add("HouseHold");
            dtSharepint.Columns.Add("HouseholdExists", typeof(Boolean));
            dtSharepint.Columns.Add("FileCopy", typeof(Boolean));




            //dt.Columns.Add("FileGuid", typeof(String));
            //dt.Columns.Add("FileName", typeof(String));
            //dt.Columns.Add("Url", typeof(String));
            //dt.Columns.Add("AnzianoID", typeof(String));
            //dt.Columns.Add("FileNameMatchBool", typeof(Boolean));
            //dt.Columns.Add("Verified", typeof(Boolean));
            //dt.Columns.Add("LegalEntity", typeof(String));
            //dt.Columns.Add("Household", typeof(String));
            //dt.Columns.Add("HouseholdExists", typeof(Boolean));
            //dt.Columns.Add("FileCopy", typeof(Boolean));



            //dtSharepint1.Columns.Add("FileName");
            //dtSharepint1.Columns.Add("AnzianoID");
            //dtSharepint1.Columns.Add("LegalEntity");
            //dtSharepint1.Columns.Add("HouseHold");

            //dtSharepint1.Columns.Add("PathUrl");

            ////dtSharepint.Columns.Add("HouseHold1");

            ///* For 2nd Gridview*/
            //dtSharepint2.Columns.Add("FileName");
            //dtSharepint2.Columns.Add("HouseHold");
            //dtSharepint2.Columns.Add("PathUrl");
            ///* For 2nd Gridview*/


            foreach (Microsoft.SharePoint.Client.File f in fcolection)
            {
                //Response.Write("<br/>");
                //Response.Write("<br/>," + f.Name.ToString());

                //FileCopy(f);
                //rowRootFolder = dtSharepint.NewRow();
                //rowRootFolder["FileName"] = f.Name.ToString();

                //rowRootFolder1 = dtSharePoint.NewRow();
                //rowRootFolder1["FileName"] = f.Name.ToString();


                //dtSharepint.Rows.Add(rowRootFolder);
                //  Label2.Text = "Anziano" + "/" + text + "/" + f.Name.ToString();
                FileNAME = f.Name.ToString();
                string path = siteUrl + "/" + "Anziano" + "/" + text + "/" + f.Name.ToString();

                string AnziId = getAnzianoIdFromFilename(FileNAME);
                //dtSharepint.Rows.Add(rowRootFolder);
                //dtSharepint.NewRow();
                if (int.TryParse(AnziId, out anz))
                {
                    dtSharepint.Rows.Add(f.UniqueId.ToString(), FileNAME, path, AnziId, false);
                    anzianoIds.Add(AnziId);
                }
                else
                    dtSharepint.Rows.Add(f.UniqueId.ToString(), FileNAME, path, "", false);
                //rowRootFolder["AnzianoID"] = AnziId;

                //dtSharepint.Rows.Add(rowRootFolder);

                //AnziID = AnziID + ",";

                //TextBox1.Text = path;

                //rowRootFolder["PathUrl"] = path;
                //rowRootFolder1["PathUrl"] = path;
                // AnziID = AnziId;
                AnziID = AnziID + "," + AnziId + "";

                //dtSharepint.Rows.Add(rowRootFolder);
                // dtSharePoint.Rows.Add(rowRootFolder1);



            }

            DataTable dtCrm = null;
            //dtCLIENT = newFolderstructure_client1();

            if (anzianoIds.Count > 0)
            {
                ds = clsDB.getDataSet("SP_S_SHAREPOINT_ANZIANO_DETAILS @AnzianoListTxt='" + AnziID + "'");
                dtCrm = ds.Tables[0];
            }

            Boolean householdExists = false;
            bool match = false;
            string legalEntity;
            string household = string.Empty;
            string[] words;
            DataRow[] drs;
            DataRow[] drs1;
            bool anzianoIDMatch = false;
            foreach (DataRow dr in dtSharepint.Rows)
            {
                DataTable dtClientNew = (DataTable)ViewState["dtSiteClientList"];
                anzianoIDMatch = false;
                dr["HouseholdExists"] = false;
                dr["FileNameMatchBool"] = false;
                dr["Verified"] = false;

                if (dr["AnzianoID"].ToString() != "")
                {
                    drs = dtCrm.Select("AnzianoID = '" + dr["AnzianoID"].ToString() + "'");
                    if (drs.Length > 0)     //if anziano id is there then match on anziano id
                    {
                        anzianoIDMatch = true;
                        legalEntity = drs[0]["LegalEntityName"].ToString();
                        household = drs[0]["HouseholdName"].ToString();
                        //household = household.ToLower().Replace("family", "").Trim();

                        dr["LegalEntity"] = legalEntity;
                        //drs1 = dtclients.Select("ClientName='" + dr["HouseHold"].ToString() + "'");
                        //if (drs1.Length > 0)
                        //{
                        //    string HouseHold0 = drs[0]["ClientName"].ToString();
                        //    if (household ==HouseHold0)
                        //    {
                        //    householdExists = true;
                        //      dr["Household"] = household;
                        //    }
                        //    else
                        //    {
                        //        householdExists = false;
                        //        dr["Household"] = "";
                        //    }
                        //}
                        //foreach (DataRow dr1 in dtClientNew.Rows)
                        //{
                        //    drs1 = dtClientNew.Select("household = '" + dr1["ClientName"].ToString() + "'");

                        //    if (drs1.Length > 0)
                        //    {
                        //        string Houshold0 = drs1[0]["ClientName"].ToString();

                        //        if (Houshold0 == household)
                        //        {
                        //            householdExists = true;
                        //            dr["Household"] = household;
                        //        }
                        //        else
                        //        {
                        //            householdExists = false;
                        //            dr["Household"] = "";
                        //        }
                        //    }
                        //}

                        foreach (DataRow dr1 in dtClientNew.Rows)
                        {

                            if (household == dr1["ClientName"].ToString())
                            {
                                householdExists = true;
                                dr["Household"] = household;
                                break;
                            }
                            else
                            {
                                householdExists = false;
                                dr["Household"] = "";

                            }
                        }

                        dr["HouseholdExists"] = householdExists;

                        match = true;
                        FileNAME = dr["FileName"].ToString();
                        words = legalEntity.Split(", .".ToCharArray());
                        foreach (string word in words)
                        {
                            if (word.Trim().Length <= 1)
                                continue;
                            if (FileNAME.IndexOf(word.Trim(), 0, StringComparison.InvariantCultureIgnoreCase) == -1)
                            {
                                match = false;
                                break;
                            }
                        }
                        dr["FileNameMatchBool"] = householdExists;
                        dr["Verified"] = householdExists;
                    }   //if (drs.Length > 0)
                }   //if (dr["AnzianoID"].ToString() != "")
                //if (match)
                //{
                //    householdExists = false;
                //    foreach (DataRow dr1 in dtClientNew.Rows)
                //    {

                //        if (household == dr1["ClientName"].ToString())
                //        {
                //            householdExists = true;
                //            dr["Household"] = household;
                //            break;
                //        }
                //        else
                //        {
                //            householdExists = false;
                //            dr["Household"] = "";
                //        }
                //    }

                //    dr["LegalEntity"] = "";
                //    dr["Household"] = household;
                //    dr["HouseholdExists"] = householdExists;
                //    dr["FileNameMatchBool"] = householdExists;
                //    dr["Verified"] = householdExists;
                //    break;
                //}
                if (!anzianoIDMatch)
                {
                    //if no anziano id is there 
                    dr["AnzianoID"] = ""; //clear out the Anziano ID

                    //if (match)
                    //foreach (DataRow drH in _householdTable.Rows)
                    if (!match)
                    {
                        dr["LegalEntity"] = "";
                        dr["Household"] = "";
                        dr["HouseholdExists"] = false;
                        dr["FileNameMatchBool"] = false;
                        dr["Verified"] = false;
                    }
                }
            }

            //dt3.Columns.Add("LegalEntity");
            //dt3.Columns.Add("HouseHold");
            //for (int i = 1; i < ds.Tables[0].Columns.Count; i++)
            //{
            //    string legalEnt = ds.Tables[0].Rows[i].ToString();

            //    for (int j = 0; j ==i; j++)
            //    {
            //        dtSharepint.Rows[j]["LegalEntity"] = legalEnt;

            //    }
            //}
            //foreach (DataRow row in ds.Tables[0].Rows)
            //{
            //    //DataRow rowRootFolder1 = dt3.NewRow();
            //    string legalEnt = row["LegalEntityName"].ToString();
            //    // dtSharepint.Rows.Add(rowRootFolder);
            //    string HouseHold1 = row["HouseHoldName"].ToString();

            //    string ID = row["AnzianoID"].ToString();
            //    string ID1 = string.Empty;
            //    // dtSharepint.Rows.Add();
            //    for (int j = 0; j < dtSharepint.Rows.Count; j++)
            //    {

            //        ID1 = dtSharepint.Rows[j]["AnzianoID"].ToString();


            //        if (ID == ID1)
            //        {
            //            if (dtSharepint.Rows[j]["LegalEntity"].ToString() == "" && dtSharepint.Rows[j]["HouseHold"].ToString() == "")
            //            {
            //                dtSharepint.Rows[j]["LegalEntity"] = legalEnt;
            //                dtSharepint.Rows[j]["HouseHold"] = HouseHold1;

            //                //rowRootFolder1 = dtSharepint1.NewRow();
            //                //rowRootFolder1["FileName"] = dtSharepint.Rows[j]["FileName"].ToString();

            //                //rowRootFolder1["AnzianoID"] = dtSharepint.Rows[j]["AnzianoID"].ToString();

            //                //rowRootFolder1["LegalEntity"] = dtSharepint.Rows[j]["LegalEntity"].ToString();

            //                //rowRootFolder1["HouseHold"] = dtSharepint.Rows[j]["HouseHold"].ToString();

            //                //dtSharepint1.Rows.Add(rowRootFolder1);
            //                //rowRootFolder1["FileName"] = dtSharepint.Rows[j].ToString();

            //                break;
            //            }

            //        }


            //        //else if (ID1 == "" && dtSharepint.Rows[j]["LegalEntity"].ToString() == "" && dtSharepint.Rows[j]["HouseHold"].ToString() == "")
            //        //{
            //        //    rowRootFolder2 = dtSharepint2.NewRow();
            //        //    rowRootFolder2["FileName"] = dtSharepint.Rows[j]["FileName"].ToString();
            //        //    dtSharepint2.Rows.Add(rowRootFolder2);
            //        //    break;


            //        //}

            //    }



            //    // rowRootFolder["HouseHold"] = row["HouseHoldName"].ToString();
            //    //dt3.Rows.Add(rowRootFolder);

            //}
            //foreach (DataRow r1 in dtSharepint.Rows)
            //{
            //    if ( r1["LegalEntity"].ToString() == "" && r1["HouseHold"].ToString() == "")
            //    {
            //        rowRootFolder2 = dtSharepint2.NewRow();
            //        rowRootFolder2["FileName"] = r1["FileName"].ToString();
            //        dtSharepint2.Rows.Add(rowRootFolder2);

            //    }
            //    else
            //    {
            //        rowRootFolder1 = dtSharepint1.NewRow();
            //        rowRootFolder1["FileName"] = r1["FileName"].ToString();

            //        rowRootFolder1["AnzianoID"] = r1["AnzianoID"].ToString();

            //        rowRootFolder1["LegalEntity"] = r1["LegalEntity"].ToString();

            //        rowRootFolder1["HouseHold"] = r1["HouseHold"].ToString();

            //        dtSharepint1.Rows.Add(rowRootFolder1);

            //    }
            //}
            // dtSharepint.Merge(dt3);

            //dt4.Merge(dtSharepint);
            //dt4.Merge(dt3);
            // dtSharepint.Rows.Add(rowRootFolder);
            //gvList.DataSource = dtSharepint;
            //gvList.DataBind();
            //gvList.Columns[0].Visible = true;
            //gvList.Columns[1].Visible = true;
            //gvList.Columns[2].Visible = true;
            //gvList.Columns[3].Visible = true;
            //gvList.Columns[4].Visible = true;
            //gvList.Columns[5].Visible = true;
            //gvList.Columns[6].Visible = true;
            //gvList.Columns[7].Visible = true;

            this.gvList.DataSource = new DataView(dtSharepint, "AnzianoID <> ''", "FileName", DataViewRowState.CurrentRows);
            this.gvList.DataBind();
            this.gvList1.DataSource = new DataView(dtSharepint, "AnzianoID = ''", "FileName", DataViewRowState.CurrentRows);
            this.gvList1.DataBind();
            ViewState["AnzianoTable"] = dtSharepint;
            //gvList.DataSource = dtSharepint1;
            //gvList.DataBind();

            //gvList1.DataSource = dtSharepint2;
            //gvList1.DataBind();



            //foreach (DataRow row1 in dtSharepint.Rows)
            //{

            //    string filename = row1["FileName"].ToString();

            //    foreach (GridViewRow grow in gvList.Rows)
            //    {

            //        DropDownList ddList1 = (DropDownList)grow.FindControl("ddlClients");


            //        CheckBox chkVerify = (CheckBox)grow.FindControl("chkVerify");

            //        CheckBox chleft = (CheckBox)grow.FindControl("chkbSelectBatch");

            //        //HyperLink hplink = ((HyperLink)grow.Cells[0].Controls[0]);

            //        //string FileName = hplink.ToString();
            //        string HouseHoldgrid = grow.Cells[5].Text.ToString().Replace("O&#39;Donnell", "O'Donnell");

            //        //string LegalEntity = grow.Cells[4].Text;

            //        string LegalEntity = grow.Cells[4].Text.Replace("Revocable", "Rev").Replace("Trust", "TR").Replace("Irrevocable", "IRRev");



            //        if (LegalEntity.ToUpper() != "&NBSP;" || HouseHoldgrid.ToUpper() != "&NBSP;")
            //        {


            //            foreach (DataRow row in dt2.Rows)
            //            {
            //                string HouseHold = row["ClientName"].ToString();
            //                if (HouseHoldgrid == HouseHold)
            //                {
            //                    ddList1.Text = HouseHoldgrid;
            //                    chkVerify.Checked = true;
            //                    break;
            //                }
            //                //else
            //                //{
            //                //    //chleft.Enabled = false;
            //                //    //break;
            //                //}

            //            }
            //        }

            //        else
            //        {
            //            chleft.Enabled = false;
            //            chkVerify.Enabled = false;
            //            //break;
            //        }
            //        //string LegalEntity1 = LegalEntity.Substring(0, 5);

            //        if (filename.Replace(",", "").Replace("_", "").Replace(".", "").Replace("'", "").ToUpper().Contains(LegalEntity.Replace(",", "").Replace("_", "").Replace(".", "").Replace("'", " ").Replace("Revocable", "Rev").Replace("Trust", "TR").Replace("Irrevocable", "IRRev").ToUpper(CultureInfo.InvariantCulture)))
            //        {
            //            if (ddList1.SelectedValue != "0")
            //            {
            //                chleft.Checked = true;
            //            }
            //            else
            //            {
            //                chleft.Enabled = false;
            //            }

            //        }
            //        else if (ddList1.SelectedValue == "0")
            //        {
            //            chleft.Enabled = false;
            //        }



            //    }

            //    //foreach(GridView grow1 in dtSharepint.Rows)
            //    //{
            //    //    DropDownList ddList1 = (DropDownList)grow1.FindControl("ddlClients");


            //    //    CheckBox chkVerify = (CheckBox)grow1.FindControl("chkVerify");

            //    //    CheckBox chleft = (CheckBox)grow1.FindControl("chkbSelectBatch");

            //    //    if (row1["AnzianoID"].ToString() == "" && row1["LegalEntity"].ToString() == "" && row1["HouseHold"].ToString() == "")
            //    //    {

            //    //    }
            //    //}
            //}
            //gvList.Columns[0].Visible = false;
            //gvList.Columns[1].Visible = false;
            //gvList.Columns[2].Visible = false;
            //gvList.Columns[3].Visible = false;
            //gvList.Columns[4].Visible = false;
            //gvList.Columns[5].Visible = false;
            //gvList.Columns[6].Visible = false;

            //ListBox1.DataSource = dtSharepint;
            //ListBox1.DataTextField = "path";
            //ListBox1.DataValueField = "path";
            //ListBox1.DataBind();
            //foreach (DataRow row in dtSharepint.Rows)
            //{
            //    ListBox1.Items.Add("row[Rootfolder]");
            //}
        }
        else
        {
            //gvList.Columns[0].Visible = false;
            //gvList.Columns[1].Visible = false;
            //gvList.Columns[2].Visible = false;
            //gvList.Columns[3].Visible = false;
            //gvList.Columns[4].Visible = false;
            //gvList.Columns[5].Visible = false;
            //gvList.Columns[6].Visible = false;
            //gvList.Columns[7].Visible = false;

            ClearGrid();
            lblMsg.Text = "No Files found in this folder";
            //lblMsg.ForeColor 
        }

    }

    public void ClearGrid()
    {
        gvList.DataSource = null;
        gvList.DataBind();
        gvList1.DataSource = null;
        gvList1.DataBind();

    }


    #region Not in Use
    //public List<string> getFolderPathAnzi(Web site, ClientContext context, string foldername, DataTable dtSharepint)
    //{

    //    List<string> folderlist = new List<string>();

    //    List<string> subfolderlist = new List<string>();


    //    // foldername = "Documents/Adams/" + f.Name;
    //    Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);
    //    //    ListCollection collList = site.Lists;

    //    FolderCollection fcolection1 = subFolder1.Folders;
    //    context.Load(fcolection1);
    //    //context.Load(collList);
    //    context.ExecuteQuery();

    //    foreach (Folder subfolder in fcolection1)
    //    {

    //        folderlist.Add(subfolder.Name.ToString());
    //        DataRow rowRootFolder = dtSharepint.NewRow();
    //        rowRootFolder["Subfolder"] = subfolder.Name.ToString();
    //        string newfoldername = foldername.Replace("Documents/Adams/", "");
    //        rowRootFolder["Path"] = newfoldername + "/" + subfolder.Name.ToString();
    //        ListBox1.Items.Add(newfoldername + "/" + subfolder.Name.ToString());
    //        dtSharepint.Rows.Add(rowRootFolder);

    //        subfolderlist = getSubFolderPath(site, context, foldername + "/" + subfolder.Name, dtSharepint);

    //    }

    //    return folderlist;
    //}
    //public List<string> getSubFolderPathAnzi(Web site, ClientContext context, string foldername, DataTable dtSharepint)
    //{

    //    List<string> folderlist = new List<string>();



    //    Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);

    //    FolderCollection fcolection1 = subFolder1.Folders;
    //    context.Load(fcolection1);
    //    //context.Load(collList);
    //    context.ExecuteQuery();

    //    foreach (Folder subfolder in fcolection1)
    //    {
    //        folderlist.Add(subfolder.Name.ToString());

    //        if (dtSharepint.Columns.Count != 4)
    //            dtSharepint.Columns.Add("subSubfolder");

    //        DataRow rowRootFolder = dtSharepint.NewRow();
    //        rowRootFolder["subSubfolder"] = subfolder.Name.ToString();
    //        string newfoldername = foldername.Replace("Documents/Adams/", "");
    //        rowRootFolder["Path"] = newfoldername + "/" + subfolder.Name.ToString();


    //        dtSharepint.Rows.Add(rowRootFolder);



    //    }

    //    return folderlist;
    //}
    #endregion


    public void FileCopy(Microsoft.SharePoint.Client.File files1)
    {
        // -- Get fIle and copy to Destination
        Stream filestrem = getFile(files1);
        string fileName = System.IO.Path.GetFileName(files1.Name);
        //string filepath = @"D:\PRACTICE\gp\";

        string filepath = Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\" + fileName;
        if (System.IO.File.Exists(filepath))
        {
            System.IO.File.Delete(filepath);
        }
        //string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
        //string filepath = siteUrl + "/" + "Documents taxonomy" + "/" + fileName;

        //FileStream fileStream = System.IO.File.Create(filepath + fileName, (int)filestrem.Length); // Test Local PAth

        FileStream fileStream = System.IO.File.Create(filepath, (int)filestrem.Length);
        //  FileStream fileStream = System.IO.File.Create(DestinationPath, (int)filestrem.Length); // Original PAth
        // Initialize the bytes array with the stream length and then fill it with data 
        byte[] bytesInStream = new byte[filestrem.Length];
        filestrem.Read(bytesInStream, 0, bytesInStream.Length);
        // Use write method to write to the file specified above 
        fileStream.Write(bytesInStream, 0, bytesInStream.Length);

        fileStream.Close();
    }


    //public void FileCopy1(string vSourcrFile)
    //{
    //    string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
    //    context = new ClientContext(siteUrl);
    //    SecureString passWord = new SecureString();
    //    foreach (var c in "w!ldWind36") passWord.AppendChar(c);
    //    context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
    //    Web site = context.Web;
    //    site = context.Web;


    //    byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
    //    System.IO.Stream stream = new System.IO.MemoryStream(bytes);

    //   // Folder currentRunFolder = site.GetFolderByServerRelativeUrl(vNewSharePointReportFolder);
    //    FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(Filenames), Overwrite = true };
    //    //currentRunFolder.Files.Add(newFile);


    //    List docs = context.Web.Lists.GetByTitle("Documents taxonomy");

    //    Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile);

    //    context.Load(uploadFile);
    //    context.Load(docs);
    //    context.ExecuteQuery();
    //}

    public Stream getFile(Microsoft.SharePoint.Client.File files1)
    {
        //string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
        //ClientContext context = new ClientContext(siteUrl);
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
            return new MemoryStream(ms.ToArray());
        }
    }

    public void CopyFilenew(string FolderPath, string destFilename, string vSourcrFile)  // string vSourcefile, string vDestinationFile
    {

      //  string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);
        Web site = context.Web;

        byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
        System.IO.Stream stream = new System.IO.MemoryStream(bytes);
        string filename = Path.GetFileName(destFilename);
        Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
        FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(destFilename), Overwrite = true };
        currentRunFolder.Files.Add(newFile);

        currentRunFolder.Update();
        //Response.Write("<br/>," + vSourcrFile);
        //Response.Write("<br/>," + destFilename);
        //Response.Write("<br/>," + FolderPath);
        context.ExecuteQuery();

        //  bool result= CheckFolderPathExis(vNewSharePointReportFolder);

    }

    protected void DeleteFile(string filename, Uri hostWeb)
    {

        //Web web = Context.Web;
        //List docs = web.Lists.GetByTitle("Documents");

        //Microsoft.SharePoint.Client.File f = web.GetFileByServerRelativeUrl("/Shared Documents/" + filename);
        //clientContext.Load(f);
        //f.DeleteObject();
        //clientContext.ExecuteQuery(); // Delete file here but throw Exception                
        //Console.Write("deleted");

    }

    public void DownloadFileUsingFileStream(Microsoft.SharePoint.Client.File file, string path)
    {
        string filename;
        filename = file.Name;
        DateTime timeModified = (DateTime)file.TimeLastModified;// Item["Modified"];

        //Copy the file
        byte[] contents = new byte[file.Length];
        // System.IO.File.WriteAllBytes(path + "\\" + filename, contents);

        FileStream outStream = new FileStream(path + "\\" + filename, FileMode.Create);
        outStream.Write(contents, 0, contents.Count());
        outStream.Close();


        ////byte[] filebytes = Encoding.Unicode.GetBytes(file.OpenBinaryStream().ToString());
        ////System.IO.File.Delete(path + "\\" + filename);
        ////System.IO.File.WriteAllBytes(path + "\\" + filename, filebytes);


        //byte[] filecontent =Encoding.Unicode.GetBytes( file.OpenBinaryStream().ToString());
        //FileStream fs = new FileStream(file.Name, FileMode.CreateNew);
        //using (fs)
        //{
        //    BinaryWriter bw = new BinaryWriter(fs);
        //    bw.Write(filecontent, 0, filecontent.Length);
        //    bw.Close();
        //}
    }

    protected void ddlFolderName_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMsg.Text = "";
        lblError.Text = "";
        if (ddlFolderName.SelectedValue != "0")
        {
            FolderPath = sh.getSPList();
            //year = dstaxonomy.Tables[2];
            Button1.Visible = true;
            lbltag.Visible = true;
            lblTest.Visible = true;
            lblMsg.Text = "";
            ChkTest.Visible = true;
            ddlPortalPath.Visible = true;
            BindFolderPathddl();
            ddlYear.Visible = true;
            BindYearddl();

            /* fucntion to display a DatawithGrid*/
            newFolderstructuretempAnzi();

        }
        else
        {
            gvList.DataSource = null;
            gvList.DataBind();
            lblMsg.Text = "Please select an Anziano folder";
        }
    }

    private string getAnzianoIdFromFilename(string fileName)
    {
        string[] fileParts = fileName.Split("_".ToCharArray());
        int id;
        foreach (string part in fileParts)
        {
            //if (part.Length == 7 && int.TryParse(part, out id))
            if (part.Length <= 7 && int.TryParse(part, out id))
                return part;
        }
        return "";


    }

    public DataTable newFolderstructure_client()
    {
       // string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;
        List list = context.Web.Lists.GetByTitle("Client");

        context.Load(list);
        context.Load(list.RootFolder);
        context.Load(list.RootFolder.Folders);
        context.Load(list.RootFolder.Files);
        context.ExecuteQuery();
        FolderCollection fcol = list.RootFolder.Folders;
        //  List<string> lstFile = new List<string>()
        List<string> lstFile = new List<string>();

        dtclients.Columns.Add("Value", typeof(String));
        dtclients.Columns.Add("Text", typeof(String));

        DataRow rowHouseHold = dtclients.NewRow();
        rowHouseHold["Value"] = "";
        rowHouseHold["Text"] = "Select Household";
        dtclients.Rows.Add(rowHouseHold);

        // DataTable dt1 = lstFile.tablecopy.ToDataTable();
        foreach (Folder f in fcol)
        {
            //Response.Write("<br/>," + f.Name.ToString());
            //dtclients.Rows.Add(f.Name.ToString());

            rowHouseHold = dtclients.NewRow();
            rowHouseHold["Value"] = f.Name.ToString();
            rowHouseHold["Text"] = f.Name.ToString();
            dtclients.Rows.Add(rowHouseHold);

        }
        dtclients.DefaultView.Sort = "Text ASC";
        return dtclients;
    }


    public DataTable newFolderstructure_client1()
    {
        try
        {
            //string siteUrl = "https://greshampartners.sharepoint.com";

            //string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);
            Web site = context.Web;

            // DataTable dt = new DataTable();
            dtclients.Columns.Add("Value", typeof(String));
            dtclients.Columns.Add("Text", typeof(String));
            DataRow dr = dtclients.NewRow();
            dr["Value"] = "";
            dr["Text"] = "Select Household";
            dtclients.Rows.Add(dr);

            //List list = site.Lists.GetByTitle("Mapping");

            List list = site.Lists.GetByTitle("Client");
            CamlQuery caml = new CamlQuery();
            Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(caml);
            context.Load(list);
            context.Load(items);
            context.ExecuteQuery();
            foreach (Microsoft.SharePoint.Client.ListItem item in items)
            {
                context.Load(item);
                context.ExecuteQuery();
                string ClientName = string.Empty;
                dr = dtclients.NewRow();
                ClientName = item["Title"].ToString();
                dr["Value"] = ClientName;
                dr["Text"] = ClientName;

                dtclients.Rows.Add(dr);
            }
        }
        catch (Exception ex)
        {

        }
        return dtclients;
    }

    protected void gvList_RowDataBound(object sender, GridViewRowEventArgs e)
    {


        count++;
        if (count == 1)
        {
            dt2 = newFolderstructure_client1();
        }
        for (int i = 0; i < e.Row.Cells.Count; i++)
        {
            e.Row.Cells[i].Attributes.Add("style", "white-space: nowrap;");

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                //gvList.Columns[5].ItemStyle.Width = 90;
                DropDownList ddList = (DropDownList)e.Row.FindControl("ddlClients");

                ////return DataTable havinf all client NAme

                //ddList.DataSource = dt2;
                //ddList.DataTextField = "HouseHold1";
                //ddList.DataValueField = "HouseHold1";
                //ddList.DataBind();
                //ddList.Items.Insert(0, "");
                //ddList.Items[0].Value = "0";

                ddList.DataSource = dt2;
                ddList.DataTextField = "ClientName";
                ddList.DataValueField = "ClientName";
                ddList.DataBind();
                //ddList.Text = "HouseHold";

                ddList.Items.Insert(0, "Select HouseHold");
                ddList.Items[0].Value = "0";

            }
        }
        //BindGridView();
    }

    protected void gvList1_RowDataBound(object sender, GridViewRowEventArgs e)
    {


        count++;
        if (count == 1)
        {
            dt2 = newFolderstructure_client1();
        }
        for (int i = 0; i < e.Row.Cells.Count; i++)
        {
            e.Row.Cells[i].Attributes.Add("style", "white-space: nowrap;");

            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                //gvList.Columns[5].ItemStyle.Width = 90;
                DropDownList ddList = (DropDownList)e.Row.FindControl("ddlClients1");

                ////return DataTable havinf all client NAme

                //ddList.DataSource = dt2;
                //ddList.DataTextField = "HouseHold1";
                //ddList.DataValueField = "HouseHold1";
                //ddList.DataBind();
                //ddList.Items.Insert(0, "");
                //ddList.Items[0].Value = "0";

                ddList.DataSource = dt2;
                ddList.DataTextField = "ClientName";
                ddList.DataValueField = "ClientName";
                ddList.DataBind();

                ddList.Items.Insert(0, "Select HouseHold");
                ddList.Items[0].Value = "0";

            }
        }
        //BindGridView();
    }



    protected void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
    {

    }



    public void BindFolderPathddl()
    {
        ddlPortalPath.DataTextField = "FolderPath";
        ddlPortalPath.DataSource = FolderPath;
        ddlPortalPath.DataBind();
        ddlPortalPath.Items.Insert(0, "Select");
        ddlPortalPath.Items[0].Value = "0";


    }


    public void BindYearddl()
    {

        ddlYear.DataTextField = "Year";
        ddlYear.DataSource = ViewState["Year"];
        ddlYear.DataBind();
        ddlYear.Items.Insert(0, "Select");
        ddlYear.Items[0].Value = "0";

    }


    protected void ddlyear_SelectedIndexChanged(object sender, EventArgs e)
    {

    }




    protected void Button1_Click(object sender, EventArgs e)
    {
        lblMsg.Text = "";
        lblError.Text = "";
        string folderPath = ddlPortalPath.SelectedItem.Text;
        string Vfolderpath = ddlPortalPath.SelectedValue.ToString();

        string vYear = ddlYear.SelectedValue.ToString();

        dtMail.Columns.Add("FileName");
        dtMail.Columns.Add("AnzianoId");
        dtMail.Columns.Add("LegalEntity");
        dtMail.Columns.Add("HouseHold");
        dtMail.Columns.Add("NameMatch");
        dtMail.Columns.Add("Verified");
        dtMail.Columns.Add("Destination");


        dtMail1.Columns.Add("FileName");
        dtMail1.Columns.Add("AnzianoId");
        dtMail1.Columns.Add("LegalEntity");
        dtMail1.Columns.Add("HouseHold");
        dtMail1.Columns.Add("NameMatch");
        dtMail1.Columns.Add("Verified");
        dtMail1.Columns.Add("Destination");

        if (Vfolderpath != "0")
        {
            string ddlyear = "";

            if (ddlYear.Visible)
            {
                ddlyear = ddlYear.SelectedValue;
            }
            if (ddlyear != "0")
            {

                DataTable dtFolderData = (DataTable)ViewState["dtFolderData"];
                DataTable dtClient = (DataTable)ViewState["dtSiteClientList"];

                string PathTaggingName = string.Empty;
                string PathTaggingID = string.Empty;
                string vIsYear = string.Empty;
                foreach (DataRow rw in dtFolderData.Rows)
                {
                    if (folderPath == rw["FolderPath"].ToString())
                    {
                        PathTaggingName = rw["Tag"].ToString();
                        vIsYear = rw["OnPortal"].ToString();
                        break;
                    }

                }

                DataSet dsTaxonomyclientPortal = (DataSet)ViewState["dsTaxonomyclientPortal"];
                DataTable dtDocumentType = dsTaxonomyclientPortal.Tables[1];
                DataTable dtClientSite = dsTaxonomyclientPortal.Tables[0];
                DataTable dtYear = dsTaxonomyclientPortal.Tables[2];


                foreach (DataRow rw in dtDocumentType.Rows)
                {
                    if (rw["DocumentType"].ToString() == PathTaggingName)
                    {
                        PathTaggingID = rw["iID"].ToString();
                        break;
                    }
                }

                string Taggingyear = ddlYear.SelectedItem.Text;
                string TaggingYearID = string.Empty;
                if (ddlYear.SelectedValue != "0")
                {
                    if (vIsYear == "True")
                    {
                        foreach (DataRow rw in dtYear.Rows)
                        {
                            if (rw["Year"].ToString() == Taggingyear)
                                TaggingYearID = rw["iID"].ToString();
                        }

                    }

                }

                foreach (GridViewRow grow in gvList.Rows)
                {


                    string cell5 = grow.Cells[5].Text;//household

                    string LegalEntity = grow.Cells[4].Text;//LegalEntityName

                    string AnzianoId = grow.Cells[3].Text;//AnzianoId

                    DropDownList ddList = (DropDownList)grow.FindControl("dtclients");
                    string cell6 = ddList.SelectedItem.Text;//client Name

                    CheckBox chkleft = (CheckBox)grow.FindControl("chkbSelectBatch");

                    CheckBox chkright = (CheckBox)grow.FindControl("chkVerify");


                    //if (chkleft.Checked == true)
                    //{

                    if (ddList.SelectedValue.ToString() != "0")
                    {
                        //string CName = grow.Cells[6].Text;
                        string TaggingClientName = cell6;
                        string taggingClientID = string.Empty;

                        //string iClientID = string.Empty;
                        foreach (DataRow rw in dtClientSite.Rows)
                        {
                            if (TaggingClientName == rw["ClientName"].ToString())
                            {
                                taggingClientID = rw["iID"].ToString();
                                break;
                            }
                        }
                        HyperLink hyper = grow.Cells[2].Controls[0] as HyperLink;


                        string fName = hyper.Text;

                       // string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

                        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
                        context = new ClientContext(siteUrl);

                        SecureString passWord = new SecureString();
                        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
                        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
                        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
                        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
                        foreach (var c in Pass) passWord.AppendChar(c);
                        context.Credentials = new SharePointOnlineCredentials(user, passWord);

                        Web site = context.Web;
                        text = ddlFolderName.SelectedItem.Text;
                        Folder subFoldercol = site.GetFolderByServerRelativeUrl("Anziano" + "/" + text);

                        //Microsoft.SharePoint.Client.File subfile = site.GetFileByServerRelativeUrl("Anziano" + "/" + text);
                        ListCollection collList = site.Lists;
                        int filecount = 0;
                        // FolderCollection fcolection = subFoldercol.Folders;
                        Microsoft.SharePoint.Client.FileCollection fcolection = subFoldercol.Files;
                        context.Load(fcolection);
                        context.Load(collList);
                        context.ExecuteQuery();

                        for (int i = 0; i < fcolection.Count; i++)
                        //foreach (Microsoft.SharePoint.Client.File f in fcolection)
                        {
                            Microsoft.SharePoint.Client.File f = fcolection[i];
                            // filecount = fcolection.Count();
                            string AnziFileName = f.Name.ToString();

                            // int a = fcolection[i];
                            //  string FolderPath = "Documents taxonomy";
                            string FolderPath = "Documents taxonomy";

                            Microsoft.SharePoint.Client.ListItem item = null;

                            Microsoft.SharePoint.Client.List docs = null;


                            if (fName == AnziFileName)
                            {

                                string path = string.Empty;
                                if (chkleft.Checked == true)
                                {
                                    FileCopy(f);
                                    //string path = @"D:\PRACTICE\gp\" + AnziFileName;

                                    path = Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\" + AnziFileName;
                                    string AnziFile = string.Empty;

                                    string exte = string.Empty;

                                    string AnziFileName1 = string.Empty;
                                    string exte1 = string.Empty;

                                    // string exte = System.IO.Path.GetExtension(path);

                                    //string AnziFile = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                    //string exte = System.IO.Path.GetExtension(path);

                                    //if (ChkTest.Checked == true)
                                    //{
                                    //#region copynewfile
                                    // CopyFilenew(FolderPath, AnziFile + "_Test" + exte, path);
                                    Folder currentRunFolder = null;
                                    FileCreationInformation newFile = null;


                                    #region Test mode File
                                    if (ChkTest.Checked == true)
                                    {
                                        AnziFile = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                        exte = System.IO.Path.GetExtension(path);
                                        //string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
                                        //ClientContext context = new ClientContext(siteUrl);
                                        //SecureString passWord = new SecureString();
                                        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
                                        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
                                        //Web site = context.Web;

                                        byte[] bytes = System.IO.File.ReadAllBytes(path);
                                        System.IO.Stream stream = new System.IO.MemoryStream(bytes);
                                        string filename = Path.GetFileName(AnziFile + "_Test" + exte);
                                        currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                                        newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFile + "_Test" + exte), Overwrite = true };
                                    }

                                    #endregion

                                    #region without test Mode
                                    else
                                    {

                                        AnziFileName1 = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                        exte1 = System.IO.Path.GetExtension(path);
                                        byte[] bytes = System.IO.File.ReadAllBytes(path);
                                        System.IO.Stream stream = new System.IO.MemoryStream(bytes);


                                        string filename = Path.GetFileName(AnziFileName1 + exte1);
                                        currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                                        newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFileName1 + exte1), Overwrite = true };
                                    }
                                    #endregion

                                    //  docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    Microsoft.SharePoint.Client.File uploadFile = currentRunFolder.Files.Add(newFile);

                                    //currentRunFolder.Update();
                                    ////Response.Write("<br/>," + vSourcrFile);
                                    ////Response.Write("<br/>," + destFilename);
                                    ////Response.Write("<br/>," + FolderPath);
                                    //context.ExecuteQuery();
                                    context.Load(uploadFile);
                                    context.Load(docs);
                                    context.ExecuteQuery();

                                    context.Load(uploadFile.ListItemAllFields);
                                    context.ExecuteQuery();

                                    // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

                                    item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
                                    context.Load(item);
                                    context.ExecuteQuery();

                                    //#endregion
                                    Success++;
                                    //Array.ForEach(Directory.GetFiles(Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\"), System.IO.File.Delete);

                                    DataRow rowMail = dtMail.NewRow();
                                    rowMail["FileName"] = fName.ToString();
                                    //dtMail.Rows.Add(rowMail);

                                    rowMail["AnzianoId"] = AnzianoId.ToString();
                                    rowMail["LegalEntity"] = LegalEntity.ToString();
                                    rowMail["HouseHold"] = cell5.ToString();
                                    if (chkleft.Checked)
                                    {
                                        rowMail["NameMatch"] = "Y";
                                    }
                                    if (chkleft.Checked)
                                    {
                                        if (chkright.Checked)
                                            rowMail["Verified"] = "Y";
                                        else
                                            rowMail["Verified"] = "N";

                                    }

                                    if (ChkTest.Checked == true)
                                    {
                                        //rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFile + "_Test" + exte;
                                        rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFile + "_Test" + exte;
                                    }
                                    else
                                    {
                                        rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFileName1 + exte1;

                                    }
                                    dtMail.Rows.Add(rowMail);

                                    #region Tagging 10/12/2016


                                    //                                    List docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    //                                    context.Load(docs);
                                    //                                    context.ExecuteQuery();
                                    //                                    CamlQuery _CamlQuery = new CamlQuery();
                                    //                                    _CamlQuery.ViewXml =
                                    //                                       @"<View Scope='RecursiveAll'>
                                    //                   <Query>
                                    //                   <OrderBy UseIndexForOrderBy = 'TRUE'> <FieldRef Name='Modified' Ascending='False' /></OrderBy>
                                    //                      </Query>
                                    //                      <RowLimit>10</RowLimit>
                                    //                      </View>";

                                    //                                    Microsoft.SharePoint.Client.ListItemCollection _ListItemCollection = docs.GetItems(_CamlQuery);
                                    //                                    context.Load(_ListItemCollection);
                                    //                                    context.ExecuteQuery();



                                    //                                    Microsoft.SharePoint.Client.ListItem item = null;
                                    //                                    foreach (Microsoft.SharePoint.Client.ListItem listItem in _ListItemCollection)
                                    //                                    {
                                    //                                        if (AnziFile + "_Test" + exte == listItem["FileLeafRef"].ToString())
                                    //                                        {
                                    //                                            item = listItem;
                                    //                                            break;
                                    //                                        }
                                    //                                        else if (AnziFileName == listItem["FileLeafRef"].ToString())
                                    //                                        {
                                    //                                            item = listItem;
                                    //                                            break;
                                    //                                        }
                                    //                                    }





                                    //                                    context.Load(_ListItemCollection);
                                    //                                    context.ExecuteQuery();






                                    //context.ExecuteQuery();

                                    TaxonomyFieldValue taxonomyFieldValueClient = new TaxonomyFieldValue();
                                    TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
                                    TaxonomyFieldValue taxonomyFieldValueYear = new TaxonomyFieldValue();

                                    taxonomyFieldValuePath.TermGuid = PathTaggingID;
                                    taxonomyFieldValuePath.Label = PathTaggingName;

                                    taxonomyFieldValueClient.TermGuid = taggingClientID;
                                    taxonomyFieldValueClient.Label = TaggingClientName;

                                    //string Taggingyear = ddlYear.SelectedItem.Text;
                                    //string TaggingYearID = string.Empty;
                                    if (ddlYear.SelectedValue != "0")
                                    {


                                        if (vIsYear == "True")
                                        {
                                            taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                            taxonomyFieldValueYear.Label = Taggingyear;
                                            if (taggingClientID != "")
                                                item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;
                                        }
                                        else
                                        {

                                            taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                            taxonomyFieldValueYear.Label = Taggingyear;
                                            if (taggingClientID != "")
                                                item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

                                        }
                                    }
                                    else
                                    {
                                        taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                        taxonomyFieldValueYear.Label = Taggingyear;
                                        if (taggingClientID != "")
                                            item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
                                    }


                                    //else
                                    //{

                                    //    taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                    //    taxonomyFieldValueYear.Label = Taggingyear;
                                    //    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

                                    //}
                                    if (taggingClientID != "")
                                        item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;

                                    if (ddList.SelectedValue.ToString() != "0" && taggingClientID != "")
                                        item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;

                                    if (taggingClientID != "")
                                    {
                                        item.Update();
                                        docs.Update();
                                        context.ExecuteQuery();
                                    }

                                    if (ChkTest.Checked == false)
                                    {
                                        f.DeleteObject();
                                        context.ExecuteQuery();
                                    }

                                    if (path != "")
                                        System.IO.File.Delete(path);
                                    //lblMsg.Text = "File Copied Successfully";
                                    break;

                                    #endregion
                                    //}

                                    //else
                                    //{
                                    //    // CopyFilenew(FolderPath, AnziFileName, path);
                                    //    #region copynewfile

                                    //    string AnziFileName1 = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                    //    string exte1 = System.IO.Path.GetExtension(path);
                                    //    byte[] bytes = System.IO.File.ReadAllBytes(path);
                                    //    System.IO.Stream stream = new System.IO.MemoryStream(bytes);


                                    //    string filename = Path.GetFileName(AnziFileName1 + exte1);
                                    //    Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                                    //    FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFileName1 + exte1), Overwrite = true };
                                    //    //  docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    //    docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    //    Microsoft.SharePoint.Client.File uploadFile = currentRunFolder.Files.Add(newFile);


                                    //    context.Load(uploadFile);
                                    //    context.Load(docs);
                                    //    context.ExecuteQuery();

                                    //    context.Load(uploadFile.ListItemAllFields);
                                    //    context.ExecuteQuery();

                                    //    // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

                                    //    item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
                                    //    context.Load(item);
                                    //    context.ExecuteQuery();

                                    //    #endregion
                                    //    Success++;

                                    //    //string[] words;


                                    //    DataRow rowMail = dtMail.NewRow();
                                    //    rowMail["FileName"] = fName.ToString();
                                    //    //dtMail.Rows.Add(rowHouseHold);

                                    //    rowMail["AnzianoId"] = AnzianoId.ToString();
                                    //    rowMail["LegalEntity"] = LegalEntity.ToString();
                                    //    rowMail["HouseHold"] = cell5.ToString();
                                    //    if (chkleft.Checked)
                                    //    {
                                    //        rowMail["NameMatch"] = "Y";
                                    //    }
                                    //    if (chkleft.Checked)
                                    //    {
                                    //        if (chkright.Checked)
                                    //            rowMail["Verified"] = "Y";
                                    //        else
                                    //            rowMail["Verified"] = "N";

                                    //    }
                                    //    rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFileName1 + exte1;
                                    //    dtMail.Rows.Add(rowMail);


                                    //    //dtMail.Columns.Add("Destination");

                                    //    f.DeleteObject();
                                    //    context.ExecuteQuery();
                                    //}


                                    // f.DeleteObject();
                                    // System.IO.File.Delete(path);
                                    //context.ExecuteQuery();





                                    //System.IO.File.Delete(siteUrl + "/" + "Anziano" + "/" + text + "/" + AnziFileName);
                                    //#region Tagging 10/12/2016


                                    ////                                    List docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    ////                                    context.Load(docs);
                                    ////                                    context.ExecuteQuery();
                                    ////                                    CamlQuery _CamlQuery = new CamlQuery();
                                    ////                                    _CamlQuery.ViewXml =
                                    ////                                       @"<View Scope='RecursiveAll'>
                                    ////                   <Query>
                                    ////                   <OrderBy UseIndexForOrderBy = 'TRUE'> <FieldRef Name='Modified' Ascending='False' /></OrderBy>
                                    ////                      </Query>
                                    ////                      <RowLimit>10</RowLimit>
                                    ////                      </View>";

                                    ////                                    Microsoft.SharePoint.Client.ListItemCollection _ListItemCollection = docs.GetItems(_CamlQuery);
                                    ////                                    context.Load(_ListItemCollection);
                                    ////                                    context.ExecuteQuery();



                                    ////                                    Microsoft.SharePoint.Client.ListItem item = null;
                                    ////                                    foreach (Microsoft.SharePoint.Client.ListItem listItem in _ListItemCollection)
                                    ////                                    {
                                    ////                                        if (AnziFile + "_Test" + exte == listItem["FileLeafRef"].ToString())
                                    ////                                        {
                                    ////                                            item = listItem;
                                    ////                                            break;
                                    ////                                        }
                                    ////                                        else if (AnziFileName == listItem["FileLeafRef"].ToString())
                                    ////                                        {
                                    ////                                            item = listItem;
                                    ////                                            break;
                                    ////                                        }
                                    ////                                    }





                                    ////                                    context.Load(_ListItemCollection);
                                    ////                                    context.ExecuteQuery();






                                    ////context.ExecuteQuery();

                                    //TaxonomyFieldValue taxonomyFieldValueClient = new TaxonomyFieldValue();
                                    //TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
                                    //TaxonomyFieldValue taxonomyFieldValueYear = new TaxonomyFieldValue();

                                    //taxonomyFieldValuePath.TermGuid = PathTaggingID;
                                    //taxonomyFieldValuePath.Label = PathTaggingName;

                                    //taxonomyFieldValueClient.TermGuid = taggingClientID;
                                    //taxonomyFieldValueClient.Label = TaggingClientName;

                                    ////string Taggingyear = ddlYear.SelectedItem.Text;
                                    ////string TaggingYearID = string.Empty;
                                    //if (ddlYear.SelectedValue != "0")
                                    //{


                                    //    if (vIsYear == "True")
                                    //    {
                                    //        taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                    //        taxonomyFieldValueYear.Label = Taggingyear;
                                    //        item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;
                                    //    }
                                    //    else
                                    //    {

                                    //        taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                    //        taxonomyFieldValueYear.Label = Taggingyear;
                                    //        item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                    //    taxonomyFieldValueYear.Label = Taggingyear;
                                    //    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
                                    //}


                                    ////else
                                    ////{

                                    ////    taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                    ////    taxonomyFieldValueYear.Label = Taggingyear;
                                    ////    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

                                    ////}

                                    //item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;
                                    //item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;
                                    //item.Update();
                                    //docs.Update();
                                    //context.ExecuteQuery();

                                    //f.DeleteObject();
                                    //context.ExecuteQuery();
                                    //System.IO.File.Delete(path);
                                    ////lblMsg.Text = "File Copied Successfully";
                                    //break;

                                    //#endregion


                                    //DownloadFileUsingFileStream(f, siteUrl+"/" +"Documents taxonomy" +"/");
                                }

                                else if (chkleft.Enabled == false && chkright.Checked == true)
                                {
                                    FileCopy(f);
                                    //string path = @"D:\PRACTICE\gp\" + AnziFileName;
                                    path = Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\" + AnziFileName;

                                    string AnziFile = string.Empty;

                                    string exte = string.Empty;

                                    string AnziFileName1 = string.Empty;
                                    string exte1 = string.Empty;

                                    // string exte = System.IO.Path.GetExtension(path);

                                    //string AnziFile = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                    //string exte = System.IO.Path.GetExtension(path);

                                    //if (ChkTest.Checked == true)
                                    //{
                                    //#region copynewfile
                                    // CopyFilenew(FolderPath, AnziFile + "_Test" + exte, path);
                                    Folder currentRunFolder = null;
                                    FileCreationInformation newFile = null;


                                    #region Test mode File
                                    if (ChkTest.Checked == true)
                                    {
                                        AnziFile = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                        exte = System.IO.Path.GetExtension(path);
                                        //string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
                                        //ClientContext context = new ClientContext(siteUrl);
                                        //SecureString passWord = new SecureString();
                                        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
                                        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
                                        //Web site = context.Web;

                                        byte[] bytes = System.IO.File.ReadAllBytes(path);
                                        System.IO.Stream stream = new System.IO.MemoryStream(bytes);
                                        string filename = Path.GetFileName(AnziFile + "_Test" + exte);
                                        currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                                        newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFile + "_Test" + exte), Overwrite = true };
                                    }

                                    #endregion

                                    #region without test Mode
                                    else
                                    {

                                        AnziFileName1 = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                        exte1 = System.IO.Path.GetExtension(path);
                                        byte[] bytes = System.IO.File.ReadAllBytes(path);
                                        System.IO.Stream stream = new System.IO.MemoryStream(bytes);


                                        string filename = Path.GetFileName(AnziFileName1 + exte1);
                                        currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                                        newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFileName1 + exte1), Overwrite = true };
                                    }
                                    #endregion

                                    //  docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    Microsoft.SharePoint.Client.File uploadFile = currentRunFolder.Files.Add(newFile);

                                    //currentRunFolder.Update();
                                    ////Response.Write("<br/>," + vSourcrFile);
                                    ////Response.Write("<br/>," + destFilename);
                                    ////Response.Write("<br/>," + FolderPath);
                                    //context.ExecuteQuery();
                                    context.Load(uploadFile);
                                    context.Load(docs);
                                    context.ExecuteQuery();

                                    context.Load(uploadFile.ListItemAllFields);
                                    context.ExecuteQuery();

                                    // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

                                    item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
                                    context.Load(item);
                                    context.ExecuteQuery();

                                    //#endregion
                                    Success++;
                                    //Array.ForEach(Directory.GetFiles(Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\"), System.IO.File.Delete);

                                    DataRow rowMail = dtMail.NewRow();
                                    rowMail["FileName"] = fName.ToString();
                                    //dtMail.Rows.Add(rowMail);

                                    rowMail["AnzianoId"] = AnzianoId.ToString();
                                    rowMail["LegalEntity"] = LegalEntity.ToString();
                                    rowMail["HouseHold"] = cell6.ToString().Replace("Select Household","");
                                    if (chkleft.Checked)
                                    {
                                        rowMail["NameMatch"] = "Y";
                                    }
                                    if (chkleft.Checked)
                                    {
                                        if (chkright.Checked)
                                            rowMail["Verified"] = "Y";
                                        else
                                            rowMail["Verified"] = "N";

                                    }
                                    else if (chkleft.Enabled == false && chkright.Checked == true)
                                    {
                                        if (chkright.Checked)
                                            rowMail["Verified"] = "Y";
                                        else
                                            rowMail["Verified"] = "N";

                                    }

                                    if (ChkTest.Checked == true)
                                    {
                                        //rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFile + "_Test" + exte;
                                        rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFile + "_Test" + exte;
                                    }
                                    else
                                    {
                                        rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFileName1 + exte1;

                                    }
                                    dtMail.Rows.Add(rowMail);

                                    #region Tagging 10/12/2016


                                    //                                    List docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    //                                    context.Load(docs);
                                    //                                    context.ExecuteQuery();
                                    //                                    CamlQuery _CamlQuery = new CamlQuery();
                                    //                                    _CamlQuery.ViewXml =
                                    //                                       @"<View Scope='RecursiveAll'>
                                    //                   <Query>
                                    //                   <OrderBy UseIndexForOrderBy = 'TRUE'> <FieldRef Name='Modified' Ascending='False' /></OrderBy>
                                    //                      </Query>
                                    //                      <RowLimit>10</RowLimit>
                                    //                      </View>";

                                    //                                    Microsoft.SharePoint.Client.ListItemCollection _ListItemCollection = docs.GetItems(_CamlQuery);
                                    //                                    context.Load(_ListItemCollection);
                                    //                                    context.ExecuteQuery();



                                    //                                    Microsoft.SharePoint.Client.ListItem item = null;
                                    //                                    foreach (Microsoft.SharePoint.Client.ListItem listItem in _ListItemCollection)
                                    //                                    {
                                    //                                        if (AnziFile + "_Test" + exte == listItem["FileLeafRef"].ToString())
                                    //                                        {
                                    //                                            item = listItem;
                                    //                                            break;
                                    //                                        }
                                    //                                        else if (AnziFileName == listItem["FileLeafRef"].ToString())
                                    //                                        {
                                    //                                            item = listItem;
                                    //                                            break;
                                    //                                        }
                                    //                                    }





                                    //                                    context.Load(_ListItemCollection);
                                    //                                    context.ExecuteQuery();






                                    //context.ExecuteQuery();

                                    TaxonomyFieldValue taxonomyFieldValueClient = new TaxonomyFieldValue();
                                    TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
                                    TaxonomyFieldValue taxonomyFieldValueYear = new TaxonomyFieldValue();

                                    taxonomyFieldValuePath.TermGuid = PathTaggingID;
                                    taxonomyFieldValuePath.Label = PathTaggingName;

                                    taxonomyFieldValueClient.TermGuid = taggingClientID;
                                    taxonomyFieldValueClient.Label = TaggingClientName;

                                    //string Taggingyear = ddlYear.SelectedItem.Text;
                                    //string TaggingYearID = string.Empty;
                                    if (ddlYear.SelectedValue != "0")
                                    {


                                        if (vIsYear == "True")
                                        {
                                            taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                            taxonomyFieldValueYear.Label = Taggingyear;
                                            if (taggingClientID != "")
                                                item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;
                                        }
                                        else
                                        {

                                            taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                            taxonomyFieldValueYear.Label = Taggingyear;
                                            if (taggingClientID != "")
                                                item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

                                        }
                                    }
                                    else
                                    {
                                        taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                        taxonomyFieldValueYear.Label = Taggingyear;
                                        if (taggingClientID != "")
                                            item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
                                    }


                                    //else
                                    //{

                                    //    taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                    //    taxonomyFieldValueYear.Label = Taggingyear;
                                    //    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

                                    //}
                                    if (taggingClientID != "")
                                        item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;

                                    if (ddList.SelectedValue.ToString() != "0" && taggingClientID != "")
                                        item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;

                                    if (taggingClientID != "")
                                    {
                                        item.Update();
                                        docs.Update();
                                        context.ExecuteQuery();
                                    }

                                    if (ChkTest.Checked == false)
                                    {
                                        f.DeleteObject();
                                        context.ExecuteQuery();
                                    }

                                    if (path != "")
                                        System.IO.File.Delete(path);
                                    //lblMsg.Text = "File Copied Successfully";
                                    break;

                                    #endregion

                                    //}

                                    //else
                                    //{
                                    //    // CopyFilenew(FolderPath, AnziFileName, path);
                                    //    #region copynewfile

                                    //    string AnziFileName1 = System.IO.Path.GetFileNameWithoutExtension(path).Replace(AnzianoId + "_", "");
                                    //    string exte1 = System.IO.Path.GetExtension(path);
                                    //    byte[] bytes = System.IO.File.ReadAllBytes(path);
                                    //    System.IO.Stream stream = new System.IO.MemoryStream(bytes);


                                    //    string filename = Path.GetFileName(AnziFileName1 + exte1);
                                    //    Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                                    //    FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFileName1 + exte1), Overwrite = true };
                                    //    //  docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    //    docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                                    //    Microsoft.SharePoint.Client.File uploadFile = currentRunFolder.Files.Add(newFile);


                                    //    context.Load(uploadFile);
                                    //    context.Load(docs);
                                    //    context.ExecuteQuery();

                                    //    context.Load(uploadFile.ListItemAllFields);
                                    //    context.ExecuteQuery();

                                    //    // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

                                    //    item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
                                    //    context.Load(item);
                                    //    context.ExecuteQuery();

                                    //    #endregion
                                    //    Success++;

                                    //    //string[] words;


                                    //    DataRow rowMail = dtMail.NewRow();
                                    //    rowMail["FileName"] = fName.ToString();
                                    //    //dtMail.Rows.Add(rowHouseHold);

                                    //    rowMail["AnzianoId"] = AnzianoId.ToString();
                                    //    rowMail["LegalEntity"] = LegalEntity.ToString();
                                    //    rowMail["HouseHold"] = cell5.ToString();
                                    //    if (chkleft.Checked)
                                    //    {
                                    //        rowMail["NameMatch"] = "Y";
                                    //    }
                                    //    if (chkleft.Checked)
                                    //    {
                                    //        if (chkright.Checked)
                                    //            rowMail["Verified"] = "Y";
                                    //        else
                                    //            rowMail["Verified"] = "N";

                                    //    }
                                    //    rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFileName1 + exte1;
                                    //    dtMail.Rows.Add(rowMail);


                                    //    //dtMail.Columns.Add("Destination");

                                    //    f.DeleteObject();
                                    //    context.ExecuteQuery();
                                    //}


                                    // f.DeleteObject();
                                    // System.IO.File.Delete(path);
                                    //context.ExecuteQuery();





                                    //System.IO.File.Delete(siteUrl + "/" + "Anziano" + "/" + text + "/" + AnziFileName);

                                }


                            }
                        }

                    }
                    else
                    {
                        //lblError.Text = "No Files Copied";
                    }

                    //}

                    //else
                    //{
                    //    //lblError.Text = "No Files Copied";
                    //}



                }


                //#region 2nd Grid

                //foreach (GridViewRow grow in gvList1.Rows)
                //{
                //    DropDownList ddList = (DropDownList)grow.FindControl("ddlClients1");

                //    CheckBox chkleft = (CheckBox)grow.FindControl("chkbSelectBatch");

                //    CheckBox chkright = (CheckBox)grow.FindControl("chkVerify");

                //    HyperLink hyper = grow.Cells[2].Controls[0] as HyperLink;

                //    if (chkleft.Checked)
                //    {

                //        // string CName = grow.Cells[6].Text;
                //        string TaggingClientName = ddList.SelectedItem.Text;
                //        string taggingClientID = string.Empty;

                //        string iClientID = string.Empty;
                //        foreach (DataRow rw in dtClientSite.Rows)
                //        {
                //            if (TaggingClientName == rw["ClientName"].ToString())
                //            {
                //                taggingClientID = rw["iID"].ToString();
                //                break;
                //            }
                //        }
                //        HyperLink hyper1 = grow.Cells[2].Controls[0] as HyperLink;


                //        string fName = hyper.Text;

                //        string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
                //        context = new ClientContext(siteUrl);
                //        SecureString passWord = new SecureString();
                //        foreach (var c in "w!ldWind36") passWord.AppendChar(c);
                //        context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
                //        Web site = context.Web;
                //        text = ddlFolderName.SelectedItem.Text;
                //        Folder subFoldercol1 = site.GetFolderByServerRelativeUrl("Anziano" + "/" + text);

                //        Microsoft.SharePoint.Client.File subfile = site.GetFileByServerRelativeUrl("Anziano" + "/" + text);
                //        ListCollection collList = site.Lists;
                //        int filecount = 0;
                //        //FolderCollection fcolection1 = subFoldercol1.Folders;
                //        Microsoft.SharePoint.Client.FileCollection fcolection1 = subFoldercol1.Files;
                //        context.Load(fcolection1);
                //        context.Load(collList);
                //        context.ExecuteQuery();

                //        for (int i = 0; i < fcolection1.Count; i++)
                //        {
                //            Microsoft.SharePoint.Client.File f1 = fcolection1[i];
                //            filecount = fcolection1.Count();
                //            string AnziFileName = f1.Name.ToString();

                //            //int a = fcolection1[i];
                //            string FolderPath = "Documents taxonomy";

                //            Microsoft.SharePoint.Client.ListItem item = null;

                //            Microsoft.SharePoint.Client.List docs = null;

                //            if (fName == AnziFileName)
                //            {
                //                FileCopy(f1);
                //                //string path = @"D:\PRACTICE\gp\" + AnziFileName;
                //                string path1 = Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\" + AnziFileName;

                //                // string exte = System.IO.Path.GetExtension(path);

                //                string AnziFile = System.IO.Path.GetFileNameWithoutExtension(path1);
                //                string exte = System.IO.Path.GetExtension(path1);

                //                if (ChkTest.Checked == true)
                //                {
                //                    // CopyFilenew(FolderPath, AnziFile + "_Test" + exte, path);
                //                    #region copynewfile
                //                    //string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
                //                    //ClientContext context = new ClientContext(siteUrl);
                //                    //SecureString passWord = new SecureString();
                //                    //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
                //                    //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
                //                    //Web site = context.Web;

                //                    byte[] bytes = System.IO.File.ReadAllBytes(path);
                //                    System.IO.Stream stream = new System.IO.MemoryStream(bytes);
                //                    string filename = Path.GetFileName(AnziFile + "_Test" + exte);
                //                    Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                //                    FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFile + "_Test" + exte), Overwrite = true };
                //                    docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                //                    Microsoft.SharePoint.Client.File uploadFile = currentRunFolder.Files.Add(newFile);

                //                    //currentRunFolder.Update();
                //                    ////Response.Write("<br/>," + vSourcrFile);
                //                    ////Response.Write("<br/>," + destFilename);
                //                    ////Response.Write("<br/>," + FolderPath);
                //                    //context.ExecuteQuery();
                //                    context.Load(uploadFile);
                //                    context.Load(docs);
                //                    context.ExecuteQuery();

                //                    context.Load(uploadFile.ListItemAllFields);
                //                    context.ExecuteQuery();

                //                    // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

                //                    item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
                //                    context.Load(item);
                //                    context.ExecuteQuery();

                //                    #endregion

                //                    Success++;
                //                    Array.ForEach(Directory.GetFiles(Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\"), System.IO.File.Delete);

                //                    DataRow rowMail = dtMail1.NewRow();
                //                    rowMail["FileName"] = fName.ToString();
                //                    //dtMail1.Rows.Add(rowMail);

                //                    rowMail["AnzianoId"] = "";
                //                    rowMail["LegalEntity"] = "";
                //                    rowMail["HouseHold"] = "";
                //                    if (chkleft.Checked)
                //                    {
                //                        rowMail["NameMatch"] = "Y";
                //                    }

                //                    if (chkleft.Checked)
                //                    {
                //                        if (chkright.Checked)
                //                            rowMail["Verified"] = "Y";
                //                        else
                //                            rowMail["Verified"] = "N";

                //                    }

                //                    rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFile + "_Test" + exte;
                //                    dtMail1.Rows.Add(rowMail);

                //                }

                //                else
                //                {
                //                    //CopyFilenew(FolderPath, AnziFileName, path);
                //                    #region copynewfile


                //                    byte[] bytes = System.IO.File.ReadAllBytes(path);
                //                    System.IO.Stream stream = new System.IO.MemoryStream(bytes);
                //                    string filename = Path.GetFileName(AnziFileName);
                //                    Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderPath);
                //                    FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(AnziFileName), Overwrite = true };
                //                    docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                //                    Microsoft.SharePoint.Client.File uploadFile = currentRunFolder.Files.Add(newFile);


                //                    context.Load(uploadFile);
                //                    context.Load(docs);
                //                    context.ExecuteQuery();

                //                    context.Load(uploadFile.ListItemAllFields);
                //                    context.ExecuteQuery();

                //                    // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

                //                    item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
                //                    context.Load(item);
                //                    context.ExecuteQuery();

                //                    #endregion

                //                    Success++;

                //                    string[] words;


                //                    DataRow rowMail = dtMail1.NewRow();
                //                    rowMail["FileName"] = fName.ToString();
                //                    //dtMail1.Rows.Add(rowMail);

                //                    rowMail["AnzianoId"] = "";
                //                    rowMail["LegalEntity"] = "";
                //                    rowMail["HouseHold"] = "";
                //                    if (chkleft.Checked)
                //                    {
                //                        rowMail["NameMatch"] = "Y";
                //                    }
                //                    if (chkleft.Checked)
                //                    {
                //                        if (chkright.Checked)
                //                            rowMail["Verified"] = "Y";

                //                    }
                //                    rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + fName.ToString();
                //                    dtMail1.Rows.Add(rowMail);


                //                    //dtMail1.Columns.Add("Destination");

                //                    f1.DeleteObject();
                //                    context.ExecuteQuery();
                //                }



                //                #region Tagging 10/12/2016


                //                //                                List docs = context.Web.Lists.GetByTitle("Documents taxonomy");
                //                //                                context.Load(docs);
                //                //                                context.ExecuteQuery();
                //                //                                CamlQuery _CamlQuery = new CamlQuery();
                //                //                                _CamlQuery.ViewXml =
                //                //                                   @"<View Scope='RecursiveAll'>
                //                //                   <Query>
                //                //                   <OrderBy UseIndexForOrderBy = 'TRUE'> <FieldRef Name='Modified' Ascending='False' /></OrderBy>
                //                //                      </Query>
                //                //                      <RowLimit>10</RowLimit>
                //                //                      </View>";

                //                //                                Microsoft.SharePoint.Client.ListItemCollection _ListItemCollection = docs.GetItems(_CamlQuery);
                //                //                                context.Load(_ListItemCollection);
                //                //                                context.ExecuteQuery();



                //                //                                Microsoft.SharePoint.Client.ListItem item = null;
                //                //                                foreach (Microsoft.SharePoint.Client.ListItem listItem in _ListItemCollection)
                //                //                                {
                //                //                                    if (AnziFile + "_Test" + exte == listItem["FileLeafRef"].ToString())
                //                //                                    {
                //                //                                        item = listItem;
                //                //                                        break;
                //                //                                    }
                //                //                                    else if (AnziFileName == listItem["FileLeafRef"].ToString())
                //                //                                    {
                //                //                                        item = listItem;
                //                //                                        break;
                //                //                                    }
                //                //                                }




                //                //                                context.ExecuteQuery();

                //                TaxonomyFieldValue taxonomyFieldValueClient = new TaxonomyFieldValue();
                //                TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
                //                TaxonomyFieldValue taxonomyFieldValueYear = new TaxonomyFieldValue();

                //                taxonomyFieldValuePath.TermGuid = PathTaggingID;
                //                taxonomyFieldValuePath.Label = PathTaggingName;

                //                taxonomyFieldValueClient.TermGuid = taggingClientID;
                //                taxonomyFieldValueClient.Label = TaggingClientName;

                //                Taggingyear = ddlYear.SelectedItem.Text;
                //                TaggingYearID = string.Empty;
                //                if (ddlYear.SelectedValue != "0")
                //                {


                //                    if (vIsYear == "True")
                //                    {
                //                        taxonomyFieldValueYear.TermGuid = TaggingYearID;
                //                        taxonomyFieldValueYear.Label = Taggingyear;
                //                        item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;
                //                    }
                //                    else
                //                    {

                //                        taxonomyFieldValueYear.TermGuid = TaggingYearID;
                //                        taxonomyFieldValueYear.Label = Taggingyear;
                //                        item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

                //                    }
                //                }
                //                else
                //                {
                //                    taxonomyFieldValueYear.TermGuid = TaggingYearID;
                //                    taxonomyFieldValueYear.Label = Taggingyear;
                //                    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
                //                }


                //                item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;
                //                item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;
                //                item.Update();
                //                docs.Update();
                //                context.ExecuteQuery();
                //                lblMsg.Text = "File Copied Successfully";
                //                break;

                //                #endregion


                //                // DownloadFileUsingFileStream(f, siteUrl+"/" +"Documents taxonomy" +"/");
                //            }
                //        }

                //    }





                //    // FileCopy(fName);
                //}

                //#endregion

                EmailBody();
                SendEmail(emailBody.ToString(), "Fund Admin Workflow Email confirmation", "skane@infograte.com", "");
                lblMsg.Text = "File Copied Successfully";

            }

            #region Not in Use
            //else if (ddlYear.Visible == true && ddlYear.SelectedValue.ToString() != "0")
            //            {
            //                DataTable dtFolderData = (DataTable)ViewState["dtFolderData"];
            //                DataTable dtClient = (DataTable)ViewState["dtSiteClientList"];

//                string PathTaggingName = string.Empty;
            //                string PathTaggingID = string.Empty;
            //                string vIsYear = string.Empty;
            //                foreach (DataRow rw in dtFolderData.Rows)
            //                {
            //                    if (folderPath == rw["FolderPath"].ToString())
            //                    {
            //                        PathTaggingName = rw["Tag"].ToString();
            //                        vIsYear = rw["OnPortal"].ToString();
            //                        break;
            //                    }

//                }

//                DataSet dsTaxonomyclientPortal = (DataSet)ViewState["dsTaxonomyclientPortal"];
            //                DataTable dtDocumentType = dsTaxonomyclientPortal.Tables[1];
            //                DataTable dtClientSite = dsTaxonomyclientPortal.Tables[0];
            //                DataTable dtYear = dsTaxonomyclientPortal.Tables[2];


//                foreach (DataRow rw in dtDocumentType.Rows)
            //                {
            //                    if (rw["DocumentType"].ToString() == PathTaggingName)
            //                    {
            //                        PathTaggingID = rw["iID"].ToString();
            //                        break;
            //                    }
            //                }

//                string Taggingyear = ddlYear.SelectedItem.Text;
            //                string TaggingYearID = string.Empty;
            //                if (ddlYear.SelectedValue != "0")
            //                {
            //                    if (vIsYear == "True")
            //                    {
            //                        foreach (DataRow rw in dtYear.Rows)
            //                        {
            //                            if (rw["Year"].ToString() == Taggingyear)
            //                                TaggingYearID = rw["iID"].ToString();
            //                        }

//                    }

//                }

//                foreach (GridViewRow grow in gvList.Rows)
            //                {


//                    string cell5 = grow.Cells[5].Text;//household

//                    string LegalEntity = grow.Cells[4].Text;//LegalEntityName

//                    string AnzianoId = grow.Cells[3].Text;//AnzianoId

//                    DropDownList ddList = (DropDownList)grow.FindControl("dtclients");
            //                    string cell6 = ddList.SelectedItem.Text;//client Name

//                    CheckBox chkleft = (CheckBox)grow.FindControl("chkbSelectBatch");

//                    CheckBox chkright = (CheckBox)grow.FindControl("chkVerify");


//                    if (chkleft.Checked == true)
            //                    {

//                        if (ddList.SelectedValue.ToString() != "0")
            //                        {
            //                            //string CName = grow.Cells[6].Text;
            //                            string TaggingClientName = cell6;
            //                            string taggingClientID = string.Empty;

//                            //string iClientID = string.Empty;
            //                            foreach (DataRow rw in dtClientSite.Rows)
            //                            {
            //                                if (TaggingClientName == rw["ClientName"].ToString())
            //                                {
            //                                    taggingClientID = rw["iID"].ToString();
            //                                    break;
            //                                }
            //                            }
            //                            HyperLink hyper = grow.Cells[2].Controls[0] as HyperLink;


//                            string fName = hyper.Text;

//                            string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
            //                            context = new ClientContext(siteUrl);
            //                            SecureString passWord = new SecureString();
            //                            foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //                            context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            //                            Web site = context.Web;
            //                            text = ddlFolderName.SelectedItem.Text;
            //                            Folder subFoldercol = site.GetFolderByServerRelativeUrl("Anziano" + "/" + text);

//                            //Microsoft.SharePoint.Client.File subfile = site.GetFileByServerRelativeUrl("Anziano" + "/" + text);
            //                            ListCollection collList = site.Lists;
            //                            int filecount = 0;
            //                            // FolderCollection fcolection = subFoldercol.Folders;
            //                            Microsoft.SharePoint.Client.FileCollection fcolection = subFoldercol.Files;
            //                            context.Load(fcolection);
            //                            context.Load(collList);
            //                            context.ExecuteQuery();

//                            for (int i = 0; i < fcolection.Count; i++)
            //                            //foreach (Microsoft.SharePoint.Client.File f in fcolection)
            //                            {
            //                                Microsoft.SharePoint.Client.File f = fcolection[i];
            //                                // filecount = fcolection.Count();
            //                                string AnziFileName = f.Name.ToString();

//                                // int a = fcolection[i];
            //                                string FolderPath = "Documents taxonomy";
            //                                if (fName == AnziFileName)
            //                                {
            //                                    FileCopy(f);
            //                                    //string path = @"D:\PRACTICE\gp\" + AnziFileName;
            //                                    string path = Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\" + AnziFileName;

//                                    // string exte = System.IO.Path.GetExtension(path);

//                                    string AnziFile = System.IO.Path.GetFileNameWithoutExtension(path);
            //                                    string exte = System.IO.Path.GetExtension(path);

//                                    if (ChkTest.Checked == true)
            //                                    {
            //                                        CopyFilenew(FolderPath, AnziFile + "_Test" + exte, path);
            //                                        Success++;
            //                                        //Array.ForEach(Directory.GetFiles(Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\"), System.IO.File.Delete);

//                                        DataRow rowMail = dtMail.NewRow();
            //                                        rowMail["FileName"] = fName.ToString();
            //                                        //dtMail.Rows.Add(rowMail);

//                                        rowMail["AnzianoId"] = AnzianoId.ToString();
            //                                        rowMail["LegalEntity"] = LegalEntity.ToString();
            //                                        rowMail["HouseHold"] = cell5.ToString();
            //                                        if (chkleft.Checked)
            //                                        {
            //                                            rowMail["NameMatch"] = "Y";
            //                                        }
            //                                        if (chkleft.Checked)
            //                                        {
            //                                            if (chkright.Checked)
            //                                                rowMail["Verified"] = "Y";
            //                                            else
            //                                                rowMail["Verified"] = "N";

//                                        }

//                                        rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFile + "_Test" + exte;
            //                                        dtMail.Rows.Add(rowMail);

//                                    }

//                                    else
            //                                    {
            //                                        CopyFilenew(FolderPath, AnziFileName, path);

//                                        Success++;

//                                        //string[] words;


//                                        DataRow rowMail = dtMail.NewRow();
            //                                        rowMail["FileName"] = fName.ToString();
            //                                        //dtMail.Rows.Add(rowHouseHold);

//                                        rowMail["AnzianoId"] = AnzianoId.ToString();
            //                                        rowMail["LegalEntity"] = LegalEntity.ToString();
            //                                        rowMail["HouseHold"] = cell5.ToString();
            //                                        if (chkleft.Checked)
            //                                        {
            //                                            rowMail["NameMatch"] = "Y";
            //                                        }
            //                                        if (chkleft.Checked)
            //                                        {
            //                                            if (chkright.Checked)
            //                                                rowMail["Verified"] = "Y";
            //                                            else
            //                                                rowMail["Verified"] = "N";

//                                        }
            //                                        rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + fName.ToString();
            //                                        dtMail.Rows.Add(rowMail);


//                                        //dtMail.Columns.Add("Destination");

//                                        f.DeleteObject();
            //                                        context.ExecuteQuery();
            //                                    }


//                                    // f.DeleteObject();
            //                                    // System.IO.File.Delete(path);
            //                                    //context.ExecuteQuery();


//                                    //System.IO.File.Delete(siteUrl + "/" + "Anziano" + "/" + text + "/" + AnziFileName);
            //                                    #region Tagging 10/12/2016


//                                    List docs = context.Web.Lists.GetByTitle("Documents taxonomy");
            //                                    context.Load(docs);
            //                                    context.ExecuteQuery();
            //                                    CamlQuery _CamlQuery = new CamlQuery();
            //                                    _CamlQuery.ViewXml =
            //                                       @"<View Scope='RecursiveAll'>
            //                   <Query>
            //                   <OrderBy UseIndexForOrderBy = 'TRUE'> <FieldRef Name='Modified' Ascending='False' /></OrderBy>
            //                      </Query>
            //                      <RowLimit>10</RowLimit>
            //                      </View>";

//                                    Microsoft.SharePoint.Client.ListItemCollection _ListItemCollection = docs.GetItems(_CamlQuery);
            //                                    context.Load(_ListItemCollection);
            //                                    context.ExecuteQuery();



//                                    Microsoft.SharePoint.Client.ListItem item = null;
            //                                    foreach (Microsoft.SharePoint.Client.ListItem listItem in _ListItemCollection)
            //                                    {
            //                                        if (AnziFile + "_Test" + exte == listItem["FileLeafRef"].ToString())
            //                                        {
            //                                            item = listItem;
            //                                            break;
            //                                        }
            //                                        else if (AnziFileName == listItem["FileLeafRef"].ToString())
            //                                        {
            //                                            item = listItem;
            //                                            break;
            //                                        }
            //                                    }





//                                    context.Load(_ListItemCollection);
            //                                    context.ExecuteQuery();






//                                    //context.ExecuteQuery();

//                                    TaxonomyFieldValue taxonomyFieldValueClient = new TaxonomyFieldValue();
            //                                    TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            //                                    TaxonomyFieldValue taxonomyFieldValueYear = new TaxonomyFieldValue();

//                                    taxonomyFieldValuePath.TermGuid = PathTaggingID;
            //                                    taxonomyFieldValuePath.Label = PathTaggingName;

//                                    taxonomyFieldValueClient.TermGuid = taggingClientID;
            //                                    taxonomyFieldValueClient.Label = TaggingClientName;

//                                    //string Taggingyear = ddlYear.SelectedItem.Text;
            //                                    //string TaggingYearID = string.Empty;
            //                                    if (ddlYear.SelectedValue != "0")
            //                                    {


//                                        if (vIsYear == "True")
            //                                        {
            //                                            taxonomyFieldValueYear.TermGuid = TaggingYearID;
            //                                            taxonomyFieldValueYear.Label = Taggingyear;
            //                                            item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;
            //                                        }
            //                                        else
            //                                        {

//                                            taxonomyFieldValueYear.TermGuid = TaggingYearID;
            //                                            taxonomyFieldValueYear.Label = Taggingyear;
            //                                            item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

//                                        }
            //                                    }
            //                                    else
            //                                    {
            //                                        taxonomyFieldValueYear.TermGuid = TaggingYearID;
            //                                        taxonomyFieldValueYear.Label = Taggingyear;
            //                                        item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
            //                                    }


//                                    //else
            //                                    //{

//                                    //    taxonomyFieldValueYear.TermGuid = TaggingYearID;
            //                                    //    taxonomyFieldValueYear.Label = Taggingyear;
            //                                    //    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

//                                    //}

//                                    item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;
            //                                    item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;
            //                                    item.Update();
            //                                    docs.Update();
            //                                    context.ExecuteQuery();
            //                                    //lblMsg.Text = "File Copied Successfully";
            //                                    break;

//                                    #endregion


//                                    //DownloadFileUsingFileStream(f, siteUrl+"/" +"Documents taxonomy" +"/");
            //                                }
            //                            }

//                        }
            //                        else
            //                        {
            //                            //lblError.Text = "No Files Copied";
            //                        }

//                    }

//                    else
            //                    {
            //                        //lblError.Text = "No Files Copied";
            //                    }



//                }


//                #region 2nd Grid

//                foreach (GridViewRow grow in gvList1.Rows)
            //                {
            //                    DropDownList ddList = (DropDownList)grow.FindControl("ddlClients1");

//                    CheckBox chkleft = (CheckBox)grow.FindControl("chkbSelectBatch");

//                    CheckBox chkright = (CheckBox)grow.FindControl("chkVerify");

//                    HyperLink hyper = grow.Cells[2].Controls[0] as HyperLink;

//                    if (chkleft.Checked)
            //                    {

//                        // string CName = grow.Cells[6].Text;
            //                        string TaggingClientName = ddList.SelectedItem.Text;
            //                        string taggingClientID = string.Empty;

//                        string iClientID = string.Empty;
            //                        foreach (DataRow rw in dtClientSite.Rows)
            //                        {
            //                            if (TaggingClientName == rw["ClientName"].ToString())
            //                            {
            //                                taggingClientID = rw["iID"].ToString();
            //                                break;
            //                            }
            //                        }
            //                        HyperLink hyper1 = grow.Cells[2].Controls[0] as HyperLink;


//                        string fName = hyper.Text;

//                        string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
            //                        context = new ClientContext(siteUrl);
            //                        SecureString passWord = new SecureString();
            //                        foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //                        context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            //                        Web site = context.Web;
            //                        text = ddlFolderName.SelectedItem.Text;
            //                        Folder subFoldercol1 = site.GetFolderByServerRelativeUrl("Anziano" + "/" + text);

//                        Microsoft.SharePoint.Client.File subfile = site.GetFileByServerRelativeUrl("Anziano" + "/" + text);
            //                        ListCollection collList = site.Lists;
            //                        int filecount = 0;
            //                        //FolderCollection fcolection1 = subFoldercol1.Folders;
            //                        Microsoft.SharePoint.Client.FileCollection fcolection1 = subFoldercol1.Files;
            //                        context.Load(fcolection1);
            //                        context.Load(collList);
            //                        context.ExecuteQuery();

//                        for (int i = 0; i < fcolection1.Count; i++)
            //                        {
            //                            Microsoft.SharePoint.Client.File f1 = fcolection1[i];
            //                            filecount = fcolection1.Count();
            //                            string AnziFileName = f1.Name.ToString();

//                            //int a = fcolection1[i];
            //                            string FolderPath = "Documents taxonomy";
            //                            if (fName == AnziFileName)
            //                            {
            //                                FileCopy(f1);
            //                                //string path = @"D:\PRACTICE\gp\" + AnziFileName;
            //                                string path = Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\" + AnziFileName;

//                                // string exte = System.IO.Path.GetExtension(path);

//                                string AnziFile = System.IO.Path.GetFileNameWithoutExtension(path);
            //                                string exte = System.IO.Path.GetExtension(path);

//                                if (ChkTest.Checked == true)
            //                                {
            //                                    CopyFilenew(FolderPath, AnziFile + "_Test" + exte, path);
            //                                    Success++;
            //                                    Array.ForEach(Directory.GetFiles(Server.MapPath("~/") + @"ExcelTemplate\AnzianoFileCopy\"), System.IO.File.Delete);

//                                    DataRow rowMail = dtMail1.NewRow();
            //                                    rowMail["FileName"] = fName.ToString();
            //                                    //dtMail1.Rows.Add(rowMail);

//                                    rowMail["AnzianoId"] = "";
            //                                    rowMail["LegalEntity"] = "";
            //                                    rowMail["HouseHold"] = "";
            //                                    if (chkleft.Checked)
            //                                    {
            //                                        rowMail["NameMatch"] = "Y";
            //                                    }

//                                    if (chkleft.Checked)
            //                                    {
            //                                        if (chkright.Checked)
            //                                            rowMail["Verified"] = "Y";
            //                                        else
            //                                            rowMail["Verified"] = "N";

//                                    }

//                                    rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + AnziFile + "_Test" + exte;
            //                                    dtMail1.Rows.Add(rowMail);

//                                }

//                                else
            //                                {
            //                                    CopyFilenew(FolderPath, AnziFileName, path);

//                                    Success++;

//                                    string[] words;


//                                    DataRow rowMail = dtMail1.NewRow();
            //                                    rowMail["FileName"] = fName.ToString();
            //                                    //dtMail1.Rows.Add(rowMail);

//                                    rowMail["AnzianoId"] = "";
            //                                    rowMail["LegalEntity"] = "";
            //                                    rowMail["HouseHold"] = "";
            //                                    if (chkleft.Checked)
            //                                    {
            //                                        rowMail["NameMatch"] = "Y";
            //                                    }
            //                                    if (chkleft.Checked)
            //                                    {
            //                                        if (chkright.Checked)
            //                                            rowMail["Verified"] = "Y";

//                                    }
            //                                    rowMail["Destination"] = siteUrl + "/" + "Documents taxonomy" + "/" + fName.ToString();
            //                                    dtMail1.Rows.Add(rowMail);


//                                    //dtMail1.Columns.Add("Destination");

//                                    f1.DeleteObject();
            //                                    context.ExecuteQuery();
            //                                }



//                                #region Tagging 10/12/2016


//                                List docs = context.Web.Lists.GetByTitle("Documents taxonomy");
            //                                context.Load(docs);
            //                                context.ExecuteQuery();
            //                                CamlQuery _CamlQuery = new CamlQuery();
            //                                _CamlQuery.ViewXml =
            //                                   @"<View Scope='RecursiveAll'>
            //                   <Query>
            //                   <OrderBy UseIndexForOrderBy = 'TRUE'> <FieldRef Name='Modified' Ascending='False' /></OrderBy>
            //                      </Query>
            //                      <RowLimit>10</RowLimit>
            //                      </View>";

//                                Microsoft.SharePoint.Client.ListItemCollection _ListItemCollection = docs.GetItems(_CamlQuery);
            //                                context.Load(_ListItemCollection);
            //                                context.ExecuteQuery();



//                                Microsoft.SharePoint.Client.ListItem item = null;
            //                                foreach (Microsoft.SharePoint.Client.ListItem listItem in _ListItemCollection)
            //                                {
            //                                    if (AnziFile + "_Test" + exte == listItem["FileLeafRef"].ToString())
            //                                    {
            //                                        item = listItem;
            //                                        break;
            //                                    }
            //                                    else if (AnziFileName == listItem["FileLeafRef"].ToString())
            //                                    {
            //                                        item = listItem;
            //                                        break;
            //                                    }
            //                                }




//                                context.ExecuteQuery();

//                                TaxonomyFieldValue taxonomyFieldValueClient = new TaxonomyFieldValue();
            //                                TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            //                                TaxonomyFieldValue taxonomyFieldValueYear = new TaxonomyFieldValue();

//                                taxonomyFieldValuePath.TermGuid = PathTaggingID;
            //                                taxonomyFieldValuePath.Label = PathTaggingName;

//                                taxonomyFieldValueClient.TermGuid = taggingClientID;
            //                                taxonomyFieldValueClient.Label = TaggingClientName;

//                                Taggingyear = ddlYear.SelectedItem.Text;
            //                                TaggingYearID = string.Empty;
            //                                if (ddlYear.SelectedValue != "0")
            //                                {


//                                    if (vIsYear == "True")
            //                                    {
            //                                        taxonomyFieldValueYear.TermGuid = TaggingYearID;
            //                                        taxonomyFieldValueYear.Label = Taggingyear;
            //                                        item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;
            //                                    }
            //                                    else
            //                                    {

//                                        taxonomyFieldValueYear.TermGuid = TaggingYearID;
            //                                        taxonomyFieldValueYear.Label = Taggingyear;
            //                                        item["p9cafa43d635492cb87a8a60d0ebb191"] = "";

//                                    }
            //                                }
            //                                else
            //                                {
            //                                    taxonomyFieldValueYear.TermGuid = TaggingYearID;
            //                                    taxonomyFieldValueYear.Label = Taggingyear;
            //                                    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
            //                                }


//                                item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;
            //                                item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;
            //                                item.Update();
            //                                docs.Update();
            //                                context.ExecuteQuery();
            //                                lblMsg.Text = "File Copied Successfully";
            //                                break;

//                                #endregion


//                                // DownloadFileUsingFileStream(f, siteUrl+"/" +"Documents taxonomy" +"/");
            //                            }
            //                        }

//                    }





//                    // FileCopy(fName);
            //                }

//                #endregion

//                EmailBody();
            //                SendEmail(emailBody.ToString(), "Fund Admin Workflow Email confirmation", "skane@infograte.com", "");
            //                lblMsg.Text = "File Copied Successfully";
            //                //lblError.Text = "Please Select Year Tag";
            //            }
            #endregion
            else
            {
                lblError.Text = "Please Select Year";
            }
        }


        else
        {
            lblError.Text = "No Document Tag Selected";
        }
        emailBody.Clear();

    }

    private void addToEmailBody(string str)
    {
        emailBody.Append(str);
    }

    public void EmailBody()
    {

        string sourceFileName;

        string destFileName;
        addToEmailBody("----------------------------------------------------------------------------------<br />");
        addToEmailBody("<b>Selected Options -</b> <br />");
        addToEmailBody("Anziano Folder: ");
        addToEmailBody(ddlFolderName.Text + "<br />");
        addToEmailBody("Client Portal Tag: ");
        if (ddlYear.Visible)
            //if (ddlYear.SelectedItem.Text != "Select")
            addToEmailBody(ddlPortalPath.Text + "/" + ddlYear.SelectedItem.Text + "<br />");
        else
            addToEmailBody(ddlPortalPath.Text + "<br />");
        if (ChkTest.Checked)
        {
            addToEmailBody("Test Mode" + "<br/>");
        }
        addToEmailBody("----------------------------------------------------------------------------------<br /><br />");
        dtMail.DefaultView.Sort = "HouseHold ASC";//added on 02/07/2018 by brijesh 
        dtMail = dtMail.DefaultView.ToTable();//added on 02/07/2018 by brijesh 
        addToEmailBody("<table border=\"1\" cellspacing=\"0\">");
        addToEmailBody("<tr><th>File Name</th><th>Anziano ID</th><th>Legal Entity</th><th>Household</th><th>Name Match</th><th>Verified</th><th>Destination</th></tr>\r\n");
        foreach (DataRow r in dtMail.Rows)
        {
            sourceFileName = r["FileName"].ToString();

            //string Anzianoid = r["Anzianoid"].ToString();

            string Anziid = r["AnzianoId"].ToString();

            if (r["AnzianoId"].ToString() == "")
            {
                destFileName = sourceFileName;
            }
            else
            {
                destFileName = r["FileName"].ToString().Replace(r["AnzianoId"].ToString() + "_", "");
            }
            addToEmailBody("<tr><td>" + r["FileName"].ToString() + "</td><td>" + r["Anzianoid"].ToString() + "</td><td>" + r["LegalEntity"].ToString() + "</td><td>" + r["HouseHold"].ToString() + "</td><td>" + r["NameMatch"].ToString() + "</td><td>" + r["Verified"].ToString() + "</td><td><a href='" + r["Destination"].ToString() + "'>" + destFileName + "</a></td></tr>\r\n");

        }

        //foreach (DataRow r in dtMail1.Rows)
        //{
        //    addToEmailBody("<tr><td>" + r["FileName"].ToString() + "</td><td>" + r["Anzianoid"].ToString() + "</td><td>" + r["LegalEntity"].ToString() + "</td><td>" + r["HouseHold"].ToString() + "</td><td>" + r["NameMatch"].ToString() + "</td><td>" + r["Verified"].ToString() + "</td><td><a href='" + r["Destination"].ToString() + "'>" + r["FileName"].ToString() + "</a></td></tr>\r\n");

        //}
        addToEmailBody("</table>\r\n");
        addToEmailBody("<b>No. Of Files Copied: </b>" + Success);

        // System.IO.File.WriteAllText(Test + "Email.html", emailBody.ToString());
    }

    public void SendEmail(string mailmessage, string subject, string mailTo, string Attachment1)
    {
        try
        {

            MailMessage myMessage = new MailMessage();
            // SmtpClient SMTPSERVER = new SmtpClient();

            string EmailID = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Password = AppLogic.GetParam(AppLogic.ConfigParam.Password);
            string SMTPHost = AppLogic.GetParam(AppLogic.ConfigParam.SMTPHost);
            // string ToEmailIDs = AppLogic.GetParam(AppLogic.ConfigParam.ToEmailIDs);

            string ToEmailIDs = "skane@infograte.com|vatwood@greshampartners.com|ereinke@greshampartners.com";   // change by abhi, ref by email
            int Port = Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.Port));

            myMessage.From = new MailAddress(EmailID, "Fund Admin Workflow");
             string[] strTo = ToEmailIDs.Split('|');


             for (int i = 0; i < strTo.Length; i++)
             {
                 if (strTo[i] != "")
                 {
                     myMessage.To.Add(new MailAddress(strTo[i]));
                 }
             }
            //myMessage.CC.Add("GBhagia@infograte.com");
            //myMessage.Bcc.Add("dshah@webdevinc.net");
           // myMessage.To.Add(new MailAddress(ToEmailIDs));

            string Bcc = txtMailBCC.Text;
            if (Bcc != "")
            {
                Session["Mail"] = txtMailBCC.Text;

                string MailCC = Session["Mail"].ToString();
                string[] strToBcc = MailCC.Split(';');
                foreach (string bccEmailId in strToBcc)
                {
                    myMessage.CC.Add(new MailAddress(bccEmailId)); //Adding Multiple CC email Id
                }
            }

            myMessage.Bcc.Add("auto-emails@infograte.com");
            // myMessage.Bcc.Add("skane@infograte.com");
            //for (int i = 0; i < strToBcc.Length; i++)
            //{
            //    if (strToBcc[i] != "")
            //    {
            //        myMessage.CC.Add(new MailAddress(strTo[i]));
            //    }
            //}

            //myMessage.Bcc.Add(strTo[0]);
            //myMessage.Bcc.Add(strTo[1]);

            //myMessage.Bcc.Add(strTo[2]);
            //myMessage.Bcc.Add(strTo[3]);

            //myMessage.Bcc.Add(strTo[4]);

            //myMessage.Bcc.Add("svaitya@webdevinc.net");
            myMessage.Subject = subject;

            //if (Attachment1 != "")
            //    myMessage.Attachments.Add(new Attachment(Attachment1));

            myMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
            myMessage.Body = mailmessage;

            myMessage.IsBodyHtml = true;

            SmtpClient SMTPSERVER = new SmtpClient(SMTPHost, Port);
            SMTPSERVER.DeliveryMethod = SmtpDeliveryMethod.Network;
            //SMTPSERVER.Host = SMTPHost;
            //SMTPSERVER.Port = Port;


            //SMTPSERVER.EnableSsl = false; for office 365
            SMTPSERVER.EnableSsl = true;
            // smtp.EnableSsl = true;
            SMTPSERVER.UseDefaultCredentials = true;
            System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(EmailID, Password);
            SMTPSERVER.Credentials = basicAuthenticationInfo;
            SMTPSERVER.Send(myMessage);
            //lblEmail.Text = "Send";
            myMessage.Dispose();
            myMessage = null;
            SMTPSERVER = null;
            mailmessage = null;
        }
        catch (Exception ex)
        {
            string strDescription = "Error sending Mail :" + ex.Message.ToString();
            //commented on 12_4_2018 Jscalise nolonger in process
               lblError.Text = lblError.Text + "  ," + "Send Error" + ex.Message.ToString();
            // lblEmail.Text = "Send Error" + ex.Message.ToString();
            
            //LogMessage(sw, strDescription);
        }
    }

    protected void ddlPortalPath_SelectedIndexChanged(object sender, EventArgs e)
    {

        lblMsg.Text = "";
        lblError.Text = "";
        DataTable FolderData = (DataTable)ViewState["dtFolderData"];
        string valuePath = ddlPortalPath.SelectedItem.Text;
        if (valuePath != "Select")
        {
            foreach (DataRow row in FolderData.Rows)
            {
                //dt.Columns.Add("FolderPath");
                //dt.Columns.Add("OnPortal");
                //dt.Columns.Add("Tag");
                if (valuePath == row["FolderPath"].ToString())
                {
                    if (row["OnPortal"].ToString() == "True")
                        ddlYear.Visible = true;
                    break;
                }
                else
                {
                    ddlYear.Visible = false;
                }
            }
        }
        else
        {
            ddlYear.Visible = false;
        }
    }

    protected void ddlYear_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMsg.Text = "";
        lblError.Text = "";
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //commented on 12_4_2018 Jscalise nolonger in process
            //lblEmail.Text = ListEmail;
            //lblEmail.Visible = true;
            //Label4.Text = ListEmail;
            newFolderstructure();
            lblMsg.Text = "";
            lblError.Text = "";
            //newFolderstructuretemp();
            FolderData = sh.getSPList();
            // BindFolderPathddl();
            txtMailBCC.Text = "";

            dstaxonomy = sh.getTaxonomyClientPortal();

            year = dstaxonomy.Tables[2];
            dtSiteClientList = sh.getSiteClientList();
            //BindYearddl();
            Button1.Visible = false;
            lbltag.Visible = false;
            lblTest.Visible = false;
            ChkTest.Visible = false;
            ddlPortalPath.Visible = false;
            ddlYear.Visible = false;
            ViewState["dtFolderData"] = FolderData;
            ViewState["dsTaxonomyclientPortal"] = dstaxonomy;
            ViewState["dtSiteClientList"] = dtSiteClientList;


            ViewState["Year"] = year;

        }
    }
}


