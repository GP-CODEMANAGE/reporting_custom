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
using Microsoft.SharePoint.Client.Taxonomy;
using System.Text;
using System.Net.Mail;

//using List = Microsoft.SharePoint;

//public static class tablecopy
//{
//    public static DataTable ToDataTable<T>(this List<T> iList)
//    {
//        DataTable dataTable = new DataTable();
//        PropertyDescriptorCollection propertyDescriptorCollection =
//            TypeDescriptor.GetProperties(typeof(T));
//        for (int i = 0; i < propertyDescriptorCollection.Count; i++)
//        {
//            PropertyDescriptor propertyDescriptor = propertyDescriptorCollection[i];
//            Type type = propertyDescriptor.PropertyType;

//            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
//                type = Nullable.GetUnderlyingType(type);


//            dataTable.Columns.Add(propertyDescriptor.Name, type);
//        }
//        object[] values = new object[propertyDescriptorCollection.Count];
//        foreach (T iListItem in iList)
//        {
//            for (int i = 0; i < values.Length; i++)
//            {
//                values[i] = propertyDescriptorCollection[i].GetValue(iListItem);
//            }
//            dataTable.Rows.Add(values);
//        }
//        return dataTable;
//    }
//}


public partial class frmClientPortalFileCopy : System.Web.UI.Page
{
    private StringBuilder emailBody = new StringBuilder("<html>");
    string filename = string.Empty;
    string path = string.Empty;
    DB cs = new DB();
    DataTable dt = new DataTable();
    sharepoint sp = new sharepoint();

    // DataTable dtSharepint = new DataTable();
    DataTable dt2 = new DataTable();
    int count = 0;
    DataTable FolderData;     // clientPortal folderPath Datatable\
    DataSet dsTaxonomyclientPortal;   // clientPortal taxonomy data
    DataTable dtSiteClientList;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {


            btnUploadFile.Attributes.Add("onclick", "return do_totals1();");

            Button1.Attributes.Add("onclick", "return do_totals1();");

            dt.Columns.Add("BatchName");
            BindFund();
            fillHouseholdType();
            //   newFolderstructure();

            // Label2.Text =  dt.Rows.Count.ToString();
            //  newFolderstructuretemp();   // old file format structure
            //BindAllGrideView(); // using taxonomy
            //Bindgridview();

            FolderData = sp.getSPList();
            dsTaxonomyclientPortal = sp.getTaxonomyClientPortal();
            dtSiteClientList = sp.getSiteClientList();

            ViewState["dtFolderData"] = FolderData;
            ViewState["dsTaxonomyclientPortal"] = dsTaxonomyclientPortal;
            ViewState["dtSiteClientList"] = dtSiteClientList;

            BindFolderPathddl();
            BindYearddl();
            BindclientGrideView();
        }
        else
        {
            if (FileUpload1.HasFile)
            {
                Session["FileUpload1"] = FileUpload1;
                txtfilename.Text = FileUpload1.FileName;
            }
            else if (Session["FileUpload1"] != null)
            {
                FileUpload1 = (FileUpload)Session["FileUpload1"];
                txtfilename.Text = FileUpload1.FileName;
            }
        }
    }

    protected void txtAsOfDate_TextChanged(object sender, EventArgs e)
    {
        if (txtAsOfDate.Text.Trim() != "")
        {
            /*split up date into different variables strMont,strday,strYear*/
            string[] datesplit = txtAsOfDate.Text.Split('/');
            string strMonth = "";
            string strday = "";
            string strYear = "";

            int month = Convert.ToInt32(datesplit[0]);
            int day = Convert.ToInt32(datesplit[1]);
            int year = Convert.ToInt32(datesplit[2]);

            //check the length of month entered and append if necessary
            if (month.ToString().Length < 2)
                strMonth = "0" + month;
            else
                strMonth = month.ToString();

            //check the length of day entered and append if necessary
            if (day.ToString().Length < 2)
                strday = "0" + day;
            else
                strday = day.ToString();

            //check the length of year entered and append if necessary
            if (year.ToString().Length == 2)
                strYear = "20" + year;
            else
                strYear = year.ToString();
            //Reformatt date using "/"
            txtAsOfDate.Text = strMonth + "/" + strday + "/" + strYear;

        }
    }

    protected void ddlFund_SelectedIndexChanged(object sender, EventArgs e)
    {

        lblMessage.Text = "";
        lablmsg.Text = "";
        checkboxbind();



        //Bindgridview();

    }

    public void checkboxbind()
    {
        //if (ddlFund.SelectedValue.ToString() == "0")
        //{


        foreach (GridViewRow gvRow in gvList.Rows)
        {
            //string client = row["HouseHoldName"];
            ////  CheckBox cbBilling = (CheckBox)grow.Cells[5].FindControl("cbExcludeBilling");
            ////  CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");

            //if (client == Convert.ToString(gvRow.Cells[2].Text))
            //{
            CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");
            cbClient.Checked = false;
            count++;
            //}
            //else
            //{
            //    //CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");
            //    //cbClient.Checked = false;

        }
        //  }


        DataSet ds = new DataSet();
        object Fund = Convert.ToString(ddlFund.SelectedValue) == "0" ? "null" : "'" + ddlFund.SelectedValue + "'";

        object AsOfDate = Convert.ToString(txtAsOfDate.Text) == "0" ? "null" : "'" + txtAsOfDate.Text + "'";

        object dStartDate = Convert.ToString(txtStartDate.Text) == "0" ? "null" : "'" + txtStartDate.Text + "'";

        string sqlstr = "[SP_S_SHAREPOINT_POSITION_HOUSEHOLD] @FundUUID = " + Fund + ", @AsOfDate =" + AsOfDate + ",@StartDt=" + dStartDate;


        ds = cs.getDataSet(sqlstr);
        dt2 = ds.Tables[0].Copy();
        if (ds.Tables[0].Rows.Count > 0)
        {

            foreach (GridViewRow gvRow in gvList.Rows)
            {
                // CheckBox cb = gvRow.Parent.Parent.FindControl("chkbSelectBatch") as CheckBox;
                string client1 = gvRow.Cells[2].Text.Replace("O&#39;Donnell", "ODonnell");

                foreach (DataRow row in dt2.Rows)
                {
                    string client = row["HouseHoldName"].ToString().Replace("'", "");

                    //  CheckBox cbBilling = (CheckBox)grow.Cells[5].FindControl("cbExcludeBilling");
                    //  CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");
                    //string client1 = row["HouseHoldName"].ToString();



                    if (client == client1 && ddlFund.SelectedValue.ToString() != "0")
                    {

                        CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");
                        cbClient.Checked = true;
                        count++;
                    }
                    else
                    {
                        //CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");
                        //cbClient.Checked = false;

                    }

                }

            }
            //Label1.Text = count.ToString(); 
        }
        else
        {

            foreach (GridViewRow gvRow in gvList.Rows)
            {
                //string client = row["HouseHoldName"];
                ////  CheckBox cbBilling = (CheckBox)grow.Cells[5].FindControl("cbExcludeBilling");
                ////  CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");

                //if (client == Convert.ToString(gvRow.Cells[2].Text))
                //{
                CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");
                cbClient.Checked = false;
                count++;
                //}
                //else
                //{
                //    //CheckBox cbClient = (CheckBox)gvRow.Cells[1].FindControl("chkbSelectBatch");
                //    //cbClient.Checked = false;

            }
        }
    }


    protected void btnUploadFile_Click(object sender, EventArgs e)
    {
        // txtfilename.Text = "";

        UploadFiles();

    }
    public void uploadFileold()
    {
        string sourceFilename = txtfilename.Text;
        string Filenames = FileName.Text;
        string FolderName = txtFolderPath.Text;

        if (FolderName != "")
        {
            if (txtFilePath.Text != "")
            {
                try
                {
                    // string FilePath = Server.MapPath("~/") + @"ExcelTemplate\ClientPortalFileCopy\" + Filenames;
                    //Response.Write("<br/>FilePath-" + FilePath);

                    //if (System.IO.File.Exists(FilePath))
                    //{
                    //    System.IO.File.Delete(FilePath);
                    //}

                    //  string FilePath = @"D:\Test\ClientPortalFileCopy" + filename;
                    //   FileUpload1.SaveAs(FilePath);
                    //   System.IO.File.Copy(txtFilePath.Text, FilePath);

                    if (ViewState["Filename"] != null)
                    {
                        string FilePath = ViewState["Filename"].ToString();
                        foreach (GridViewRow gvRow in gvList.Rows)
                        {
                            CheckBox ClientCheckbox = (CheckBox)gvRow.FindControl("chkbSelectBatch");
                            bool CheckNameChecks = ClientCheckbox.Checked;
                            if (ClientCheckbox.Checked)
                            {
                                string clientName = gvRow.Cells[2].Text;

                                string destFilePath = @"Documents/" + clientName + "/" + FolderName;


                                CopyFilenew(destFilePath, Filenames, FilePath);

                            }

                        }

                        lablmsg.Text = "File Copied Successfully";
                        txtfilename.Text = "";
                        txtFolderPath.Text = "";
                        txtFilePath.Text = "";
                        lblMessage.Text = "";
                        FileName.Text = "";
                        BindFund();
                        txtAsOfDate.Text = "";
                    }
                    else
                    {
                        lblMessage.Text = "Please Select The File..";
                    }
                }
                catch (Exception ex)
                {
                    lblMessage.Text = ex.ToString();
                    // Label5.Text = "Upload status: The file could not be uploaded. The following error occured: " + ex.Message;
                }
            }
            else
            {
                lblMessage.Text = "Please Select The File";
            }
        }
        else
        {
            lblMessage.Text = "Please Select Folder";
        }


        //path = FileUpload1.FileName.ToString();
        //filename = FileUpload1.PostedFile.FileName.ToString();
        //String Path = Server.MapPath(FileUpload1.FileName);



    }

    protected void txtfilename_TextChanged(object sender, EventArgs e)
    {
        filename = txtfilename.Text;
    }

    protected void gvList_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    protected void gvList_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        //if (e.Row.RowType == DataControlRowType.DataRow)
        //{
        //    CheckBox chk = (CheckBox)e.Row.FindControl("chkbSelectBatch");
        //    chk.Attributes.Add("onclick", "ClearLabel();");
        //}
    }
    public void fillHouseholdType()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("SP_S_HH_Relationship_Status @ReportFlg = 1");
        ddlHouseHoldType.Items.Clear();
        ddlHouseHoldType.Items.Add(new System.Web.UI.WebControls.ListItem("ALL", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHouseHoldType.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][1])));
        }

    }
    protected void BindFund()//function to bind fund DropDown
    {
        try
        {

            string sqlstr = "SP_S_FUND_LKUP @FundTypeIdNmb = '2 , 3' , @AvailabilityIdNmb = '4, 5, 6'";//fetch data from storedprocedure
            BindDropdown(ddlFund, sqlstr, "ssi_Name", "ssi_FundId");//fucntion bind all data
            ddlFund.Items.Insert(0, "Select Fund");
            ddlFund.Items[0].Value = "0";
        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while fetching values for dropdownlists. Details: " + ex.Message;
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

    }

    public void newFolderstructure()
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
        Microsoft.SharePoint.Client.List list = context.Web.Lists.GetByTitle("Documents");

        //  Microsoft.SharePoint.Client.List list1 = context.Web.Lists.GetByTitle("Documents");

        context.Load(list);
        context.Load(list.RootFolder);
        context.Load(list.RootFolder.Folders);
        context.Load(list.RootFolder.Files);
        context.ExecuteQuery();
        FolderCollection fcol = list.RootFolder.Folders;
        //  List<string> lstFile = new List<string>()
        List<string> lstFile = new List<string>();
        Microsoft.SharePoint.Client.ListItem list1;
        //DataRow row = dt.NewRow();
        //list1 = f.ListItemAllFields;
        //context.Load(list1);
        //context.ExecuteQuery();

        dt.Columns.Add("BatchName1");
        dt.Columns.Add("onPortal");
        dt.Columns.Add("clientPortalTempFilename");
        // DataTable dt1 = lstFile.tablecopy.ToDataTable();
        foreach (Folder f in fcol)
        {

            dt.Rows.Add(f.Name.ToString());

        }
        dt.DefaultView.Sort = "BatchName1 ASC";
    }

    public void Bindgridview()
    {
        gvList.DataSource = dt;
        gvList.DataBind();
    }

    public object list1 { get; set; }

    public string DBConnectionstring { get; set; }
    protected void gvList_RowCreated(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow &&
   (e.Row.RowState == DataControlRowState.Normal ||
    e.Row.RowState == DataControlRowState.Alternate))
        {
            CheckBox chkBxSelect = (CheckBox)e.Row.Cells[1].FindControl("chkbSelectBatch");
            CheckBox chkBxHeader = (CheckBox)this.gvList.HeaderRow.FindControl("chkboxSelectAll");
            chkBxSelect.Attributes["onclick"] = string.Format
                                                   (
                                                      "javascript:ChildClick(this,'{0}');",
                                                      chkBxHeader.ClientID
                                                   );
        }
    }
    protected void txtStartDate_TextChanged(object sender, EventArgs e)
    {
        if (txtStartDate.Text.Trim() != "")
        {
            /*split up date into different variables strMont,strday,strYear*/
            string[] datesplit = txtStartDate.Text.Split('/');
            string strMonth = "";
            string strday = "";
            string strYear = "";

            int month = Convert.ToInt32(datesplit[0]);
            int day = Convert.ToInt32(datesplit[1]);
            int year = Convert.ToInt32(datesplit[2]);

            //check the length of month entered and append if necessary
            if (month.ToString().Length < 2)
                strMonth = "0" + month;
            else
                strMonth = month.ToString();

            //check the length of day entered and append if necessary
            if (day.ToString().Length < 2)
                strday = "0" + day;
            else
                strday = day.ToString();

            //check the length of year entered and append if necessary
            if (year.ToString().Length == 2)
                strYear = "20" + year;
            else
                strYear = year.ToString();
            //Reformatt date using "/"
            txtStartDate.Text = strMonth + "/" + strday + "/" + strYear;

        }
    }

    #region getdata from sharepoint and bind

    public void newFolderstructuretemp()
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

        Folder subFoldercol = site.GetFolderByServerRelativeUrl("Documents/Adams");
        ListCollection collList = site.Lists;

        FolderCollection fcolection = subFoldercol.Folders;
        context.Load(fcolection);
        context.Load(collList);
        context.ExecuteQuery();

        DataTable dtSharepint = new DataTable();
        dtSharepint.Columns.Add("Rootfolder");
        dtSharepint.Columns.Add("SubFolder");
        dtSharepint.Columns.Add("Path");
        dtSharepint.Columns.Add("subSubfolder");

        //if (dtSharepint.Columns.Count != 4)
        //    dtSharepint.Columns.Add("subSubfolder");
        foreach (Folder f in fcolection)
        {
            //Response.Write("<br/>");
            //Response.Write("<br/>," + f.Name.ToString());
            DataRow rowRootFolder = dtSharepint.NewRow();
            rowRootFolder["Rootfolder"] = f.Name.ToString();
            rowRootFolder["Path"] = f.Name.ToString();
            dtSharepint.Rows.Add(rowRootFolder);

            string foldername = "Documents/Adams/" + f.Name;

            List<string> folderlist = getFolderPath(site, context, foldername, dtSharepint);

            foreach (string FolderName in folderlist)
            {
                DataRow rowRootFolder1 = dtSharepint.NewRow();
                rowRootFolder1["SubFolder"] = FolderName;
                //   dt.Rows.Add(rowRootFolder1);


            }
            //    Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);
            ////    ListCollection collList = site.Lists;

            //    FolderCollection fcolection1 = subFolder1.Folders;
            //    context.Load(fcolection1);
            //   // context.Load(collList);
            //    context.ExecuteQuery();

            //    foreach (Folder subfolder in fcolection1)
            //    {
            //        //Response.Write("<br/>," + subfolder.Name.ToString());
            //        DataRow rowRootFolder1 = dt.NewRow();
            //        rowRootFolder1["SubFolder"] = subfolder.Name.ToString();
            //        dt.Rows.Add(rowRootFolder1);
            //    }

            //  Folder subFoldercol = site.GetFolderByServerRelativeUrl("Documents/Adams/Gresham Statements");

        }
        // dtSharepint.DefaultView.Sort = "SubFolder ASC";
        GridView1.DataSource = dtSharepint;
        GridView1.DataBind();



    }
    public List<string> getFolderPath(Web site, ClientContext context, string foldername, DataTable dtSharepint)
    {

        List<string> folderlist = new List<string>();

        List<string> subfolderlist = new List<string>();


        // foldername = "Documents/Adams/" + f.Name;
        Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);
        //    ListCollection collList = site.Lists;

        FolderCollection fcolection1 = subFolder1.Folders;
        context.Load(fcolection1);
        //context.Load(collList);
        context.ExecuteQuery();

        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add("Rootfolder");
        dtTemp.Columns.Add("Subfolder");
        dtTemp.Columns.Add("subSubfolder");
        dtTemp.Columns.Add("Path");

        foreach (Folder subfolder in fcolection1)
        {

            folderlist.Add(subfolder.Name.ToString());
            if (!foldername.Contains("/Gresham Statements") && !foldername.Contains("Client Meetings"))
            {
                DataRow rowRootFolder = dtSharepint.NewRow();
                rowRootFolder["Subfolder"] = subfolder.Name.ToString();
                string newfoldername = foldername.Replace("Documents/Adams/", "");
                rowRootFolder["Path"] = newfoldername + "/" + subfolder.Name.ToString();
                dtSharepint.Rows.Add(rowRootFolder);

                subfolderlist = getSubFolderPath(site, context, foldername + "/" + subfolder.Name, dtSharepint, subfolder.Name.ToString());
            }
            else
            {
                DataRow rowRootFolder = dtTemp.NewRow();
                // rowRootFolder["subSubfolder"] = subfolder.Name.ToString();
                rowRootFolder["Subfolder"] = subfolder.Name.ToString();

                string newfoldername = foldername.Replace("Documents/Adams/", "");
                rowRootFolder["Path"] = newfoldername + "/" + subfolder.Name.ToString();
                dtTemp.Rows.Add(rowRootFolder);
            }
        }
        if (dtTemp.Rows.Count > 0)
        {
            dtTemp.DefaultView.Sort = "Subfolder ASC";


            DataView dv = dtTemp.DefaultView;
            dv.Sort = "Subfolder ASC";
            DataTable sortedDT = dv.ToTable();
            foreach (DataRow rw in sortedDT.Rows)
            {
                DataRow rowRootFolder = dtSharepint.NewRow();
                rowRootFolder["subSubfolder"] = rw["subSubfolder"];
                rowRootFolder["Path"] = rw["Path"];
                rowRootFolder["Subfolder"] = rw["Subfolder"];
                dtSharepint.Rows.Add(rowRootFolder);
            }
        }





        return folderlist;
    }
    public List<string> getSubFolderPath(Web site, ClientContext context, string foldername, DataTable dtSharepint, String folderName)
    {

        List<string> folderlist = new List<string>();

        Folder subFolder1 = site.GetFolderByServerRelativeUrl(foldername);

        FolderCollection fcolection1 = subFolder1.Folders;
        context.Load(fcolection1);
        //context.Load(collList);
        context.ExecuteQuery();
        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add("Subfolder");
        dtTemp.Columns.Add("subSubfolder");
        dtTemp.Columns.Add("Path");

        foreach (Folder subfolder in fcolection1)
        {
            folderlist.Add(subfolder.Name.ToString());

            //if (dtSharepint.Columns.Count != 4)
            //    dtSharepint.Columns.Add("subSubfolder");

            DataRow rowRootFolder = dtTemp.NewRow();
            rowRootFolder["subSubfolder"] = subfolder.Name.ToString();
            //   rowRootFolder["Subfolder"] = folderName;

            string newfoldername = foldername.Replace("Documents/Adams/", "");
            rowRootFolder["Path"] = newfoldername + "/" + subfolder.Name.ToString();
            dtTemp.Rows.Add(rowRootFolder);

        }



        if (dtTemp.Rows.Count > 0)
        {

            DataView dv = dtTemp.DefaultView;

            dtTemp.DefaultView.Sort = "subSubfolder ASC";
            dv.Sort = "subSubfolder ASC";
            DataTable sortedDT = dv.ToTable();
            foreach (DataRow rw in sortedDT.Rows)
            {
                DataRow rowRootFolder = dtSharepint.NewRow();
                rowRootFolder["subSubfolder"] = "    " + rw["subSubfolder"].ToString();
                rowRootFolder["Path"] = rw["Path"];
                //  rowRootFolder["Subfolder"] = rw["Subfolder"];
                dtSharepint.Rows.Add(rowRootFolder);
            }
        }


        return folderlist;
    }

    public void CopyFilenew(string FolderPath, string destFilename, string vSourcrFile)  // string vSourcefile, string vDestinationFile
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

    public void BindAllGrideView()
    {
        try
        {
           // string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            // foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);
            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);
            Web site = context.Web;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

            TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_4BLQTxDzt3+F9JB2YxRRiQ==");

            TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);
            TermGroup termGroup1 = termStore.Groups.GetByName("Client Portal");  //GUID = {94c3c53d-2351-3b5e-bfcb-c4f1b941157c}

            TermSet termsetClientName = termGroup1.TermSets.GetByName("Client Name");
            TermSet termSetDocumentType = termGroup1.TermSets.GetByName("Document Type");
            TermSet termSetYear = termGroup1.TermSets.GetByName("Year");

            TermCollection tcClientName = termsetClientName.GetAllTerms();
            TermCollection tcDocType = termSetDocumentType.GetAllTerms();
            TermCollection tcyear = termSetYear.GetAllTerms();

            context.Load(taxonomySession);
            context.Load(termStore);
            context.Load(termGroup);
            context.Load(termGroup1);

            context.Load(termsetClientName);
            context.Load(termSetDocumentType);
            context.Load(termSetYear);

            context.Load(tcClientName);
            context.Load(tcDocType);
            context.Load(tcyear);

            context.ExecuteQuery();


            DataTable dtClient = new DataTable();
            DataTable dtDocumentType = new DataTable();
            DataTable dtYear = new DataTable();

            dtClient.Columns.Add("clientName");
            dtDocumentType.Columns.Add("DocumentType");
            dtDocumentType.Columns.Add("Year");
            dtYear.Columns.Add("Year");
            foreach (Term ts in tcClientName)
            {
                DataRow row = dtClient.NewRow();
                row["clientName"] = ts.Name;
                dtClient.Rows.Add(row);

            }

            foreach (Term ts in tcDocType)
            {
                DataRow row = dtDocumentType.NewRow();
                row["DocumentType"] = ts.Name;
                dtDocumentType.Rows.Add(row);

            }

            foreach (Term ts in tcyear)
            {
                DataRow row = dtYear.NewRow();
                row["Year"] = ts.Name;
                dtYear.Rows.Add(row);

            }

            #region Bind Gridview
            foreach (DataRow row in dtClient.Rows)
            {
                DataRow row1 = dt.NewRow();
                row1["BatchName"] = row["clientName"].ToString();
                dt.Rows.Add(row1);
            }

            dt.DefaultView.Sort = "BatchName ASC";
            int i = 0;
            foreach (DataRow row in dtYear.Rows)
            {
                dtDocumentType.Rows[i]["Year"] = row["Year"].ToString();
                i++;
            }
            GridView1.DataSource = dtDocumentType;
            GridView1.DataBind();
            // dtDocumentType.DefaultView.Sort = "clientName ASC";

            #endregion
        }
        catch
        {

        }
    }


    #endregion

    protected void FileUpload1_Unload(object sender, EventArgs e)
    {
        string FilePath = Server.MapPath("~/") + @"ExcelTemplate\ClientPortalFileCopy\" + FileUpload1.FileName;
        if (System.IO.File.Exists(FilePath))
        {
            System.IO.File.Delete(FilePath);
        }
        //  string FilePath = @"D:\Test\ClientPortalFileCopy" + filename;
        FileUpload1.SaveAs(FilePath);
    }
    protected void txtAsOfDate_TextChanged1(object sender, EventArgs e)
    {
        checkboxbind();
        lblMessage.Text = "";
        lablmsg.Text = "";

    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        try
        {
            string FileName = FileUpload1.FileName;
            int ind = FileName.IndexOf(".");
            int len = FileName.Length;
            string extention = FileName.Substring(ind, len - ind);
            string newFilename = "TestCopyfile" + extention;


            string FilePath = Server.MapPath("~/") + @"ExcelTemplate\ClientPortalFileCopy\" + newFilename;
            //Response.Write("<br/>FilePath-" + FilePath);

            if (System.IO.File.Exists(FilePath))
            {
                System.IO.File.Delete(FilePath);
            }


            //Response.Write("<br/>FilePath-" + FilePath);
            //  string FilePath = @"D:\Test\ClientPortalFileCopy" + filename;
            FileUpload1.SaveAs(FilePath);

            // Response.Write("<br/>FilePath -" + FilePath);
            ViewState["Filename"] = FilePath;

            lblMessage.Text = "";
            lablmsg.Text = "";
        }
         catch(Exception ex )
        {
            lblMessage.Text = "Error :" + ex.Message.ToString();
        }
    }

    public void BindFolderPathddl()
    {
        DataTable dtFolderData = new DataTable();
        dtFolderData.Columns.Add("FolderPath");
        DataRow dr = dtFolderData.NewRow();
        dr["FolderPath"] = "Select";
        dtFolderData.Rows.Add(dr);
        dtFolderData.Merge(FolderData);

        //foreach(DataRow row in FolderData.Rows )
        //{
        //    DataRow drow = dtFolderData.NewRow();
        //    dr["FolderPath"] = drow["FolderPath"].ToString();
        //    dtFolderData.Rows.Add(dr);
        //}
        ddlPortalPath.DataTextField = "FolderPath";
        ddlPortalPath.DataSource = dtFolderData;
        ddlPortalPath.DataBind();

    }

    public void BindYearddl()
    {
        DataTable year = dsTaxonomyclientPortal.Tables[2];
        ddlYear.DataTextField = "Year";
        ddlYear.DataSource = year;
        ddlYear.DataBind();

        ddlYear.Items.Insert(0, "Select");
        ddlYear.Items[0].Value = "0";


    }

    public void BindclientGrideView()
    {
        DataSet dsData = (DataSet)ViewState["dsTaxonomyclientPortal"];
        DataTable dtClient = dsData.Tables[0];
        foreach (DataRow Clientrow in dtClient.Rows)
        {
            DataRow row = dt.NewRow();
            row["BatchName"] = Clientrow["clientName"];
            dt.Rows.Add(row);
        }
        //dt.Merge(dtClient);
        Bindgridview();

    }


    protected void ddlClientPortalPath_SelectedIndexChanged(object sender, EventArgs e)
    {
        string valuePath = ddlPortalPath.SelectedItem.Text;
        foreach (DataRow row in FolderData.Rows)
        {
            //dt.Columns.Add("FolderPath");
            //dt.Columns.Add("OnPortal");
            //dt.Columns.Add("Tag");
            if (valuePath == row["FolderPath"].ToString())
            {
                if (row["OnPortal"].ToString() == "Yes")
                    ddlYear.Visible = true;
            }
            else
            {
                ddlYear.Visible = false;
            }
        }
        lblMessage.Text = "";
        lablmsg.Text = "";
    }
    protected void ddlPortalPath_SelectedIndexChanged(object sender, EventArgs e)
    {
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
        lblMessage.Text = "";
        lablmsg.Text = "";

    }

    public void UploadFiles()
    {
        if (txtFilePath.Text != "")
        {
            try
            {


                string sourceFilename = txtfilename.Text;
                string Filenames = FileName.Text;
                string FolderName = txtFolderPath.Text;
                string folderPath = ddlPortalPath.SelectedItem.Text;

                DataTable dtFolderData = (DataTable)ViewState["dtFolderData"];

                //#region taxonomy declaration
                //string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
                //ClientContext context = new ClientContext(siteUrl);
                //SecureString passWord = new SecureString();
                //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
                //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
                //Web site = context.Web;
                //TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
                //TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_4BLQTxDzt3+F9JB2YxRRiQ==");

                //TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);
                //TermGroup termGroup1 = termStore.Groups.GetByName("Client Portal");  //GUID = {94c3c53d-2351-3b5e-bfcb-c4f1b941157c}
                //TermGroup tgClientName = termStore.Groups.GetByName("Client Name");
                //TermGroup tgYear = termStore.Groups.GetByName("Year");

                //TermSet termsetClientName = tgClientName.TermSets.GetByName("Client Name");
                //TermSet termSetDocumentType = termGroup1.TermSets.GetByName("Document Type");
                //TermSet termSetYear = tgYear.TermSets.GetByName("Year");

                //TermCollection tcClientName = termsetClientName.GetAllTerms();
                //TermCollection tcDocType = termSetDocumentType.GetAllTerms();
                //TermCollection tcyear = termSetYear.GetAllTerms();

                //context.Load(taxonomySession);
                //context.Load(termStore);
                //context.Load(termGroup);
                //context.Load(termGroup1);
                //context.Load(tgClientName);

                //context.Load(termsetClientName);
                //context.Load(termSetDocumentType);
                //context.Load(termSetYear);

                //context.Load(tcClientName);
                //context.Load(tcDocType);
                //context.Load(tcyear);

                //context.ExecuteQuery();
                //#endregion
                if (Filenames != "")
                {
                    if (folderPath != "Select")
                    {
                        string year = "";
                        if (ddlYear.Visible)
                        {
                            year = ddlYear.SelectedValue;
                        }

                        if (year != "0" )
                        {

                            string PathTaggingName = string.Empty;
                            string PathTaggingID = string.Empty;
                            string vIsYear = string.Empty;
                            string FolderDisplypath = folderPath;
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
                            //row["DocumentType"] = ts.Name;
                            //row["iID"] = ts.Id.ToString();
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
                                    FolderDisplypath = FolderDisplypath + "/" + Taggingyear;
                                }
                            }
                            string filenameWithoutext = Filenames.Substring(0, Filenames.LastIndexOf("."));
                            string vNewSharePointReportFolder = "Documents taxonomy";
                            //string vSourcrFile = @"E:\Log.txt";
                            string vSourcrFile = ViewState["Filename"].ToString();


                            if (vSourcrFile != null)
                            {
                                string exte = System.IO.Path.GetExtension(vSourcrFile);

                                DataTable dtClient = (DataTable)ViewState["dtSiteClientList"];

                                DataTable dtDisply = new DataTable();
                                dtDisply.Columns.Add("ClientName");
                                dtDisply.Columns.Add("FileUpload");
                                dtDisply.Columns.Add("FileUploadingStatus");
                                dtDisply.Columns.Add("Messeage");
                                dtDisply.Columns.Add("UploadFilepath");

                                int iClientCount = 0;
                                foreach (GridViewRow gvRow in gvList.Rows)
                                {
                                    iClientCount++;
                                    DataRow row = dtDisply.NewRow();
                                    row["UploadFilepath"] = FolderDisplypath;
                                    try
                                    {

                                        CheckBox ClientCheckbox = (CheckBox)gvRow.FindControl("chkbSelectBatch");
                                        bool CheckNameChecks = ClientCheckbox.Checked;
                                        string TaggingClientName = gvRow.Cells[2].Text;
                                        TaggingClientName = TaggingClientName.Replace("&#39;", "'");
                                        row["ClientName"] = TaggingClientName;

                                        string taggingClientID = string.Empty;
                                        if (ClientCheckbox.Checked)
                                        {
                                            row["FileUpload"] = "True";
                                            string iClientID = string.Empty;
                                            foreach (DataRow rw in dtClient.Rows)
                                            {
                                                if (TaggingClientName == rw["ClientName"].ToString())
                                                {
                                                    iClientID = rw["iID"].ToString();
                                                    break;
                                                }
                                            }

                                            if (cbTest.Checked)
                                                Filenames = filenameWithoutext + "_" + iClientID + "_Test" + exte;
                                            else
                                                Filenames = filenameWithoutext + "_" + iClientID + exte;

                                            foreach (DataRow rw in dtClientSite.Rows)
                                            {
                                                if (TaggingClientName == rw["clientName"].ToString())
                                                {
                                                    taggingClientID = rw["iID"].ToString();
                                                    break;
                                                }

                                            }

                                         //   string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";

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

                                            List docs = context.Web.Lists.GetByTitle("Documents taxonomy");

                                            byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
                                            System.IO.Stream stream = new System.IO.MemoryStream(bytes);

                                            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(vNewSharePointReportFolder);
                                            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(Filenames), Overwrite = true };
                                            Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile);

                                            context.Load(uploadFile);
                                            context.Load(docs);
                                            context.ExecuteQuery();

                                            context.Load(uploadFile.ListItemAllFields);
                                            context.ExecuteQuery();

                                           // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

                                            Microsoft.SharePoint.Client.ListItem item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
                                            context.Load(item);
                                            context.ExecuteQuery();

//                                            context.Load(docs.Fields.GetByTitle("p9cafa43d635492cb87a8a60d0ebb191"));

//                                            context.Load(docs.Fields.GetByTitle("p9cafa43d635492cb87a8a60d0ebb191"));

//                                            //var name = docs.Fields.GetByTitle("p9cafa43d635492cb87a8a60d0ebb191").InternalName;

//                                            uploadFile.ListItemAllFields["p9cafa43d635492cb87a8a60d0ebb191"] = "2012";

//                                            uploadFile.ListItemAllFields.Update();
//                                            context.Load(uploadFile);
//                                            context.ExecuteQuery();


//                                            //using (var fs = new FileStream(vSourcrFile, FileMode.Open))
//                                            //{
//                                            //    var fi = new FileInfo(vSourcrFile);
//                                            //    var list = context.Web.Lists.GetByTitle("Documents");
//                                            //    context.Load(list.RootFolder);
//                                            //    context.ExecuteQuery();

//                                            //    var fileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, Path.GetFileName(Filenames));

//                                            //    Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, fs, true);
//                                            //}

//                                            //CamlQuery _CamlQuery = new CamlQuery();
//                                            //_CamlQuery.ViewXml = "<View><Query><Where><Contains><FieldRef Name='FileLeafRef'/><Value Type='Text'>" + Filenames + "</Value>" +
//                                            //                     "</Contains></Where></Query><RowLimit>100</RowLimit></View>";
//                                            CamlQuery _CamlQuery = new CamlQuery();
//                                            _CamlQuery.ViewXml =
//                                            @"<View Scope='RecursiveAll'>
//                   <Query>
//                   <OrderBy UseIndexForOrderBy = 'TRUE'> <FieldRef Name='Modified' Ascending='False' /></OrderBy>
//                      </Query>
//                      <RowLimit>10</RowLimit>
//                      </View>";

//                                            Microsoft.SharePoint.Client.ListItemCollection _ListItemCollection = docs.GetItems(_CamlQuery);
//                                            context.Load(_ListItemCollection);
//                                            context.ExecuteQuery();



//                                            Microsoft.SharePoint.Client.ListItem item = null;
//                                            foreach (Microsoft.SharePoint.Client.ListItem listItem in _ListItemCollection)
//                                            {
//                                                if (Filenames == listItem["FileLeafRef"].ToString())
//                                                {
//                                                    item = listItem;
//                                                    break;
//                                                }
//                                            }

//                                            context.Load(item);
//                                            context.ExecuteQuery();

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
                                                    item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;
                                                }
                                                else
                                                {
                                                    taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                                    taxonomyFieldValueYear.Label = Taggingyear;
                                                    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
                                                }
                                            }
                                            else
                                            {
                                                taxonomyFieldValueYear.TermGuid = TaggingYearID;
                                                taxonomyFieldValueYear.Label = Taggingyear;
                                                item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
                                            }

                                            item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;
                                            item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;


                                            item.Update();
                                            docs.Update();
                                            context.ExecuteQuery();
                                            row["FileUploadingStatus"] = "Success";

                                        }
                                        else
                                        {
                                            row["FileUpload"] = "False";

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        row["FileUploadingStatus"] = "Fail";
                                        row["Messeage"] = ex.Message;

                                    }

                                    dtDisply.Rows.Add(row);

                                }
                                Session["dtDisply"] = dtDisply;
                                lblMessage.Text = "File Copied Successfully";


                                //added -sasmit 11/9/20116
                                #region sesson to send to frmClientCopyResult
                                int count = 0;
                                foreach (DataRow row in dtDisply.Rows)
                                {
                                    if (row["FileUpload"].ToString() == "True")
                                        count++;

                                }
                                if (cbTest.Checked == true)
                                {
                                    string mode = "Yes";
                                    Session["Mode"] = mode;

                                }
                                else
                                {
                                    string mode = "No";
                                    Session["Mode"] = mode;
                                }



                                Session["FileName"] = FileName.Text;
                                Session["StartDate"] = txtStartDate.Text;
                                Session["AsOfDate"] = txtAsOfDate.Text;
                                Session["Fund"] = ddlFund.SelectedItem.Text;
                                #endregion
                                //added-sasmit 11/9/2016

                                if (dtDisply.Rows.Count > 0)
                                {
                                    # region Send Email
                                    try
                                    {
                                        EmailBody();
                                        SendEmail(emailBody.ToString(), "UPLOAD COMPLETED", "skane@infograte.com", "");

                                    }
                                    catch (Exception ex)
                                    {
                                        lblMessage.Text = "Error sending mail.";
                                        Response.Write("Error sending mail.-"+ex.ToString());
                                    }
                                    #endregion
                                    string url = "frmClientCopyResult.aspx";
                                    //StringBuilder sb = new StringBuilder();
                                    //sb.Append("<script type = 'text/javascript'>");
                                    //sb.Append("window.open('");
                                    //sb.Append(url);
                                    //sb.Append("');");
                                    //sb.Append("</script>");
                                    //ClientScript.RegisterStartupScript(this.GetType(),
                                    //        "script", sb.ToString());
                                    Response.Write("<script>");
                                    //  string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
                                    Response.Write("window.open('" + url + "', 'mywindow')");
                                    //  Response.Write("window.open('ViewReport.aspx?" + fsFinalLocation + "', 'mywindow')");

                                    Response.Write("</script>");
                                }

                                txtfilename.Text = "";
                                txtFolderPath.Text = "";
                                txtFilePath.Text = "";
                                //lblMessage.Text = "";
                                FileName.Text = "";
                                BindFund();
                                txtAsOfDate.Text = "";

                            }
                            else
                            {
                                lblMessage.Text = "Please Select The File..";
                            }

                        }
                        else
                        {
                            lblMessage.Text = "Please Select Year";
                        }

                    }
                    else
                    {
                        lblMessage.Text = "Please Select File Path";
                    }
                }
                else
                {
                    lblMessage.Text = "Please Enter File Name";
                }
            }
            catch
            { }
        }
        else
        {
            lblMessage.Text = "Please Select File Name";

        }

    }

    //added sasmit 11/9/2016
    #region Email Section
    public void EmailBody()
    {
        DataTable dtDisply = (DataTable)Session["dtDisply"];
        if (dtDisply.Rows.Count > 0)
        {
            int count = 0;
            foreach (DataRow row in dtDisply.Rows)
            {
                if (row["FileUpload"].ToString() == "True")
                    count++;

            }

            string fund ="";
            if (ddlFund.SelectedItem.Text != "Select Fund")
            {
                fund = ddlFund.SelectedItem.Text;
            }

            //int count = dtDisply.Rows.Count;
            string Path = dtDisply.Rows[0]["UploadFilepath"].ToString();

            addToEmailBody("<b>Client Portal Upload Completed </b><br><br><br>");
            addToEmailBody("<b>Count Of Households Successfully Uploaded: </b>" + count + "<br><br>");
            addToEmailBody("<b>Upload to Folder: </b>" + Path + "<br>");
            addToEmailBody("<b>FileName: </b>" + FileName.Text + "<br>");

            addToEmailBody("<b>Start Date: </b>" + txtStartDate.Text + "<br>");
            addToEmailBody("<b>As Of Date: </b>" + txtAsOfDate.Text + "<br>");
            addToEmailBody("<b>Fund: </b>" + fund + "<br>");
            if (cbTest.Checked == true)
            {
                string mode = "Yes";
                addToEmailBody("<b>Test Mode: </b>" + mode + "<br>");

            }
            else
            {
                string mode = "No";
                addToEmailBody("<b>Test Mode: </b>" + mode + "<br>");
            }





            addToEmailBody("<table border=\"1\" cellspacing=\"0\">");
            addToEmailBody("<tr><td><b>Client</b></td><td><b>Upload File?</b></td><td><b>Status</b></td><td><b>Message</b></td></tr>\r\n");
            foreach (DataRow r in dtDisply.Rows)
            {
                string ClientName = r["ClientName"].ToString();
                string FileUpload = r["FileUpload"].ToString();
                string Status = r["FileUploadingStatus"].ToString();
                string Message = r["Messeage"].ToString();

                addToEmailBody("<tr><td>" + ClientName + "</td><td>" + FileUpload + "</td><td>" + Status + "</td><td>" + Message + "</td></tr>\r\n");


            }
            addToEmailBody("</table>\r\n");


        }
    }

    //added sasmit 11/9/2016
    private void addToEmailBody(string str)
    {
        emailBody.Append(str);
    }

    //added sasmit 11/9/2016
    public void SendEmail(string mailmessage, string subject, string mailTo, string Attachment1)
    {
        
        try
        {
            MailMessage myMessage = new MailMessage();
            // SmtpClient SMTPSERVER = new SmtpClient();

        string EmailID = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
        string Password = AppLogic.GetParam(AppLogic.ConfigParam.Password);
        string SMTPHost = AppLogic.GetParam(AppLogic.ConfigParam.SMTPHost);
        string ToEmailIDs1 = AppLogic.GetParam(AppLogic.ConfigParam.ToEmailIDs1);
        int Port = Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.Port));

          //  Response.Write("EmailID" + EmailID);
          //  Response.Write("Password" + Password);
          // /  Response.Write("SMTPHost" + SMTPHost);
          ////  Response.Write("ToEmailIDs1" + ToEmailIDs1);
          //   Response.Write("Port" + Port);

            //string EmailID = "CRMAdmin@greshampartners.com";
            //string Password = "W!gmxF26ggw]";
            //string SMTPHost = "Smtp.office365.com";
            //int Port = 25;
            //string ToEmailIDs1 = "skane@infograte.com";

            //int Port = = ConfigurationSettings.AppSettings["EmailId"].ToString();
            //string Password = ConfigurationSettings.AppSettings["Password"].ToString();
            //string SMTPHost = ConfigurationSettings.AppSettings["SMTPHost"].ToString();
            //string ToEmailIDs = ConfigurationSettings.AppSettings["ToEmailIDs"].ToString();
            //int Port = Convert.ToInt32(ConfigurationSettings.AppSettings["Port"]);

            myMessage.From = new MailAddress(EmailID, "Gresham Client Portal FileCopy");
            string[] strTo = ToEmailIDs1.Split('|');


            for (int i = 0; i < strTo.Length; i++)
            {
                if (strTo[i] != "")
                {
                    myMessage.To.Add(new MailAddress(strTo[i]));
                }
            }
            //myMessage.CC.Add("GBhagia@infograte.com");
            //myMessage.Bcc.Add("dshah@webdevinc.net");
            myMessage.Bcc.Add("skane@infograte.com");
            myMessage.Bcc.Add(new MailAddress("auto-emails@infograte.com"));
            // myMessage.Bcc.Add("svaitya@webdevinc.net");
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


            //SMTPSERVER.EnableSsl = false; for office 365 mailing
            SMTPSERVER.EnableSsl = true;
            // smtp.EnableSsl = true;
            SMTPSERVER.UseDefaultCredentials = true;
            System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(EmailID, Password);
            SMTPSERVER.Credentials = basicAuthenticationInfo;
            SMTPSERVER.Send(myMessage);

            myMessage.Dispose();
            myMessage = null;
            SMTPSERVER = null;
        }
        catch (Exception ex)
        {
           

            string strDescription = "Error sending Mail :" + ex.ToString();
            Response.Write(strDescription);

        }
    }
    #endregion

    protected void ddlHouseHoldType_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lablmsg.Text = "";
    }
}