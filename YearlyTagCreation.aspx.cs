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


public partial class YearlyTagCreation : System.Web.UI.Page
{

    int year1 = 0;
    int year2 = 0;
    sharepoint sh = new sharepoint();
    bool result = false;
    public StreamWriter sw = null;
    DataTable dtCLientList = new DataTable();

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            CheckYearTag();
            ClearLabel();
            dtCLientList = FolderList();
            ViewState["dtCLientList"] = dtCLientList;
            Bindddl(dtCLientList);

        }
    }
    public void CheckYearTag()
    {
        try
        {
            //lblMsg.Text = "Processing....";
            string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            // foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);
            string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);
            Web site = context.Web;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
           // TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_4BLQTxDzt3+F9JB2YxRRiQ=="); // commented 1_10_2019

 Guid SharepointTermStoreID =  new Guid("8f0311806e7c4e72aa9d55d7cf0d8400");
            TermStore termStore = taxonomySession.TermStores.GetById(SharepointTermStoreID);

            TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);
            TermGroup tgYear = termStore.Groups.GetByName("Year");
            TermSet termsetYear = tgYear.TermSets.GetByName("Year");
            TermCollection tcYear = termsetYear.GetAllTerms();

            int cnt = 0;
            context.Load(taxonomySession);
            context.Load(tgYear);
            context.Load(termsetYear);
            context.Load(tcYear);
            context.ExecuteQuery();
            int count = tcYear.Count();



            foreach (Term ts in tcYear)
            {
                cnt++;

                if (count == cnt)
                {
                    int Year = Convert.ToInt32(ts.Name);


                    if (DateTime.Now.Year != Year)
                    {
                        year1 = Year + 1;
                        Session["Year1"] = year1;
                    }
                    else
                    {
                        year1 = Year + 1;
                        Session["Year1"] = year1;
                        year2 = Year + 2;
                        Session["Year2"] = year2;
                    }
                    Bindddl(ddlYear);
                    break;
                }

            }



        }
        catch
        { }

    }

    public void Bindddl(DataTable ddl)
    {

        DataTable dt = new DataTable();
        dt.Columns.Add("Tag");
        foreach (DataRow dr in ddl.Rows)
        {
            dt.Rows.Add(dr["Tag"].ToString());
        }

        dt.DefaultView.Sort = "Tag ASC";
        ddlFolderName.DataSource = dt;
        ddlFolderName.DataTextField = "Tag";

        ddlFolderName.DataValueField = "Tag";
        ddlFolderName.DataBind();

        ddlFolderName.Items.Insert(0, "Select");
        ddlFolderName.Items[0].Value = "0";
    }


    public bool CreateYearTag()
    {
        try
        {
            //lblMsg.Text = "Processing....";
            string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            // foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);
            string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);
            TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_4BLQTxDzt3+F9JB2YxRRiQ==");
            TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);
            TermGroup tgYear = termStore.Groups.GetByName("Year");
            TermSet termsetYear = tgYear.TermSets.GetByName("Year");
            TermCollection tcYear = termsetYear.GetAllTerms();

            int cnt = 0;
            context.Load(taxonomySession);
            context.Load(tgYear);
            context.Load(termsetYear);
            context.Load(tcYear);
            context.ExecuteQuery();
            int count = tcYear.Count();
            // bool result = false;

            string year = ddlYear.SelectedItem.Text;

            foreach (Term ts in tcYear)
            {
                cnt++;
                if (count == cnt)
                {
                    string Year = ts.Name;

                    if (Year != ddlYear.SelectedItem.Text)
                    {
                        result = true;
                        break;
                    }

                }

            }

            if (result == true)
            {
                Guid newTermId = Guid.NewGuid();
                Term newTerm = termsetYear.CreateTerm(year, 1033, newTermId);
                context.Load(newTerm);
                context.Load(termsetYear);
                context.ExecuteQuery();
            }



        }


        catch
        { }
        return result;

    }

    public void Bindddl(DropDownList ddl)
    {
        ddl.Items.Clear();
        ddl.Items.Insert(0, "Select");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;

        if (Convert.ToInt32(Session["Year1"]) != 0 && Convert.ToInt32(Session["Year2"]) != 0)
        {
            ddl.Items.Insert(1, Convert.ToString(year1));
            ddl.Items[1].Value = "1";

            ddl.Items.Insert(2, Convert.ToString(year2));
            ddl.Items[2].Value = "2";
        }
        else
        {
            ddl.Items.Insert(1, Convert.ToString(year1));
            ddl.Items[1].Value = "1";
        }
        //ddl.Items[1].Value = "1";
        //ddl.SelectedIndex = 1;
        Session.Remove("Year1");
        Session.Remove("Year2");
    }

    protected void Button1_Click(object sender, EventArgs e)
    {

        DataTable dtCLientList1 = (DataTable)ViewState["dtCLientList"];

        ClearLabel();
        

        #region LogFile
        ///* Creating Log File */
        //string LogFileName = string.Empty;
        //LogFileName = "Log-" + DateTime.Now;
        ////string FolderName = "Logs";
        //LogFileName = LogFileName.Replace(":", "-");
        //LogFileName = LogFileName.Replace("/", "-");
        ////sw = new StreamWriter(Server.MapPath("\\" + FolderName + "\\" + LogFileName)  + ".txt", true);
        //// sw = new StreamWriter(Server.MapPath(@"\Logs" + LogFileName)+".txt", true);
        ////sw = new StreamWriter(Server.MapPath("/Logs"+"/"+LogFileName) + DateTime.Now);
        //sw = new StreamWriter(Server.MapPath("") + @"\Logs" + "/" + LogFileName +".txt",true);


        #endregion

        if (ddlYear.SelectedValue != "0" && ddlYear.Enabled ==true)
        {

            foryear(dtCLientList1);
        }

        else if (ddlYear.Enabled == true)
        {
            lblMsg.Text = "Please Select Year";
        }

        else if (ddlFolderName.SelectedValue != "0" && ddlFolderName.Enabled == true)
        {
            forFolder(dtCLientList1);
        }
        else if (ddlFolderName.Enabled == true)
        {
            lblMsg.Text = "Please Select Folder";
        }
      
    }


    public void foryear(DataTable dtCLientList1)
    {
        
        int cnt = 0;
        CreateYearTag();
        string YearTag = ddlYear.SelectedItem.Text;

        //string SourcePath = @"\\grpao1-vwfs01\shared$\infograte\Summitas Test\";
        string SourcePath = AppLogic.GetParam(AppLogic.ConfigParam.SummitasPath).ToString();

        string user = AppLogic.GetParam(AppLogic.ConfigParam.UserName).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();

        using (new Impersonation("corp", user, Pass))
        {
            DirectoryInfo source = new DirectoryInfo(SourcePath);

            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {

                string ClientName = diSourceSubDir.Name;

             
                DirectoryInfo sourceClient = new DirectoryInfo(SourcePath + "\\" + ClientName);

                foreach (DirectoryInfo diSourceSuBDir in sourceClient.GetDirectories())
                {
                    string SubFolderName = diSourceSuBDir.Name;
                   

                    foreach (DataRow dr in dtCLientList1.Rows)
                    {

                        string year = dr["Year"].ToString();
                        string FolderPath = dr["FolderPath"].ToString();

                        if (year.ToLower() == "true")
                        {
                            cnt++;
                            bool check = Directory.Exists(SourcePath + "\\" + ClientName + "\\" + SubFolderName + "\\" + FolderPath);


                            if (check == true)
                            {
                                string YearFolder = SourcePath + "\\" + ClientName + "\\" + SubFolderName + "\\" + FolderPath + "\\" + YearTag;
                                if (!Directory.Exists(YearFolder))
                                {
                                    Directory.CreateDirectory(YearFolder);

                                }

                            }
                        }

                    }
                }


            }
        }
        lblMsg.Text = "";
        lblMsg0.Text = ddlYear.SelectedItem.Text + "" + "Created!";
        lblMsg1.Text = ddlYear.SelectedItem.Text + " " + " added to Summitas Folders";
        ddlYear.Enabled = true;
        ddlFolderName.Enabled = true;
    }
    protected void ddlYear_SelectedIndexChanged(object sender, EventArgs e)
    {
        //Bindddl(ddlYear);
        ClearLabel();
        ddlFolderName.Enabled = false;
    }


    public DataTable FolderList()
    {

        DataTable dtClient = new DataTable();
        string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";


        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        Web site = context.Web;

        List list = site.Lists.GetByTitle("CP Mapping");
        CamlQuery caml = new CamlQuery();
        Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(caml);
        context.Load(list);
        context.Load(items);
        context.ExecuteQuery();

        //  DataTable dt = new DataTable();
        dtClient.Columns.Add("FolderPath");
        dtClient.Columns.Add("Year");
        dtClient.Columns.Add("Tag");

        foreach (Microsoft.SharePoint.Client.ListItem item in items)
        {
            context.Load(item);
            context.ExecuteQuery();
            string folderpath = null;
            string Year = null;
            string Tag = null;
            try
            {
                Year = item["On_x0020_Portal"].ToString();
            }
            catch
            {

            }
            folderpath = item["Title"].ToString();
            Tag = item["_x0070_bi4"].ToString();

            DataRow dr = null;
            dr = dtClient.NewRow();
            dr["FolderPath"] = folderpath;
            dr["Year"] = Year;
            dr["Tag"] = Tag;

            dtClient.Rows.Add(dr);


        }
        return dtClient;
    }

    public void ClearLabel()
    {
        lblMsg.Text = "";
        lblMsg0.Text = "";
        lblMsg1.Text = "";
    }
    protected void ddlFolderName_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearLabel();
        ddlYear.Enabled = false;
    }


    public void forFolder(DataTable dtCLientList1)
    {
        
        int cnt = 0;
       // string SourcePath = @"C:\Users\svaity\Desktop\12_26_2017\Summitastest\";
        string SourcePath = AppLogic.GetParam(AppLogic.ConfigParam.SummitasPath).ToString();

        string user = AppLogic.GetParam(AppLogic.ConfigParam.UserName).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        //using (new Impersonation("corp", "gbhagia", "51ngl3malt"))
        using (new Impersonation("corp", user, Pass))
        {
            DirectoryInfo source = new DirectoryInfo(SourcePath);

            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {

                string ClientName = diSourceSubDir.Name;

                
                DirectoryInfo sourceClient = new DirectoryInfo(SourcePath + "\\" + ClientName);

                foreach (DirectoryInfo diSourceSuBDir in sourceClient.GetDirectories())
                {
                    string SubFolderName = diSourceSuBDir.Name;
                    

                    foreach (DataRow dr in dtCLientList1.Rows)
                    {

                        string year = dr["Year"].ToString();
                        string FolderPath = dr["FolderPath"].ToString().Replace("/", "\\");



                        if (ddlFolderName.SelectedItem.Text.ToLower() == dr["Tag"].ToString().ToLower())
                        {
                            //if (year.ToLower() == "true")
                            //{

                            //    int currentyear = DateTime.Now.Year;
                            //    string Comm = SourcePath + "\\" + ClientName + "\\" + SubFolderName + "\\" + FolderPath;
                            //    bool checkFolder = Directory.Exists(Comm);

                            //    if (checkFolder == true)
                            //    {

                            //        result = createyearfolder(Comm, currentyear);
                            //        if (result == true)
                            //            cnt++;
                            //    }

                            //    else
                            //    {
                            //        Directory.CreateDirectory(Comm);
                            //        result = createyearfolder(Comm, currentyear);
                            //        if (result == true)
                            //            cnt++;
                            //    }

                            //}
                            //else
                            //{
                                string Comm = SourcePath + "\\" + ClientName + "\\" + SubFolderName + "\\" + FolderPath;
                                result = !Directory.Exists(Comm);
                                if (result == true)
                                {
                                    Directory.CreateDirectory(Comm);
                                    cnt++;
                                }

                            //}
                            break;
                        }

                    }
                }

            }
        }

        if (cnt > 0)
        {
            lblMsg.Text = "";
            //lblMsg0.Text = ddlFolderName.SelectedItem.Text + "" + " New Tag Created!";
            lblMsg0.Text =  ddlFolderName.SelectedItem.Text +" Created!";
            lblMsg1.Text =  ddlFolderName.SelectedItem.Text + " " + " added to No. Of Clients" + cnt;
        }

        else
        {
            lblMsg.Text = "Folder Already Exists!!";
        }
        ddlYear.Enabled = true;
        ddlFolderName.Enabled = true;
    }

    public bool createyearfolder(string Comm, int currentyear)
    {
        bool flagYear = false;
        try
        {
            bool checkcurrentyear = Directory.Exists(Comm + "\\" + currentyear);

            if (!checkcurrentyear)
            {
                Directory.CreateDirectory(Comm + "\\" + currentyear);
                flagYear = true;
            }
            currentyear = currentyear + 1;
            bool checknextyear = Directory.Exists(Comm + "\\" + currentyear);
            if (!checknextyear)
            {

                Directory.CreateDirectory(Comm + "\\" + currentyear);
                flagYear = true;
            }
            return flagYear;
        }
        catch (Exception ex)
        {
            flagYear = false;
            return flagYear;
        }

    }
}