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

public partial class frmCommBilling : System.Web.UI.Page
{

    DataTable dtCLientList = new DataTable();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
             dtCLientList = FolderList();
             ViewState["dtCLientList"] = dtCLientList;
            Bindddl(dtCLientList);
        }
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

    protected void ddlYear_SelectedIndexChanged(object sender, EventArgs e)
    {
        //Bindddl(ddlFolderName);
        // ClearLabel();
    }


    protected void Button1_Click(object sender, EventArgs e)
    {
        DataTable dtCLientList1 = (DataTable)ViewState["dtCLientList"];
        bool result = false;

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

        if (ddlFolderName.SelectedValue != "0")
        {
            //bool YearTaxonomy = true;
            int cnt = 0;

            // bool YearTaxonomy = CreateYearTag();

            //if (YearTaxonomy == true)
            //{
            // FolderList();

            string YearTag = ddlFolderName.SelectedItem.Text;

            //DirectoryInfo source = new DirectoryInfo(@"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas\");

            //string SourcePath = @"\\grpao1-vwfs01\shared$\Client Portal\STest\";

            // string SourcePath = @"\\grpao1-vwfs01\Shared$\Client Portal\Summitas\";

            string SourcePath = @"\\grpao1-vwfs01\shared$\infograte\Summitas Test1\";

            //  string SourcePath =  @"\\grpao1-vwfs01\shared$\infograte\Summitas Test\";

            //  string SourcePath = @"\\grpao1-vwfs01\shared$\infograte\S Test";
            string user = AppLogic.GetParam(AppLogic.ConfigParam.UserName).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
           // using (new Impersonation("corp", "gbhagia", "51ngl3malt"))
            
            using (new Impersonation("corp", user, Pass))
            {
                DirectoryInfo source = new DirectoryInfo(SourcePath);

                foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
                {

                    string ClientName = diSourceSubDir.Name;

                    //string[] filesindirectory = Directory.GetDirectories(Server.MapPath("~/Test" + "/" + ClientName));
                    DirectoryInfo sourceClient = new DirectoryInfo(SourcePath + "\\" + ClientName);

                    foreach (DirectoryInfo diSourceSuBDir in sourceClient.GetDirectories())
                    {
                        string SubFolderName = diSourceSuBDir.Name;
                        //DirectoryInfo sourceClient = new DirectoryInfo(@"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName);

                        foreach (DataRow dr in dtCLientList1.Rows)
                        {

                            string year = dr["Year"].ToString();
                            string FolderPath = dr["FolderPath"].ToString().Replace("/","\\");



                            if (ddlFolderName.SelectedItem.Text.ToLower() == dr["Tag"].ToString().ToLower())
                            {
                                if (year.ToLower() == "true")
                                {

                                    int currentyear = DateTime.Now.Year;

                                    //if (ddlFolderName.SelectedItem.Text.ToLower() == "communication")
                                    //{
                                    string Comm = SourcePath + "\\" + ClientName + "\\" + SubFolderName + "\\" + FolderPath;
                                    bool checkFolder = Directory.Exists(Comm);

                                    if (checkFolder == true)
                                    {

                                        result = createyearfolder(Comm, currentyear);
                                        cnt++;
                                    }

                                    else
                                    {
                                        Directory.CreateDirectory(Comm);
                                        result = createyearfolder(Comm, currentyear);
                                        cnt++;
                                    }

                                }
                                else
                                {
                                    string Comm = SourcePath + "\\" + ClientName + "\\" + SubFolderName + "\\" + FolderPath;
                                     result = !Directory.Exists(Comm);
                                    if (result == true)
                                    {
                                        Directory.CreateDirectory(Comm);
                                        cnt++;
                                    }

                                }
                                break;
                            }





                        }
                    }


                    #region Not in Use
                    //DirectoryInfo sourceClient = new DirectoryInfo(@"D:\PRACTICE\Test" + "\\" + ClientName);


                    //foreach (DirectoryInfo diSourceSuBDir in sourceClient.GetDirectories())
                    //{
                    //    string SubFolderName = diSourceSuBDir.Name;

                    //    //DirectoryInfo SubFolder = new DirectoryInfo(@"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName);

                    //    DirectoryInfo SubFolder = new DirectoryInfo(@"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName);

                    //    //foreach (DirectoryInfo diSubFolder in SubFolder.GetDirectories())
                    //    //{
                    //        //string Folder = diSubFolder.Name;
                    //    /*********************************CrmTest3 Path ********************** */
                    //    //string pathCM = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Client Meetings";
                    //    //string pathITax_K1 = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Income Tax" + "\\" + "K-1s";
                    //    //string pathITax_tx = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Income Tax" + "\\" + "Tax Returns";
                    //    //string pathInv_NonM = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Investment Activity" + "\\" + "NonMarketable";
                    //    //string pathInv_M = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Investment Activity" + "\\" + "Marketable";
                    //    //string pathThirdP_AF = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Third Party Reports" + "\\" + "Audited Financials";
                    //    //string pathThirdP_ls = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Third Party Reports" + "\\" + "Investment Statements";
                    //    //string pathGS = @"\\GRPAO1-VWFS01\Shared$\Client Portal\Summitas" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Gresham Statements";
                    //    /*********************************CrmTest3 Path ********************** */

                    //        /********Local Path *****************************************************/

                    //    string pathCM = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Client Meetings";
                    //    string pathITax_K1 = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Income Tax" + "\\" + "K-1s";
                    //    string pathITax_tx = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Income Tax" + "\\" + "Tax Returns";
                    //    string pathInv_NonM = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Investment Activity" + "\\" + "NonMarketable";
                    //    string pathInv_M = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Investment Activity" + "\\" + "Marketable";
                    //    string pathThirdP_AF = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Third Party Reports" + "\\" + "Audited Financials";
                    //    string pathThirdP_ls = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Third Party Reports" + "\\" + "Investment Statements";
                    //    string pathGS = @"D:\PRACTICE\Test" + "\\" + ClientName + "\\" + SubFolderName + "\\" + "Gresham Statements";

                    //        /********Local Path *****************************************************/


                    //        bool check = Directory.Exists(pathCM);

                    //        if (Directory.Exists(pathCM))
                    //        {
                    //            string TPath = pathCM +"\\" +YearTag;
                    //            if (!Directory.Exists(TPath))
                    //            Directory.CreateDirectory(TPath);
                    //        }
                    //        if (Directory.Exists(pathITax_K1))
                    //        {
                    //            string TPath = pathITax_K1 + "\\" + YearTag;
                    //            if (!Directory.Exists(TPath))
                    //                Directory.CreateDirectory(TPath);
                    //        }
                    //        if (Directory.Exists(pathITax_tx))
                    //        {
                    //            string TPath = pathITax_tx + "\\" + YearTag;
                    //            if (!Directory.Exists(TPath))
                    //                Directory.CreateDirectory(TPath);
                    //        }
                    //        if (Directory.Exists(pathInv_NonM))
                    //        {
                    //            string TPath = pathInv_NonM + "\\" + YearTag;
                    //            if (!Directory.Exists(TPath))
                    //                Directory.CreateDirectory(TPath);
                    //        }
                    //        if (Directory.Exists(pathInv_M))
                    //        {
                    //            string TPath = pathInv_M + "\\" + YearTag;
                    //            if (!Directory.Exists(TPath))
                    //                Directory.CreateDirectory(TPath);
                    //        }
                    //        if (Directory.Exists(pathThirdP_AF))
                    //        {
                    //            string TPath = pathThirdP_AF + "\\" + YearTag;
                    //            if (!Directory.Exists(TPath))
                    //                Directory.CreateDirectory(TPath);
                    //        }
                    //        if (Directory.Exists(pathThirdP_ls))
                    //        {
                    //            string TPath = pathThirdP_ls + "\\" + YearTag;
                    //            if (!Directory.Exists(TPath))
                    //                Directory.CreateDirectory(TPath);
                    //        }
                    //        if (Directory.Exists(pathGS))
                    //        {
                    //            string TPath = pathGS + "\\" + YearTag;
                    //            if (!Directory.Exists(TPath))
                    //                Directory.CreateDirectory(TPath);

                    //        }
                    //    //}
                    //}
                    ////DirectoryInfo nextTargetSubDir =
                    ////    target.CreateSubdirectory(diSourceSubDir.Name);
                    ////CopyAll(diSourceSubDir, nextTargetSubDir);
                    #endregion

                }
            }

            if (result == true  )
            {
                lblMsg.Text = "";
                lblMsg0.Text = ddlFolderName.SelectedItem.Text + "" + " New Tag Created!";
                lblMsg1.Text = "New Tag" + " " + ddlFolderName.SelectedItem.Text + " " + " added to Summitas Folders" + cnt;
            }

            else
            {
                lblMsg.Text = "Folder Already Found!!";
            }

        }
        //    else if(YearTaxonomy==false)
        //    {
        //        lblMsg.Text = lblMsg1.Text + ddlFolderName.SelectedItem.Text + "" + " Tag Already Exist" + YearTaxonomy.ToString();

        //    }
        //}
        else
        {
            lblMsg.Text = "Please Select Year";
        }
    }

    public void ClearLabel()
    {
        lblMsg.Text = "";
        lblMsg0.Text = "";
        lblMsg1.Text = "";
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
    protected void ddlFolderName_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearLabel();
    }
}