
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class HouseholdDashboard : System.Web.UI.Page
{
    bool bRoles = false;
    string CurrentUser = string.Empty;
    string Check = string.Empty;
    bool isEmpty = false;
    // DataTable dtDocumentTaxonomy;//Commented 7_24_2019 - Page SLow Issue
    DataTable dtActiveClientList;
    sharepoint spcls = new sharepoint();
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            CurrentUser = GetcurrentUser("");
            CurrentUser = CurrentUser == "0" ? "null" : "'" + CurrentUser + "'";
            if (!IsPostBack)
            {
                BindddlRoles(CurrentUser);
                BindddlAdvisor();
                //Commented 7_24_2019 - Page SLow Issue
                //if (ViewState["dtDocumentTaxonomy"] == null && ViewState["dtActiveClientList"] == null)
                //{
                //    try
                //    {
                //        dtDocumentTaxonomy = spcls.getTaxonomyClientService();
                //        dtActiveClientList = spcls.getActiveClientList();
                //        ViewState["dtDocumentTaxonomy"] = dtDocumentTaxonomy;
                //        ViewState["dtActiveClientList"] = dtActiveClientList;
                //    }
                //    catch (Exception Ex)
                //    {
                //        lblError.Visible = true;
                //        Label1.Text = "Error Fetching Taxonomy List";
                //    }
                //}

                if (Application["dtActiveClientList"] == null)
                {
                    try
                    {
                        Application["dtActiveClientList"] = spcls.getActiveClientList();
                    }
                    catch (Exception Ex)
                    {
                        lblError.Visible = true;
                        Label1.Text = "Error Fetching Taxonomy List";
                    }

                }




            }
            string accountid = ddlAdvisor.SelectedValue;


            if (accountid == "0")
            {
                //if (bRoles)
                //{
                // int RoleId = ddlRoles.SelectedIndex;
                if (!isEmpty)
                {
                    string RoleId = ddlRoles.SelectedValue;
                    if (ViewState["RoleId"] != null)
                    {
                        Check = ViewState["RoleId"].ToString();
                    }
                    else
                    {
                        Check = RoleId;
                    }

                    if (Check == RoleId && RoleId != "")
                    {
                        DataTable dtHousehold = GetHouseHoldList(CurrentUser, RoleId);
                        if (dtHousehold.Rows.Count > 0)
                        {
                            ShowHousehold(dtHousehold);
                            ViewState["RoleId"] = RoleId;
                        }
                        else
                        {
                            Label1.Visible = true;
                            Label1.Text = "No Records Found";
                        }
                    }
                    else if (RoleId == "")
                    {
                        ddlRoles.Visible = false;
                        lblRoles.Visible = false;
                        Label1.ForeColor = System.Drawing.Color.Red;
                        Label1.Visible = true;
                        Label1.Text = "This dashboard is for Client Services (Advisor, Associates, Supervisor and Overseer). You do not have permission to view this dashboard.";
                    }

                }
            }
            else
            {
                ddlRoles.Visible = false;
                lblRoles.Visible = false;
                ddlAccountchange();
            }
            //}
            //else
            //{

            //    lblRoles.Visible = false;
            //    ddlRoles.Visible = false;
            //    Label1.Visible = true;
            //    Label1.Text = "You don't have permissions to view this page";
            //}
            //}
            //DataTable dtHousehold = GetHouseHoldList(CurrentUser);
            //if (dtHousehold.Rows.Count > 0)
            //{
            //    ShowHousehold(dtHousehold);
            //}
            //else
            //{
            //    Label1.Visible = true;
            //    Label1.Text = "No Records Found";
            //}

        }
        catch (Exception ex)
        {
            Label1.Visible = true;
            Label1.Text = "ERROR:" + ex.Message;
        }
    }
    public void BindddlRoles(string CurrentUser)
    {
        try
        {
            DB clsDB = new DB();

            string query = "SP_S_Role_DashBoard @CurrentUserID =" + CurrentUser;
            DataSet ds = clsDB.getDataSet(query);

            if (ds.Tables[0].Rows.Count > 0)
            {
                bRoles = true;
                ddlRoles.DataTextField = "RoleName";
                ddlRoles.DataValueField = "RoleID";
                ddlRoles.DataSource = ds.Tables[0];
                ddlRoles.DataBind();
            }
        }
        catch (Exception ex)
        {
            bRoles = false;
        }
    }
    public void ShowHousehold(DataTable dtHousehold)
    {

        int dtHouseholdCount = dtHousehold.Rows.Count;
        //int index = ddlItems.SelectedIndex;
        for (int i = 0; i < dtHouseholdCount; i++)
        {
            string Accountid = string.Empty;
            string Name = string.Empty;
            string Advisor = string.Empty;
            string Associate = string.Empty;
            string Supervisor = string.Empty;
            string Overseer = string.Empty;
            string AUM = string.Empty;
            string AsofDate = string.Empty;

            string PrimaryConatctName = string.Empty;
            string PrimaryAddressLine1 = string.Empty;
            string PrimaryAddressLine2 = string.Empty;
            string PrimaryAddressLine3 = string.Empty;
            string PrimaryContactAddPhone = string.Empty;

            Accountid = dtHousehold.Rows[i]["Accountid"].ToString();
            Name = dtHousehold.Rows[i]["Name"].ToString();
            Advisor = dtHousehold.Rows[i]["AdvisorNameTxt"].ToString();
            Associate = dtHousehold.Rows[i]["AssociateNameTxt"].ToString();
            Supervisor = dtHousehold.Rows[i]["SupervisorNameTxt"].ToString();
            Overseer = dtHousehold.Rows[i]["OverseerNameTxt"].ToString();
            AUM = dtHousehold.Rows[i]["AUMMny"].ToString();
            AsofDate = dtHousehold.Rows[i]["AsofDate"].ToString();

            PrimaryConatctName = dtHousehold.Rows[i]["PrimaryContactName"].ToString();
            PrimaryAddressLine1 = dtHousehold.Rows[i]["PrimaryContactAddLine1"].ToString();
            PrimaryAddressLine2 = dtHousehold.Rows[i]["PrimaryContactAddLine2"].ToString();
            PrimaryAddressLine3 = dtHousehold.Rows[i]["PrimaryContactAddLine3"].ToString();
            PrimaryContactAddPhone = dtHousehold.Rows[i]["PrimaryContactAddPhone"].ToString();

            string AccountContactUUID = dtHousehold.Rows[i]["AccountContactUUID"].ToString();
            string AccountContactName = dtHousehold.Rows[i]["AccountContactName"].ToString();
            //adding pannel
            Panel panel = new Panel();

            //panel.HorizontalAlign = HorizontalAlign.Center;
            // panel.CssClass = "panel-group";
            panel.CssClass = "panel panel-default";
            panel.Style.Add("flex-align", "center");
            panel.BorderColor = System.Drawing.Color.FromName("#006699");
            // panel.BackColor = System.Drawing.Color.FromName("#B7DDE8");
            panel.EnableViewState = true;

            panel.ID = "pnl" + i;
            panel.BorderWidth = 1;
            panel.Width = 1300;
            HtmlTable table = new HtmlTable();
            //table.Border = 2;
            table.Align = "center";
            table.Width = "1300px";
            HtmlTableRow HeaderRow = new HtmlTableRow();
            HtmlTableCell HeaderCell = new HtmlTableCell();
            HtmlTableRow row = new HtmlTableRow();
            HtmlTableCell cell = new HtmlTableCell();
            HtmlTableRow row2 = new HtmlTableRow();

            HeaderCell.InnerText = Name;
            HeaderCell.Width = "250px";
            HeaderCell.Style.Add("font-weight", "bold");
            HeaderCell.Style.Add("cssclass", "panel-heading");
            HeaderCell.Style.Add("color", "White");

            //adding label

            #region comment
            //HtmlTable tbl = new HtmlTable();
            //// tbl.Border = 5;
            //// tbl.Width = "1000px";

            // HtmlTableCell cell0 = new HtmlTableCell();
            // lbl.Width = 100;
            #endregion


            HtmlTableRow row1 = new HtmlTableRow();
            HtmlTableCell cell1 = new HtmlTableCell();

            //lbl.ID = "lbl0";
            cell1.InnerHtml = "<b>Advisor: </b>" + Advisor;
            cell1.Align = "Left";
            cell1.Width = "220px";
            #region comment
            //  cell0.Controls.Add(lbl);

            //Label lbl = new Label();
            //lbl.ID = "lbl0";
            //lbl.Text = "Advisor:" + Advisor + "&nbsp;&nbsp;&nbsp;";
            //cell0.Controls.Add(lbl);

            // HtmlTableCell cell1 = new HtmlTableCell();
            // cell1.Width = 100;
            #endregion
            HtmlTableCell cell2 = new HtmlTableCell();
            //lbl1.ID = "lbl1";
            cell2.InnerHtml = "<b>Associate: </b>" + Associate;
            cell2.Width = "230px";
            // cell1.Controls.Add(lbl1);

            // HtmlTableCell cell2 = new HtmlTableCell();
            //cell2.Width = 100;
            HtmlTableCell lbl2 = new HtmlTableCell();
            // lbl2.ID = "lbl2";
            AUM = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C0}", Convert.ToDecimal(AUM));
            // lbl2.InnerHtml = "<b>AUM: </b>" + AUM + " &nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>As Of: </b>"+ AsofDate;
            lbl2.InnerHtml = " ";
            lbl2.Style.Add("color", "White");
            lbl2.Width = "250px";

            #region comment
            // cell2.Controls.Add(lbl2);

            ////HtmlTableCell cell3 = new HtmlTableCell();
            //// cell3.Width = 100;
            //HtmlTableCell lbl3 = new HtmlTableCell();
            ////lbl3.ID = "lbl3";
            //lbl3.InnerHtml = "<b>Overseer: </b>" + Overseer;
            ////cell3.Controls.Add(lbl3);

            // HtmlTableCell cell4 = new HtmlTableCell();
            //  cell4.Width = 100;
            #endregion
            HtmlTableCell cell3 = new HtmlTableCell();
            // lbl4.ID = "lbl4";
            // cell3.InnerHtml = "<b>Supervisor: </b>" + Supervisor + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Overseer: </b>" + Overseer; ;
            cell3.InnerHtml = "<b>Supervisor: </b>" + Supervisor;
            cell3.Align = "Left";
            cell3.Width = "200px";
            // cell4.Controls.Add(lbl4);

            HtmlTableCell cell4 = new HtmlTableCell();
            cell4.InnerHtml = "<b>Overseer: </b>" + Overseer; ;
            cell4.Align = "Left";
            cell4.Width = "200px";
            // cell4.Controls.Add(lbl4);


            HtmlTableCell emptycol = new HtmlTableCell();
            emptycol.Width = "200px";
            emptycol.InnerHtml = "";

            HtmlTableCell emptycol1 = new HtmlTableCell();
            emptycol1.Width = "200px";
            emptycol1.InnerHtml = "";

            #region Button
            HtmlTableCell cell5 = new HtmlTableCell();
            cell5.Width = "350px";
            //cell5.Style.Add("cell-Padding", "200px");
            HtmlButton btn = new HtmlButton();
            btn.ID = AccountContactUUID; //Accountid;
            //cell.Align = "Right";
            btn.Style.Add("width", "150px");
            btn.Style.Add("height", "22px");
            btn.InnerText = "Detail View";
            btn.Style.Add("align", "right");
            //cell5.Height = "10px";
            //cell5.Width = "80px";
            //btn.Style.Add( "margin-left", "25px");
            //btn.OnClientClick += DetailView1(btn,new EventArgs());
            //btn.Click += (se, ev) => DetailView(se, ev, btn.ID);
            if (ddlAdvisor.SelectedValue != "0")
                btn.ServerClick += new EventHandler(DetailView1);
            else
                btn.ServerClick += new EventHandler(DetailViewnew);
            //btn.Attributes.Add("OnClick", "javascript :reply_click(this);");


            //Sharepoint Button
            HtmlButton btn1 = new HtmlButton();
            // btn1.ID = AccountContactName.Replace("'","");// Name;
            btn1.ID = Accountid; // ID           

            btn1.Style.Add("width", "150px");
            btn1.Style.Add("height", "22px");
            btn1.InnerText = "Sharepoint";
            btn1.Style.Add("align", "right");
            btn1.ServerClick += new EventHandler(Sharepoint);
            #endregion
            //string test = "https://greshampartners.sharepoint.com/ClientServ/Documents/Clients/Active/"+Name.Replace("Family","");
            // cell5.InnerHtml = "<a href=" + test + ">HTML5 tutorial!</a>";
            cell5.Controls.Add(btn1);
            cell5.Controls.Add(btn);



            #region 2ndRow
            HtmlTableCell cellnewrow1 = new HtmlTableCell();
            cellnewrow1.InnerHtml = "<b>Contact Name: </b>" + PrimaryConatctName; ;
            cellnewrow1.Align = "Left";
            cellnewrow1.Width = "250px";

            HtmlTableCell cellnewrow2 = new HtmlTableCell();
            cellnewrow2.InnerHtml = "<b>Mobile No: </b>" + PrimaryContactAddPhone;
            cellnewrow2.Align = "Left";
            cellnewrow2.Width = "200px";


            HtmlTableCell cellnewrow3 = new HtmlTableCell();
            cellnewrow3.InnerHtml = "<b>Address: </b>" + PrimaryAddressLine1 + " " + PrimaryAddressLine2 + " " + PrimaryAddressLine3; ;
            cellnewrow3.Align = "Left";
            cellnewrow3.Width = "400px";
            cellnewrow3.ColSpan = 3;

            #endregion




            HeaderRow.Cells.Add(HeaderCell);
            HeaderRow.Cells.Add(lbl2);
            HeaderRow.Cells.Add(emptycol);
            HeaderRow.Cells.Add(emptycol1);
            HeaderRow.Cells.Add(cell5); // Button
            // HeaderRow.Align = "Center";
            HeaderRow.BgColor = "#006699";


            row.Cells.Add(cell1);
            row.Cells.Add(cell2);
            row.Cells.Add(cell3);
            row.Cells.Add(cell4);
            row.Height = "30px";


            row2.Cells.Add(cellnewrow1);
            row2.Cells.Add(cellnewrow2);
            row2.Cells.Add(cellnewrow3);
            row2.Height = "30px";


            row.Cells.Add(cell);
            table.Rows.Add(HeaderRow);
            table.Rows.Add(row);
            table.Rows.Add(row2);

            panel.Controls.Add(table);
            form1.Controls.Add(panel);

            // form1.InnerHtml = "</br></br></br>";

        }
    }

    protected void Sharepoint(object sender, EventArgs e)
    {
        try
        {


            HtmlButton button = (HtmlButton)sender;
            string AccountId = button.ID.ToLower();

            string folderpath = string.Empty;

            string[] ids = AccountId.Split('|');
            #region old
            // commented 2_20_2019
            //AccountId = ids[0];
            //if (AccountId.ToLower() == "odonnell family")
            //{
            //    //  AccountId = "o%27donnell";
            //    AccountId = "o%27donnell family";
            //}

            // folderpath = "https://greshampartners.sharepoint.com/ClientServ/Documents/Clients/Active/" + AccountId.Replace("family", "");

            //if (CheckFolderPathExists(folderpath))
            //{


            //Commented 7_24_2019 - Page SLow Issue
            //  AccountId = AccountId.Replace("family", "");
            //if (ViewState["dtDocumentTaxonomy"] != null || ViewState["dtActiveClientList"] != null)
            //{
            //    DataTable dtDocTax = (DataTable)ViewState["dtDocumentTaxonomy"];
            //    DataTable dtActiveClient = (DataTable)ViewState["dtActiveClientList"];

            //    //Fetch URL from Sharepoint
            //    //folderpath = spcls.FetchSharepointLink(dtDocTax, dtActiveClient, AccountId);
            //    folderpath = spcls.FetchNewSpURL(dtActiveClient, AccountId);
            //    #endregion
            //    if (folderpath != null)
            //    {
            //        System.Text.StringBuilder sb = new System.Text.StringBuilder();
            //        Type tp = this.GetType();
            //        sb.Append("\n<script type=text/javascript>\n");

            //        //string test = "https://greshampartners.sharepoint.com/ClientServ/Documents/Clients/Active/" + Name.Replace("Family", "");
            //        // sb.Append("\nwindow.open('ClientServicesDashboard.aspx?id=" + buttonId + "', '_blank');");
            //        sb.Append("\nwindow.open('" + folderpath + "', 'mywindow');");

            //        sb.Append("</script>");
            //        ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
            //        //}
            //    }
            //    else
            //    {
            //        lblError.Visible = true;
            //        lblError.Text = "Sharepoint URL not found";
            //    }

            //}
            //else
            //{
            //    lblError.Visible = true;
            //    lblError.Text = "ERROR on Sharepoint List Error";
            //}
            #endregion


            if (Application["dtActiveClientList"] != null )
            {
                DataTable dtActiveClient = (DataTable)Application["dtActiveClientList"];                

                //Fetch URL from Sharepoint
                //folderpath = spcls.FetchSharepointLink(dtDocTax, dtActiveClient, AccountId);
                folderpath = spcls.FetchNewSpURL(dtActiveClient, AccountId);
            
                if (folderpath != null)
                {
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");

                    //string test = "https://greshampartners.sharepoint.com/ClientServ/Documents/Clients/Active/" + Name.Replace("Family", "");
                    // sb.Append("\nwindow.open('ClientServicesDashboard.aspx?id=" + buttonId + "', '_blank');");
                    sb.Append("\nwindow.open('" + folderpath + "', 'mywindow');");

                    sb.Append("</script>");
                    ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
                    //}
                }
                else
                {
                    lblError.Visible = true;
                    lblError.Text = "Sharepoint URL not found";
                }

            }
            else
            {
                lblError.Visible = true;
                lblError.Text = "ERROR on Sharepoint List Error";
            }
        }
        catch (Exception Ex)
        {
            lblError.Visible = true;
            lblError.Text = "ERROR on sharpoint Click: " + Ex.Message.ToString();
        }
    }

    #region COmment
    //public bool CheckFolderPathExists(String folderPath)
    //{
    //    string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
    //    string filename = @"E:\devlopment\GP\SharepointCode\DemoTest.txt";
    //    ClientContext context = new ClientContext(siteUrl);
    //    SecureString passWord = new SecureString();
    //    foreach (var c in "w!ldWind36") passWord.AppendChar(c);
    //    context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
    //    Web site = context.Web;
    //    try
    //    {
    //        //Get the required RootFolder
    //        //string barRootFolderRelativeUrl = "Shared Documents/test 2/";
    //        //  string barRootFolderRelativeUrl = folderPath;

    //        //folderPath = folderPath.Replace("\\", "/");
    //        //folderPath = "Documents/" + folderPath;
    //        //Folder barFolder = site.GetFolderByServerRelativeUrl(folderPath);

    //        //int len = folderPath.Length;
    //        //int indexlen = folderPath.IndexOf("Documents");
    //        //indexlen = indexlen + 10;
    //        //int cnt = len - indexlen;
    //        //string vNewSharePointReportFolder = folderPath.Substring(indexlen, cnt);
    //        //vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/").Replace(@"\", "/");




    //        //vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/");
    //        //vNewSharePointReportFolder = "Documents/" + vNewSharePointReportFolder;
    //        Folder barFolder = site.GetFolderByServerRelativeUrl(folderPath);

    //        // context.Load(barFolder);
    //        context.ExecuteQuery();

    //        return true;
    //    }
    //    catch (Exception ex)
    //    {

    //        return false;
    //    }

    //    //  return true;
    //}
    #endregion
    protected void DetailView1(object sender, EventArgs e)
    {

        HtmlButton button = (HtmlButton)sender;
        string AccountId = button.ID;

        string[] ids = AccountId.Split('|');
        AccountId = ids[0];

        string AccountName = string.Empty;

        if (AccountId != "")
        {
            //string sqlstr = "select top 1 Name from Account where accountid = '" + AccountId + "'";
            //DB clsDB = new DB();
            //DataSet lodataset = clsDB.getDataSet(sqlstr);
            //if (lodataset.Tables[0].Rows.Count > 0)
            //{
            //    AccountName = Convert.ToString(lodataset.Tables[0].Rows[0]["Name"]);
            //}

            //Response.Redirect("ClientServicesDashboard.aspx?id=" + buttonId);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");
            // sb.Append("\nwindow.open('ClientServicesDashboard.aspx?id=" + buttonId + "', 'mywindow');");
            //  sb.Append("\nwindow.open('ClientServicesDashboard.aspx?Id=" + AccountId + "&Name=" + AccountName + "', '_blank');");

            sb.Append("\nwindow.open('ClientServicesDashboard.aspx?Id=" + AccountId + "', '_blank');");

            sb.Append("</script>");
            ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
        }
    }
    protected void DetailViewnew(object sender, EventArgs e)
    {

        HtmlButton button = (HtmlButton)sender;
        string AccountId = button.ID;
        string AccountName = string.Empty;

        string[] ids = AccountId.Split('|');
        AccountId = ids[0];

        if (AccountId != "")
        {
            //string sqlstr = "select top 1 Name from Account where accountid = '" + AccountId + "'";
            //DB clsDB = new DB();
            //DataSet lodataset = clsDB.getDataSet(sqlstr);
            //if (lodataset.Tables[0].Rows.Count > 0)
            //{
            //    AccountName = Convert.ToString(lodataset.Tables[0].Rows[0]["Name"]);
            //}

            //Response.Redirect("ClientServicesDashboard.aspx?id=" + buttonId);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");
            // sb.Append("\nwindow.open('ClientServicesDashboard.aspx?id=" + buttonId + "', 'mywindow');");
            //sb.Append("\nwindow.open('ClientServicesDashboard.aspx?Id=" + AccountId + "&Name=" + AccountName + "', '_blank');");

            sb.Append("\nwindow.open('ClientServicesDashboard.aspx?Id=" + AccountId + "', '_blank');");

            sb.Append("</script>");
            ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
        }
    }
    private DataTable GetHouseHoldList(string CurrentUser, string Role)
    {
        string greshamquery;
        int totalCount = 0;

        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {
            // CurrentUser = CurrentUser == "0" ? "null" : "'" + CurrentUser + "'";
            // greshamquery = "SP_S_HouseHoldList_DashBoard @CurrentUserID =" + CurrentUser;
            greshamquery = "SP_S_HouseHoldList_DashBoard @CurrentUserID =" + CurrentUser + ", @RoleID =" + Role;
            // greshamquery = "SP_S_HouseHoldList_DashBoard @CurrentUserID =" + CurrentUser + ", @HouseholdUUID =Null" + ", @RoleID =" + Role;
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;

        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            //bProceed = false;
            totalCount = 0;
            Label1.Visible = true; ;
            Label1.Text = "SP_S_HouseHoldList_DashBoard fails error desc:" + exc.Detail.InnerText;


        }
        catch (Exception exc)
        {
            //bProceed = false;
            totalCount = 0;
            Label1.Visible = true; ;
            Label1.Text = "SP_S_HouseHoldList_DashBoard fails error desc:" + exc.Message;

        }

        return ds_gresham.Tables[0];
    }
    private string GetcurrentUser(string Type)
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        //Changed Windows to - ADFS Claims Login 8_9_2019

        string strName = "";
        //Response.Write("Name1 =" + strName);
        if (Request.Url.AbsoluteUri.Contains("localhost"))
            strName = "corp\\jpilson";
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            strName= claimsIdentity.Name;
        }
        // strName = "corp\\DColton";
        //   strName = "corp\\david";
        // strName = "corp\\allison";
        //strName = "corp\\jmasa";
        // strName = "corp\\gbhagia";
        //////////
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        sqlstr = "select top 1 internalemailaddress,systemuserid,fullName from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            if (Type == "Name")
                return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["fullName"]);
            else
                return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
            //return UserID = "DFCE21B1-B81E-E211-A2B7-0002A5443D86";
        }
        else
        {
            return UserID = "";
        }
    }
    protected void ddlRoles_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblRoles.Visible = true;
        ddlRoles.Visible = true;
        Label1.Visible = false;

        string RoleId = ddlRoles.SelectedValue;

        DataTable dtHousehold = GetHouseHoldList(CurrentUser, RoleId);
        if (dtHousehold.Rows.Count > 0)
        {
            ShowHousehold(dtHousehold);
            ViewState["RoleId"] = RoleId;
        }
        else
        {
            Label1.Visible = true;
            Label1.Text = "No Records Found";
        }
    }

    public void BindddlAdvisor()
    {
        try
        {
            DB clsDB = new DB();

            string query = "SP_S_HouseHoldList_DashBoard";
            DataSet ds = clsDB.getDataSet(query);

            DataTable newdt = new DataTable();
            newdt.Columns.Add("Name");
            newdt.Columns.Add("AccountID");
            //  newdt = ds.Tables[0];
            foreach (DataRow rw in ds.Tables[0].Rows)
            {
                DataRow row2 = newdt.NewRow();
                row2["Name"] = rw["Name"];
                row2["AccountID"] = rw["AccountID"];
                newdt.Rows.Add(row2);
            }

            DataView dv = new DataView(newdt);
            DataTable distinctValues = dv.ToTable(true, "Name", "AccountID");


            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("AccountID");
            DataRow row = dt.NewRow();
            row["Name"] = "Select";
            row["AccountID"] = "0";
            dt.Rows.Add(row);

            foreach (DataRow rw in distinctValues.Rows)
            {
                DataRow row1 = dt.NewRow();
                row1["Name"] = rw["Name"];
                row1["AccountID"] = rw["AccountID"];
                dt.Rows.Add(row1);
            }

            if (distinctValues.Rows.Count > 0)
            {
                bRoles = true;
                ddlAdvisor.DataTextField = "Name";
                ddlAdvisor.DataValueField = "AccountID";
                ddlAdvisor.DataSource = dt;//ds.Tables[0];
                ddlAdvisor.DataBind();
            }


            //   ddlAdvisor.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0", true));

        }
        catch (Exception ex)
        {
            bRoles = false;
        }
    }

    protected void ddlAdvisor_SelectedIndexChanged1(object sender, EventArgs e)
    {
        string HouseholdId = ddlAdvisor.SelectedValue;
        //  HouseholdId = HouseholdId.Replace("-", "");
        string greshamquery;
        int totalCount = 0;
        if (HouseholdId != "0")
        {
            string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 400;
            SqlDataAdapter dagersham = new SqlDataAdapter();
            DataSet ds_gresham = new DataSet();

            try
            {

                greshamquery = "SP_S_HouseHoldList_DashBoard @HouseholdUUID ='" + HouseholdId + "'";
                // greshamquery = "SP_S_HouseHoldList_DashBoard @CurrentUserID =" + CurrentUser + ", @RoleID =" + Role;
                dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                ds_gresham = new DataSet();
                dagersham.Fill(ds_gresham);
                totalCount = ds_gresham.Tables[0].Rows.Count;

            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                //bProceed = false;
                totalCount = 0;
                Label1.Visible = true; ;
                Label1.Text = "SP_S_HouseHoldList_DashBoard fails error desc:" + exc.Detail.InnerText;


            }
            catch (Exception exc)
            {
                //bProceed = false;
                totalCount = 0;
                Label1.Visible = true; ;
                Label1.Text = "SP_S_HouseHoldList_DashBoard fails error desc:" + exc.Message;

            }
            Label1.Text = "";
            isEmpty = true;
            ShowHousehold(ds_gresham.Tables[0]);
        }
    }

    public void ddlAccountchange()
    {
        string HouseholdId = ddlAdvisor.SelectedValue;
        //  HouseholdId = HouseholdId.Replace("-", "");
        string greshamquery;
        int totalCount = 0;
        if (HouseholdId != "0")
        {
            string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 400;
            SqlDataAdapter dagersham = new SqlDataAdapter();
            DataSet ds_gresham = new DataSet();

            try
            {

                greshamquery = "SP_S_HouseHoldList_DashBoard @HouseholdUUID ='" + HouseholdId + "'";
                // greshamquery = "SP_S_HouseHoldList_DashBoard @CurrentUserID =" + CurrentUser + ", @RoleID =" + Role;
                dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                ds_gresham = new DataSet();
                dagersham.Fill(ds_gresham);
                totalCount = ds_gresham.Tables[0].Rows.Count;

            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                //bProceed = false;
                totalCount = 0;
                Label1.Visible = true; ;
                Label1.Text = "SP_S_HouseHoldList_DashBoard fails error desc:" + exc.Detail.InnerText;


            }
            catch (Exception exc)
            {
                //bProceed = false;
                totalCount = 0;
                Label1.Visible = true; ;
                Label1.Text = "SP_S_HouseHoldList_DashBoard fails error desc:" + exc.Message;

            }
            Label1.Text = "";
            isEmpty = true;
            ShowHousehold(ds_gresham.Tables[0]);
        }

    }
    protected void btnReset_Click(object sender, EventArgs e)
    {
        Response.Redirect(Request.RawUrl);
    }
}