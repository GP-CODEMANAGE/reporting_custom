using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using Spire.Xls;
using Microsoft.Xrm.Sdk;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

public partial class BatchReport_TestData : System.Web.UI.Page
{
    Logs lg = new Logs();
    public StreamWriter sw = null;
    public string Filename = "";
    public DateTime NearestQuarterEnd(DateTime date)
    {
        IEnumerable<DateTime> candidates =//Enumarateor used To show candiadtes by LastDate
            QuartersInYear(date.Year).Union(QuartersInYear(date.Year - 1));
        return candidates.Where(d => d < date.Date).OrderBy(d => d).Last();
    }
    IEnumerable<DateTime> QuartersInYear(int year)
    {
        return new List<DateTime>() {
        new DateTime(year, 3, 31),
        new DateTime(year, 6, 30),
        new DateTime(year, 9, 30),
        new DateTime(year, 12, 31),
    };
    }
    public IOrganizationService GetCrmService()
    {
        ClientCredentials Credentials = new ClientCredentials();
        string Server = "TEst";
        string str = string.Empty;
        //if (Server.ToLower() == "test")
        //{
        //    str = "corp\\skane";
        //}
        //else
        //{
        //    IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        //    str = claimsIdentity.Name;

        //}

        //string UserID = string.Empty;



        // *** for specific user credential *********  //
        Credentials.UserName.UserName = "GirishB.Infograte@wefamilyoffices.com";
        Credentials.UserName.Password = "infogr@te2019";




        // Credentials.UserName.UserName = "corp\\crmadmin";
        // Credentials.UserName.Password = "W!gmxF26ggw]";

        //This URL needs to be updated to match the servername and Organization for the environment.
        Uri OrganizationUri = new Uri("https://wefamilyoffices.api.crm.dynamics.com/XRMServices/2011/Organization.svc");
        // Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
        Uri HomeRealmUri = null;



        //OrganizationServiceProxy serviceProxy; 
        Microsoft.Xrm.Sdk.IOrganizationService service;
        Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy serviceProxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null);

        // This statement is required to enable early-bound type support.
        serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

        service = (Microsoft.Xrm.Sdk.IOrganizationService)serviceProxy;
        //Guid _UserID = new Guid(UserID);

        //serviceProxy.CallerId = _UserID;

        return service;
    }


    public IOrganizationService ConnectToCRM()
 
        {
            
                string orgName = "wefamilyoffices";// ConfigurationManager.AppSettings["OrgName"].ToString();

                string userName = "GirishB.Infograte@wefamilyoffices.com";

                string password = "gbzncqtbllflldhv";

                string crmRegion = "crm";

                ClientCredentials clientCredentials = new ClientCredentials();

                clientCredentials.UserName.UserName = userName;//userName + "@" + orgName + ".onmicrosoft.com";

                clientCredentials.UserName.Password = password;

                // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                // Get the URL from CRM, Navigate to Settings -> Customizations -> Developer Resources

                // Copy and Paste Organization Service Endpoint Address URL

                IOrganizationService organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri("https://" + orgName + ".api." + crmRegion + ".dynamics.com/XRMServices/2011/Organization.svc"),

                 null, clientCredentials, null);

                if (organizationService != null)
                {

                    Guid userid = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).UserId;

                }
           
            return organizationService;
        
 
        }
    public void Fetch(IOrganizationService _orgService)
        {
            ConditionExpression condition1 = new ConditionExpression();
            condition1.AttributeName = "lastname";
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add("Signer 1");

            ConditionExpression condition2 = new ConditionExpression();
            condition2.AttributeName = "firstname";
            condition2.Operator = ConditionOperator.Equal;
            condition2.Values.Add("Test");

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);
            filter1.Conditions.Add(condition2);

            QueryExpression query = new QueryExpression("contact");
          //  query.ColumnSet.AddColumns("firstname", "lastname", "mobilephone", "new_mfaflg");
            query.ColumnSet.AddColumns("firstname", "new_mfaflg");
            query.Criteria.AddFilter(filter1);
           
            EntityCollection result1 = _orgService.RetrieveMultiple(query);
            foreach (var a in result1.Entities)
            {
             //   Console.WriteLine("Name: " + a.Attributes["firstname"] + " " + a.Attributes["lastname"]);
                ColumnSet c1 = new ColumnSet(true);
               // Guid ContactId =  a.Attributes["contactid"];
                Entity retEnt = (Entity)_orgService.Retrieve("contact", new Guid(a.Attributes["contactid"].ToString()), c1);

                bool MFAflg = Convert.ToBoolean( retEnt.Attributes["new_mfaflg"]);
                    string PhoneNumber =retEnt.Attributes["mobilephone"].ToString();
                    string CountryCode = retEnt.Attributes["new_countrycode"].ToString();
                    string contactid = retEnt.Attributes["contactid"].ToString();


                    UpdateBankAccount(contactid, PhoneNumber, CountryCode, MFAflg, _orgService, "new_signer_1", "", "", "new_signer1mfa");
            }

        }
    public void Fetch1(IOrganizationService _orgService)
    {
        ConditionExpression condition1 = new ConditionExpression();
        condition1.AttributeName = "lastname";
        condition1.Operator = ConditionOperator.Equal;
        condition1.Values.Add("Banker1");

        ConditionExpression condition2 = new ConditionExpression();
        condition2.AttributeName = "firstname";
        condition2.Operator = ConditionOperator.Equal;
        condition2.Values.Add("Test");

        FilterExpression filter1 = new FilterExpression();
        filter1.Conditions.Add(condition1);
        filter1.Conditions.Add(condition2);

        QueryExpression query = new QueryExpression("contact");
        //  query.ColumnSet.AddColumns("firstname", "lastname", "mobilephone", "new_mfaflg");
        query.ColumnSet.AddColumns("firstname", "new_mfaflg");
        query.Criteria.AddFilter(filter1);

        EntityCollection result1 = _orgService.RetrieveMultiple(query);
        foreach (var a in result1.Entities)
        {
            //   Console.WriteLine("Name: " + a.Attributes["firstname"] + " " + a.Attributes["lastname"]);
            ColumnSet c1 = new ColumnSet(true);
            // Guid ContactId =  a.Attributes["contactid"];
            Entity retEnt = (Entity)_orgService.Retrieve("contact", new Guid(a.Attributes["contactid"].ToString()), c1);



            string contactid = retEnt.Attributes["contactid"].ToString();
            string Gender = retEnt.FormattedValues["gendercode"].ToString();
            string FullName = retEnt.Attributes["fullname"].ToString();
            string EmailID = retEnt.Attributes["emailaddress1"].ToString();


            // Check for Banker
            UpdateProvider(contactid, Gender, FullName, EmailID, _orgService, "new_bankerid", "new_bankernamewithsalutation", "new_bankersname", "new_bankersemailid");
        }

    }
    public void UpdateProvider(string Contact_ID, string Gender, string FullName, string EmailID, IOrganizationService _orgService, string Banker, string Banker_namewithsalutation, string Banker_name, string BankerEmailID)
    {
        ConditionExpression condition1 = new ConditionExpression();
        condition1.AttributeName = Banker;
        condition1.Operator = ConditionOperator.Equal;
        condition1.Values.Add(Contact_ID);

        FilterExpression filter1 = new FilterExpression();
        filter1.Conditions.Add(condition1);


        QueryExpression query = new QueryExpression("new_provider");
        query.ColumnSet.AddColumns("new_name", "new_name");
        query.Criteria.AddFilter(filter1);

        EntityCollection result1 = _orgService.RetrieveMultiple(query);
        //All Provider having selected Banker found needs to be updated
        foreach (var a in result1.Entities)
        {
            // Update Name,EmailID,NamewithSaluation field 
            Entity Provider = new Entity("new_provider");
            Provider = _orgService.Retrieve(Provider.LogicalName, a.Id, new ColumnSet(true));
            if (Gender.ToLower() == "male")
            {
                Provider[Banker_namewithsalutation] = "Mr " + FullName;
                Provider[Banker_name] = FullName;
                Provider[BankerEmailID] = EmailID;

            }
            else if (Gender.ToLower() == "female")
            {
                Provider[Banker_namewithsalutation] = "Mrs " + FullName;
                Provider[Banker_name] = FullName;
                Provider[BankerEmailID] = EmailID;
            }

            _orgService.Update(Provider);
            Console.WriteLine("Updated contact");

        }
    }
    public void UpdateBankAccount(string Contact_ID, string Contact_phone, string Contact_countrycode, bool Contact_mfaflg, IOrganizationService _orgService, string Signer, string BankAccount_PhoneNumber, string BankAccount_Countcode, string BankAccount_SignerMFAflg)
    {
         ConditionExpression condition1 = new ConditionExpression();
         condition1.AttributeName = Signer;
            condition1.Operator = ConditionOperator.Equal;
            condition1.Values.Add(Contact_ID);

            FilterExpression filter1 = new FilterExpression();
            filter1.Conditions.Add(condition1);


            QueryExpression query = new QueryExpression("new_bankaccount");
            query.ColumnSet.AddColumns("new_name", "new_name");
            query.Criteria.AddFilter(filter1);
           
            EntityCollection result1 = _orgService.RetrieveMultiple(query);
            foreach (var a in result1.Entities)
            {
                // Update with new phone number
                Entity contact = new Entity("new_bankaccount");
                contact = _orgService.Retrieve(contact.LogicalName, a.Id, new ColumnSet(true));
                contact[BankAccount_SignerMFAflg] = Contact_mfaflg;
                _orgService.Update(contact);
                Console.WriteLine("Updated contact");

            }
    }
    protected void Page_Load(object sender, EventArgs e)
    {



      IOrganizationService service =   ConnectToCRM();
      Fetch1(service);

    //  Fetch(service);
       // ConnectToCRM_D365_Method2();

       //int numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
       //int rowsintable = 39;
       // int numIndexPageSize=0 ;
       // double total = 0;
       //      int liTotalPage = 0;
       // if(rowsintable <40) // 1st index page can be hold 39 rows
       // {
       //     numIndexPageSize = 40;

       //     total = rowsintable / numIndexPageSize;
       //     liTotalPage = Convert.ToInt32(Math.Ceiling(total));
       //     numIndexPageCount = numIndexPageCount + liTotalPage;
       // }
       // else
       // {
       //    numIndexPageCount = 2;
       //     numIndexPageSize = 45;
       //     rowsintable = rowsintable - 39;
       //     total =rowsintable/ numIndexPageSize;
       //     liTotalPage = Convert.ToInt32(Math.Ceiling(total));
       //     numIndexPageCount = numIndexPageCount + liTotalPage;
       // }
     
      






        //string EndDate = null;
        //string BatchFileName = "GenerateExcel(Consolida tePdfFi leName, dtset)";
        //BatchFileName = BatchFileName.Remove(BatchFileName.Length - 9, 5);

        //DateTime dtAsOfDate1 = DateTime.Now;
        //DateTime lastDay = new DateTime(dtAsOfDate1.Year, dtAsOfDate1.Month, 1); //1st Day of Current Month
        //lastDay = lastDay.AddDays(-1);  //last date of previous month

        //DateTime date = DateTime.Now;
        //DateTime quarterEnd = NearestQuarterEnd(date);
        //EndDate = quarterEnd.ToShortDateString();
        //EndDate = "'" + EndDate + "'";

        //string ConsolidatePdfFileName = "ABC.xlsx";

        ////DataSet dtset = ExcellProcedure(EndDate, "01004B53-0D2B-E011-81E9-0019B9E7EE05");

        ////string FilePath = GenerateExcel(ConsolidatePdfFileName, dtset);







        //DB clsDB = new DB();
        //DateTime dtmain = DateTime.Now;
        //string LogFileName = string.Empty;
        //LogFileName = "Log-" + DateTime.Now;
        //LogFileName = LogFileName.Replace(":", "-");
        //LogFileName = LogFileName.Replace("/", "-");
        //LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        //sw = new StreamWriter(LogFileName);
        //sw.Close();
        //HttpContext.Current.Session["Filename"] = LogFileName;
        //ViewState["Filename"] = LogFileName;


        //LogFileName = (string)ViewState["Filename"];

        //Session["Filename"] = LogFileName;
        //// clsCombined.LogFileName = LogFileName;
        //clsCombinedReports clsCom = new clsCombinedReports();
        //clsCom.LogFileName = LogFileName;

        //string filenale = (string)Session["Filename"];

        //lg.AddinLogFile(Session["Filename"].ToString(), "Start Page Load " + dtmain);

        //// lg.AddinLogFile(Session["Filename"].ToString(), "Option selected " + ddlAction.SelectedItem.ToString());

        //DataSet ds = clsDB.getDataSet("Exec SP_R_WORST_MONTH_MAXDD_NEW_GA_BASEDATA @GroupName = 'Anderson, J. - J&S Anderson GA', @PositionGAFlagTxt = 'GA', @TrxnGAFlagTxt = 'GA', @AsOfDate = '9/30/2019 12:00:00 AM', @BenchMarkName = null, @AssetNameTxt = 'Cash and Equivalents,Diversified Growth,Emerging Markets,Fixed Income,Global Equity,International Equity,Liquid Real Assets,Low Volatility Hedged,Opportunistic Growth,Private Equity,Private Real Assets,U.S. Equity'");
        //Write(ds.Tables[0]);


    }

    public void Write(DataTable dt)
    {
        int[] maxLengths = new int[dt.Columns.Count];

        for (int i = 0; i < dt.Columns.Count; i++)
        {
            maxLengths[i] = dt.Columns[i].ColumnName.Length;

            foreach (DataRow row in dt.Rows)
            {
                if (!row.IsNull(i))
                {
                    int length = row[i].ToString().Length;

                    if (length > maxLengths[i])
                    {
                        maxLengths[i] = length;
                    }
                }
            }
        }

        string val1 = string.Empty;
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            val1 = val1 + "|" + dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2);

            //  sw.Write(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2));
        }
        lg.AddinLogFile(Session["Filename"].ToString(), val1);

        string val2 = string.Empty;
        foreach (DataRow row in dt.Rows)
        {
            val2 = "";
            for (int i = 0; i < dt.Columns.Count; i++)
            {
               
                if (!row.IsNull(i))
                {
                    val2 = val2 + "|" + row[i].ToString().PadRight(maxLengths[i] + 2);
                    //  sw.Write(row[i].ToString().PadRight(maxLengths[i] + 2));
                }
                else
                {
                    val2 = val2 + "|" + new string(' ', maxLengths[i] + 2);
                    //  sw.Write(new string(' ', maxLengths[i] + 2));
                }
            }
            lg.AddinLogFile(Session["Filename"].ToString(), val2);

        }



    }


    private DataSet ExcellProcedure(string EndDate, string BatchId)
    {
        string greshamquery;
        int totalCount = 0;

        //  SqlConnection Gresham_con = new SqlConnection("Password=slater6;Persist Security Info=False;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=GP-PRODDB");
        SqlConnection Gresham_con = new SqlConnection("");
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {
         //   greshamquery = "SP_S_SEC_DATADUMP @BatchUUID='" + BatchId + "',@AsofDT=" + EndDate;
            greshamquery = "SP_S_SEC_DATADUMP @BatchUUID = '01004B53-0D2B-E011-81E9-0019B9E7EE05',@AsofDT = '20190630'";
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            dagersham.SelectCommand.CommandTimeout = 600;
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
           // sw.WriteLine("Stored procedure executed succesfully" + DateTime.Now);
            ////----------------------------------TAble1----------------------------
            //greshamquery = "SP_S_BATCH_STATUS @EndDT=" + EndDate;

            //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            //ds_gresham = new DataSet();

            //dagersham.Fill(ds_gresham, "Table1");
            ////----------------------------------TAble2----------------------------
            //greshamquery = "SP_S_BATCH_STATUS @EndDT=" + EndDate;
            //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            //dagersham.Fill(ds_gresham, "Table2");

            ////totalCount = ds_gresham.Tables[0].Rows.Count;

        }
        catch (Exception exc)
        {
            Response.Write("ERROR: " + exc.Message.ToString());
            //totalCount = 0;
            //lblError.Visible = true;
            //lblError.Text = "sp_S_batch sp fails error desc:" + exc.Message;
        }

        return ds_gresham;
    }
    public string GenerateExcel(string FileName, DataSet ds)
    {
        #region Spire License Code
        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
     //   string License = ConfigurationSettings.AppSettings["SpireLicense"].ToString();
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion
        // DataSet ds=null;
        // DataTable dtTable1 = null;
        try
        {
            //if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\ReportOutput"))
            //{
            //    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\ReportOutput");
            //}

            //string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            //string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            //string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            String lsFileNamforFinalXls = FileName;// "ManagerPerformance_" + strMonth + "_" + strDay + "_" + strYear + ".xlsx";
           // string ExcelFilePath = System.Windows.Forms.Application.StartupPath + "\\ReportOutput\\" + lsFileNamforFinalXls;
            string ExcelFilePath = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FileName);

            if (System.IO.File.Exists(ExcelFilePath))
            {
                System.IO.File.Delete(ExcelFilePath);
            }

            //  string ExcelFilePath = ExcelfilePath + "TradingAppRecon" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

            #region EPP code
            //FileInfo newFile = new FileInfo(ExcelFilePath);

            //using (OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(newFile))
            //{
            //    //if (ds.Tables["Table1"] != null)
            //    //{
            //    //    OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add("sheet1");
            //    //    ws.Cells["A1"].LoadFromDataTable(ds.Tables["Table1"], true);
            //    //    WorksheetFormatting(ws);
            //    //}
            //    //if (ds.Tables["Table2"] != null)
            //    //{
            //    //    OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add("sheet2");
            //    //    ws.Cells["A1"].LoadFromDataTable(ds.Tables["Table2"], true);
            //    //    WorksheetFormatting(ws);
            //    //}
            //    for (int i = 0; i < ds.Tables.Count; i++)
            //    {
            //        string SheetNme = ds.Tables[i].Rows[0][0].ToString();
            //        string GroupName = ds.Tables[i].Rows[0][1].ToString();
            //        i++;
            //        ds.Tables[i].Columns.Add("GroupName");
            //        for (int k = 0; k < ds.Tables[i].Rows.Count; k++)
            //        {
            //            ds.Tables[i].Rows[k]["GroupName"] = GroupName;
            //        }
            //        OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add(SheetNme);
            //        if (ds.Tables[i].Rows.Count > 0)
            //        {
            //            ws.Cells["A1"].LoadFromDataTable(ds.Tables[i], true);
            //            WorksheetFormatting(ws);
            //        }
            //        else
            //        {
            //            ws.Cells["J9:L10"].Merge = true;
            //            ws.Cells["J9:L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            //            ws.Cells["J9:L10"].Value = "No Data Found";
            //            ws.Cells["J9:L10"].Style.Font.Size = 16;
            //            //ws.Cells["J9:L10"].Style.Fill.BackgroundColor.SetColor(Color.Red);
            //        }

            //    }
            //    pck.Save();
            //}
            #endregion
            int SheetNo = 0;
            Workbook book = new Workbook();
            book.Version = ExcelVersion.Version2016;
            book.CreateEmptySheets(ds.Tables.Count / 2);
            for (int i = 0; i < ds.Tables.Count; i++)
            {

                string SheetNme = ds.Tables[i].Rows[0][0].ToString();
                string GroupName = ds.Tables[i].Rows[0][1].ToString();
                i++;
                ds.Tables[i].Columns.Add("GroupName");
                for (int k = 0; k < ds.Tables[i].Rows.Count; k++)
                {
                    ds.Tables[i].Rows[k]["GroupName"] = GroupName;
                }


                Worksheet sheet = book.Worksheets[SheetNo];
                sheet.Name = SheetNme;
                if (ds.Tables[i].Rows.Count > 0)
                {
                    sheet.Range[1, 1, 1, ds.Tables[i].Columns.Count].Style.Font.IsBold = true;

                    sheet.InsertDataTable(ds.Tables[i], true, 1, 1);
                    sheet.Range[1, 1, ds.Tables[i].Rows.Count, ds.Tables[i].Columns.Count].AutoFitColumns();
                    sheet.Range[1, 1, ds.Tables[i].Rows.Count, ds.Tables[i].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;
                }
                SheetNo++;
            }

            book.SaveToFile(ExcelFilePath);
          //  book.SaveToFile(ExcelFilePath, ExcelVersion.Version2016);

            string vContain = "Excel Report Generated Succesfully ";
            sw.WriteLine(vContain);
            return ExcelFilePath;
        }
        catch (Exception e)
        {

            string vContain = "Excel Report Genration Fail,  Error " + e.ToString();
            sw.WriteLine(vContain);
            //LG.AddinLogFile(Form1.vLogFile, vContain);
            //   lblMsg.Text = vContain;

            return "";
        }
    }
}