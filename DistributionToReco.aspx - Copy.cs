using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Data;
using System.IO;
using Microsoft.SqlServer.Management.Common;
using System.Data.SqlClient;
using Microsoft.SqlServer.Management.Smo.Agent;
//using CrmSdk;
using System.util;
using Microsoft.SqlServer.Management.Smo;
using System.Security.Principal;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.Configuration;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using Spire.Xls;
using System.IO.Compression;
using System.Threading;
using Microsoft.IdentityModel.Claims;

public partial class DistributionToReco : System.Web.UI.Page
{
    GeneralMethods clsGM = new GeneralMethods();
    string sqlstr = string.Empty;
    bool bProceed = false;
    int Dload = 0;
    DB clDB = null;
    Dictionary<string, string> lstFile = new Dictionary<string, string>();
    string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
    List<Int32> count = new List<int>();
    string con = AppLogic.GetParam(AppLogic.ConfigParam.DBTransactions);//"Data Source=sql01;User ID=MPIUser;Initial Catalog=TransactionLoad_DB;Persist Security Info=True;Password=slater6;";
    string DTSFilePath = AppLogic.GetParam(AppLogic.ConfigParam.DTSFilePath);
    bool bpopup = false;
    string msg = string.Empty;
    string TempPath = string.Empty;
    int ZipFileCount = 0;
    // bool bProceed = false;
    int countSelectedFund = 0;
    bool bcapital = false;
    bool bdistribution = false;
    string ErrorOccured = string.Empty;
    IOrganizationService service = null;
    int countpopup = 0;
    int jobfailcont = 0;
    int jobsucesscnt = 0;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Fillddl();
        }
        GetcurrentUser();
    }

    protected void btnUpload_Click(object sender, EventArgs e)
    {
        string vCall = string.Empty;
        lblError.Text = "";
        lblError1.Text = "";
        string FundValue = "";
        string vFundName = "";
        ErrorOccured = "";
        string vDate = txtDate.Text;
        string ALpsFile = string.Empty;
        clDB = new DB();
        if (fuDist.HasFile)
        {
            if (System.IO.Path.GetExtension(fuDist.FileName).ToUpper() == ".ZIP")
            {
                if (txtDate.Text != "")
                {
                    if (rbCapitalCall.Checked || rbDistribution.Checked)
                    {
                        if (rbCapitalCall.Checked)
                        {
                            vCall = "1";
                            bcapital = true;
                        }
                        else
                        {
                            vCall = "2";
                            bdistribution = true;
                        }
                        Random rndNo = new Random();
                        string iRnNo = rndNo.Next(1, 9999999).ToString();
                        try
                        {
                            string conPath = AppLogic.GetParam(AppLogic.ConfigParam.DTSFilePath);  //@"\\GRPAO1-VWFS01\Shared$\Invoice\CRM2016\";

                            //     ConfigurationManager.AppSettings["CapitalCallFilePath"].ToString();

                            string user = AppLogic.GetParam(AppLogic.ConfigParam.UserName).ToString();
                            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();


                            // using (new Impersonation("corp", "gbhagia", "51ngl3malt"))
                            DateTime dt = DateTime.Now;

                            string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
                            string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
                            string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();

                            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
                            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

                            // string strUserName = HttpContext.Current.User.Identity.Name.ToString();

                            //Changed Windows to - ADFS Claims Login 8_9_2019
                            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
                            string strUserName = claimsIdentity.Name;

                            strUserName = strUserName.Substring(strUserName.IndexOf("\\") + 1);

                            string dateTime = "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;
                            string TempFolderName = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;

                            TempPath = Server.MapPath("") + @"\ExcelTemplate\Tempfolder\" + TempFolderName + "\\";

                            ViewState["Temppath"] = TempPath;

                            ZipFileCount = ProcessZip(TempPath, dateTime);

                            //If Funds are not selected -Run for all files present in the Zip
                            if (countSelectedFund == 0)
                            {
                                Proces_ALL_FUND(TempPath, service, bcapital, bdistribution);
                                ViewState["msg"] = msg;
                                ViewState["error"] = ErrorOccured;
                            }


                            #region Not in use
                            //using (new Impersonation("corp", user, Pass))
                            //{
                            //    if (rbCapitalCall.Checked)
                            //    {
                            //        string extension = System.IO.Path.GetExtension(fuDist.FileName);
                            //        string FilenameWOExtension = System.IO.Path.GetFileNameWithoutExtension(fuDist.FileName);
                            //        string strFileName = clDB.FileName("DistToRecoCapitalcall");
                            //        //  string strFileName = "GDGSCapitalCallSampleFileLayout" + extension;
                            //        string ExcelSavePath = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FilenameWOExtension + extension);
                            //        if (File.Exists(ExcelSavePath))
                            //            File.Delete(ExcelSavePath);
                            //        fuDist.PostedFile.SaveAs(ExcelSavePath);

                            //        /** clean up Alps file*/
                            //        ALpsFile = ReadAlPsFile(ExcelSavePath, "Capital Call", strFileName);

                            //        if (File.Exists(Path.Combine(conPath, strFileName)))
                            //        {
                            //            File.Delete(Path.Combine(conPath, strFileName));
                            //            //File.Move(ExcelSavePath, conPath, strFileName);
                            //            File.Move(ALpsFile, Path.Combine(conPath, strFileName));
                            //        }
                            //        else
                            //            // fuDist.PostedFile.SaveAs(Path.Combine(conPath, strFileName));
                            //            File.Move(ALpsFile, Path.Combine(conPath, strFileName));
                            //    }
                            //    else if (rbDistribution.Checked)
                            //    {
                            //        string extension = System.IO.Path.GetExtension(fuDist.FileName);
                            //        string FilenameWOExtension = System.IO.Path.GetFileNameWithoutExtension(fuDist.FileName);

                            //        string strFileName = clDB.FileName("DistToRecoDistribution");
                            //        // string strFileName = "GDGSDistributionSampleFileLayout" + extension;

                            //        string ExcelSavePath = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FilenameWOExtension + extension);
                            //        if (File.Exists(ExcelSavePath))
                            //            File.Delete(ExcelSavePath);
                            //        fuDist.PostedFile.SaveAs(ExcelSavePath);

                            //        /** clean up Alps file ****/
                            //        ALpsFile = ReadAlPsFile(ExcelSavePath, "Distribution", strFileName);

                            //        if (File.Exists(Path.Combine(conPath, strFileName)))
                            //        {
                            //            File.Delete(Path.Combine(conPath, strFileName));
                            //            //File.Move(ExcelSavePath, conPath, strFileName);
                            //            File.Move(ALpsFile, Path.Combine(conPath, strFileName));
                            //        }
                            //        else
                            //            // fuDist.PostedFile.SaveAs(Path.Combine(conPath, strFileName));
                            //            File.Move(ALpsFile, Path.Combine(conPath, strFileName));



                            //    }
                            //    ViewState["IsUpload"] = 1;
                            //    //lblError.Text = "File Upload Successfully.";
                            //    //lblError.ForeColor = System.Drawing.Color.Green;
                            //    //btnGDGSReco.Visible = true;
                            //    //btnUpload.Visible = false;...................

                            //}

                            #endregion

                        }


                        //}
                        //    string FilePath = HttpContext.Current.Server.MapPath(@"~\UploadFiles\" + iRnNo + fuDist.FileName);
                        //    fuDist.PostedFile.SaveAs(FilePath);
                        //    lblError.Text = "File Upload Successfully.";

                        catch (Exception ee)
                        {
                            lblError.Text = "File Upload Fail. Please Try Agian." + ee.Message;
                            ViewState["IsUpload"] = 0;
                        }


                        //  lblError1.Text = ErrorOccured;

                        if (bpopup || countpopup > 0)
                        {
                            //Button3_Click(Button1, EventArgs.Empty);


                            Button1.Visible = true;
                            System.Text.StringBuilder sb = new System.Text.StringBuilder();
                            Type tp = this.GetType();
                            sb.Append("\n<script type=text/javascript>\n");
                            sb.Append("var bt = window.document.getElementById('Button1');\n");
                            //sb.Append("if(confirm('would you like to continue creating the GDGS Recommendation?.'))\n{");

                            sb.Append("if(!alert('Done creating fund recommendations, ready to create the GDGS recommendations.  Click OK to continue.'))\n{");
                            sb.Append("\nwindow.document.getElementById('Hidden2').value='1';");
                            sb.Append("\n txtDateClear();");
                            sb.Append(("\n bt.click();\n"));
                            sb.Append("\n}");
                            sb.Append("else\n{");
                            // sb.Append(("\nwindow.document.getElementById('Hidden2').value='0';"));
                            sb.Append(("\nwindow.document.getElementById('Hidden2').value='0';"));
                            sb.Append(("\n bt.click();\n"));
                            sb.Append("\n txtDateClear();");
                            sb.Append("\n}");
                            sb.Append("</script>");
                            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                            //Button1_Click(sender,this);
                        }

                    }
                    else
                    {
                        lblError.Text = "Please Select File Type";
                        lblError.ForeColor = System.Drawing.Color.Red;
                    }

                }
                else
                {
                    lblError.Text = "Please Select Date";
                    lblError.ForeColor = System.Drawing.Color.Red;
                }
            }
            else
            {
                lblError.Text = "Please select .Zip File";
                lblError.ForeColor = System.Drawing.Color.Red;
            }

        }
        else
        {
            lblError.Text = "Please select File";
            lblError.ForeColor = System.Drawing.Color.Red;
        }

        if (jobsucesscnt == 0 && jobfailcont > 0)
            lblError1.Text = ErrorOccured;


    }

    public void Fillddl()
    {
        if (!IsPostBack)
        {
            clDB = new DB();
            DataSet ds = clDB.getDataSet("SP_S_GRESHAM_NON_MARKETABLE_FUND @Flag='0', @TypeListTxt='3,9,12'");
            if (ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (ViewState["dtFund"] == null)
                {
                    ViewState["dtFund"] = dt;
                }
            }
        }
    }



    public bool ReadAlPsFile(string inputfile, string sheetname, string Destinationfilename)
    {
        bool bproceed = true;
        bool totalflag = false;
        try
        {
            string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
            Spire.License.LicenseProvider.SetLicenseKey(License);
            Spire.License.LicenseProvider.LoadLicense();


            // string lsFileNamforFinalXls = "Alps distrubutionfile.xlsx";

            string lsFileNamforFinalXls = inputfile;
            // string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);

            string strDirectory = lsFileNamforFinalXls;


            //   string strDirectory = inputfile;


            Workbook workbook = new Workbook();

            //open an excel file

            workbook.LoadFromFile(strDirectory);


            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = sheetname;


            DataTable dt = new DataTable();
            //dt = sheet.ExportDataTable();



            /* store 9th row and 3rd column value of to identify whether it is capital call file or distribution file 
             beacuse for some file Jobs not working properly */
            string char1 = sheet.Rows[9]["C9"].Value;


            // string text = dt.Rows[5][2].ToString();
            ViewState["FileType"] = char1;

            // sheet.Name = "Capital Call";

            for (int i = sheet.Pictures.Count - 1; i >= 0; i--)
            {
                sheet.Pictures[i].Remove();
            }


            sheet.DeleteRow(1, 15);


            string freezpane = sheet.IsFreezePanes.ToString();//checking for excel contains frrezpane

            sheet.RemovePanes();//remove freezpane



            int cnt = 0;
            int Rowcnt = 0;

            foreach (CellRange range in sheet.Columns[3])
            {
                var str = range.Text;

                cnt++;
                if (range.Text != null)
                {
                    if (range.Text.ToLower().Contains("total"))
                    {
                        //int rowcount = range.RowCount;
                        Rowcnt = cnt;
                        totalflag = true;
                    }
                }
            }

            //sheet.DeleteRow(103, 104);//delete total line
            //sheet.DeleteRow(105, 109);//delete footnote
            //                          //save the excel file

            if (totalflag)
            {
                sheet.DeleteRow(Rowcnt, Rowcnt + 1);//delete total line
                sheet.DeleteRow(Rowcnt + 2, Rowcnt + 6);//delete footnote
                                                        //save the excel file
            }
            sheet.DeleteColumn(1);//delete # column that contains id  

            // string destinationfilepath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + Destinationfilename;

            workbook.SaveToFile(Destinationfilename, ExcelVersion.Version2016);

            /* destination folder */
            // return destinationfilepath;

            return bproceed;
        }
        catch (Exception ex)
        {
            return false;
        }

    }
    // insert from distribution file type
    protected List<Int32> InsertDistibuteReco(string FundId)
    {
        List<Int32> count = new List<int>();
        count.Add(0);
        count.Add(0);
        try
        {

            int iTotalCount = 0, DeleteRecord = 0;
            #region CRM Connection
            string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
            //string crmServerURL = "http://server:5555/";
            string orgName = "GreshamPartners";
            //string orgName = "Webdev";
            //  CrmService service = null;

            IOrganizationService service = null;

            try
            {
                string UserId = GetcurrentUser();

                service = clsGM.GetCrmService();
                ///  strDescription = "Crm Service starts successfully";
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
            {
                //  bProceed = false;
                //strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                //lblMessage.Text = strDescription;
            }
            catch (Exception exc)
            {
                //bProceed = false;
                //strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                //lblMessage.Text = strDescription;
            }

            #endregion
            string date = txtDate.Text;
            DataSet dsData = getDataSet("SP_S_DISTRIBUTION_GDGS @GDGSFlg = 0, @AsOfDate='" + date + "',@FundId='" + FundId + "'"); // get Excel data in DS Fromat

            if (dsData.Tables.Count > 0)
            {
                if (dsData.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in dsData.Tables[1].Rows)
                    {

                        string status = row["status"].ToString();
                        string ssi_transactionrecommendationid = row["ssi_transactionrecommendationid"].ToString();

                        Guid gUId = new Guid(ssi_transactionrecommendationid);
                        //  service.Timeout = 1000000;


                        //  service.Delete(EntityName.ssi_position.ToString(), UUID);
                        service.Delete("ssi_transactionrecommendation", gUId);

                        DeleteRecord++;
                    }
                }
            }
            count[0] = DeleteRecord;

            if (dsData.Tables.Count > 0)
            {
                if (dsData.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in dsData.Tables[0].Rows)
                    {
                        string vHouseHold = row["AccountID"].ToString();
                        string vLegalEntityId = row["Ssi_LegalEntityId"].ToString();
                        string vAccountId = row["Ssi_AccountId"].ToString();

                        string vproposedamount = row["proposedamount"].ToString();
                        string vconfirmedamount = row["confirmedamount"].ToString();
                        string vClassbfee = row["ssi_classbfee"].ToString();

                        string vTransactionTypes = row["ssi_transactiontypes"].ToString();
                        string vStatus = row["Ssi_Status"].ToString();
                        string vStatusDate = row["Ssi_StatusDate"].ToString();

                        string ssi_GreshamAdvised = row["ssi_GreshamAdvised"].ToString();
                        string vFundId = FundId;

                        try
                        {
                            //task objTask = new task();

                            //objTask.activityid = new Key();
                            //objTask.activityid.Value = Guid.NewGuid();

                            //objTask.subject = "test1";

                            //objTask.ssi_status = new Picklist();
                            //objTask.ssi_status.Value = 1;

                            //objTask.scheduledend = new CrmDateTime();
                            //objTask.scheduledend.Value = DateTime.Now.AddDays(1).ToString();

                            //service.Create(objTask);

                            //  ssi_transactionrecommendation objTranRecom = new ssi_transactionrecommendation();

                            Microsoft.Xrm.Sdk.Entity objTranRecom = new Microsoft.Xrm.Sdk.Entity("ssi_transactionrecommendation");
                            //transaction reco ID
                            // objTranRecom.ssi_transactionrecommendationid = new Key();
                            // objTranRecom.ssi_transactionrecommendationid.Value = Guid.NewGuid();

                            //HouseHold ID
                            // objTranRecom.ssi_householdid = new Lookup();
                            // objTranRecom.ssi_householdid.type = EntityName.account.ToString();
                            // objTranRecom.ssi_householdid.Value = new Guid(vHouseHold);

                            objTranRecom["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(vHouseHold));

                            //legal Entity ID
                            // objTranRecom.ssi_legalentityid = new Lookup();
                            // objTranRecom.ssi_legalentityid.type = EntityName.ssi_legalentity.ToString();
                            // objTranRecom.ssi_legalentityid.Value = new Guid(vLegalEntityId);


                            objTranRecom["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(vLegalEntityId));

                            //Fund ID
                            // objTranRecom.ssi_fundid = new Lookup();
                            // objTranRecom.ssi_fundid.type = EntityName.ssi_fund.ToString();
                            // objTranRecom.ssi_fundid.Value = new Guid(vFundId);

                            objTranRecom["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(vFundId));

                            //Account ID
                            if (vAccountId != "")
                            {
                                // objTranRecom.ssi_accountid = new Lookup();
                                // objTranRecom.ssi_accountid.type = EntityName.ssi_account.ToString();
                                // objTranRecom.ssi_accountid.Value = new Guid(vAccountId);

                                objTranRecom["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(vAccountId));
                            }
                            else
                            {
                                //  objTranRecom.ssi_accountid = new Lookup();
                                //  objTranRecom.ssi_accountid.IsNull = true;
                                //  objTranRecom.ssi_accountid.IsNullSpecified = true;

                                objTranRecom["ssi_accountid"] = null;
                            }

                            // TransactionType
                            //  objTranRecom.ssi_transactiontypes = new Picklist();
                            //  objTranRecom.ssi_transactiontypes.Value = Convert.ToInt32(vTransactionTypes);

                            objTranRecom["ssi_transactiontypes"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vTransactionTypes));


                            ////Close date investment
                            //if (vCloseDateInvestment != "")
                            //{
                            //  objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            //  objTranRecom.ssi_closedateinvestment.Value = DateTime.Parse(txtDate.Text).ToString();

                            objTranRecom["ssi_closedateinvestment"] = Convert.ToDateTime((txtDate.Text));
                            //}
                            //else
                            //{
                            //    objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            //    objTranRecom.ssi_closedateinvestment.IsNull = true;
                            //    objTranRecom.ssi_closedateinvestment.IsNullSpecified = true;
                            //}

                            ////Praposed Amount
                            if (vproposedamount != "")
                            {
                                //objTranRecom.ssi_proposedamount = new CrmMoney();
                                //objTranRecom.ssi_proposedamount.Value = Convert.ToDecimal(vproposedamount);

                                objTranRecom["ssi_proposedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vproposedamount));
                            }
                            else
                            {
                                // objTranRecom.ssi_proposedamount = new CrmMoney();
                                // objTranRecom.ssi_proposedamount.IsNull = true;
                                // objTranRecom.ssi_proposedamount.IsNullSpecified = true;

                                objTranRecom["ssi_proposedamount"] = null;
                            }

                            ////Conformed Amount
                            if (vconfirmedamount != "")
                            {
                                //objTranRecom.ssi_confirmedamount = new CrmMoney();
                                //objTranRecom.ssi_confirmedamount.Value = Convert.ToDecimal(vconfirmedamount);

                                objTranRecom["ssi_confirmedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vconfirmedamount));
                            }
                            else
                            {
                                //  objTranRecom.ssi_confirmedamount = new CrmMoney();
                                //  objTranRecom.ssi_confirmedamount.IsNull = true;
                                //  objTranRecom.ssi_confirmedamount.IsNullSpecified = true;

                                objTranRecom["ssi_confirmedamount"] = null;
                            }

                            ////class B Fee
                            if (vClassbfee != "")
                            {

                                //  objTranRecom.ssi_classbfee = new CrmMoney();
                                //  objTranRecom.ssi_classbfee.Value = Convert.ToDecimal(vClassbfee);

                                objTranRecom["ssi_classbfee"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vClassbfee));
                            }
                            else
                            {
                                // objTranRecom.ssi_classbfee = new CrmMoney();
                                // objTranRecom.ssi_classbfee.IsNull = true;
                                // objTranRecom.ssi_classbfee.IsNullSpecified = true;

                                objTranRecom["ssi_classbfee"] = null;
                            }

                            //Status 
                            if (vStatus != "")
                            {
                                // objTranRecom.ssi_status = new Picklist();
                                // objTranRecom.ssi_status.Value = Convert.ToInt16(vStatus);

                                objTranRecom["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vStatus));
                            }
                            ////Status date
                            if (vStatusDate != "")
                            {
                                //  objTranRecom.ssi_statusdate = new CrmDateTime();
                                //  objTranRecom.ssi_statusdate.Value = DateTime.Parse(vStatusDate).ToString();

                                objTranRecom["ssi_statusdate"] = Convert.ToDateTime(vStatusDate);
                            }
                            ////Response.Write(vStatusDate);

                            ////grasham Adviced checks
                            if (ssi_GreshamAdvised == "1")
                            {

                                //   objTranRecom.ssi_greshamadvised = new CrmBoolean();
                                //   objTranRecom.ssi_greshamadvised.Value = true;

                                objTranRecom["ssi_greshamadvised"] = true;
                            }
                            else
                            {
                                // objTranRecom.ssi_greshamadvised = new CrmBoolean();
                                // objTranRecom.ssi_greshamadvised.Value = false;

                                objTranRecom["ssi_greshamadvised"] = false;
                            }

                            // service.Timeout = 1000000;
                            service.Create(objTranRecom);

                            iTotalCount++;
                        }
                        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                        {
                            lblError.Text = "Error in insert reco=" + exc.ToString();
                            lblError.ForeColor = System.Drawing.Color.Green;
                            Response.Write("Error in insert reco=" + exc.ToString());
                        }
                    }
                    count[1] = iTotalCount;

                }
                else
                {
                    lblError.Text = "empty Row";
                    lblError.ForeColor = System.Drawing.Color.Green;
                }

            }
            else
            {
                lblError.Text = "Empty DS";
                lblError.ForeColor = System.Drawing.Color.Green;

            }

            //dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat for GDGS Data

            //int GDGScount = InsertDistibuteRecoGDGS(dsData);
            //count.Add(iTotalCount);
            //count.Add(GDGScount);

            return count;
        }
        catch (Exception e)
        {
            lblError.Text = "Distribution Insert Fail Please Try Again-" + e.ToString();
            lblError.ForeColor = System.Drawing.Color.Green;
            Response.Write(e.Message);

            return count;
        }

    }
    protected List<Int32> InsertCapitalCall(string FundID)
    {
        List<Int32> count = new List<int>();
        count.Add(0);
        count.Add(0);
        try
        {

            int iTotalCount = 0, DeleteRecord = 0;

            #region CRM Connection
            string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
            //string crmServerURL = "http://server:5555/";
            string orgName = "GreshamPartners";
            //string orgName = "Webdev";
            //CrmService service = null;
            IOrganizationService service = null;

            try
            {
                string UserId = GetcurrentUser();

                service = clsGM.GetCrmService();
                ///  strDescription = "Crm Service starts successfully";
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                //  bProceed = false;
                //strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                //lblMessage.Text = strDescription;
            }
            catch (Exception exc)
            {
                //bProceed = false;
                //strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                //lblMessage.Text = strDescription;
            }

            #endregion

            string date = txtDate.Text;
            //  string FundValue = lstbxPartnership.SelectedValue.ToString();
            //DataSet dsData = getDataSet("SP_S_CAPITALCALLLETTER_GDGS @GDGSFlg = 0, @AsOfDate='" + date + "',@FundId='" + FundValue + "'");

            DataSet dsData = getDataSet("SP_S_CAPITALCALLLETTER_GDGS @GDGSFlg = 0, @AsOfDate='" + date + "',@FundId='" + FundID + "'");
            //DataSet dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat

            if (dsData.Tables.Count > 0)
            {
                if (dsData.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in dsData.Tables[1].Rows)
                    {

                        string status = row["status"].ToString();
                        string ssi_transactionrecommendationid = row["ssi_transactionrecommendationid"].ToString();

                        Guid gUId = new Guid(ssi_transactionrecommendationid);
                        // service.Timeout = 1000000;
                        //   service.Delete(EntityName.ssi_transactionrecommendation.ToString(), gUId);

                        service.Delete("ssi_transactionrecommendation", gUId);
                        DeleteRecord++;
                    }
                }
            }
            count[0] = DeleteRecord;
            // count.Add(DeleteRecord);
            if (dsData.Tables.Count > 0)
            {
                if (dsData.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow row in dsData.Tables[0].Rows)
                    {
                        string vHouseHold = row["AccountID"].ToString();
                        string vLegalEntityId = row["Ssi_LegalEntityId"].ToString();

                        //   string vFundId = row["Ssi_FundId"].ToString();
                        //string vFundId = lstbxPartnership.SelectedValue.ToString();

                        string vFundId = FundID;


                        string vAccountId = row["Ssi_AccountId"].ToString();

                        //   string vCloseDateInvestment = row["Ssi_CloseDateInvestment"].ToString();

                        string vproposedamount = row["proposedamount"].ToString();
                        string vconfirmedamount = row["confirmedamount"].ToString();


                        //   string vClassbfee = row["ssi_classbfee"].ToString();

                        string vTransactionTypes = row["ssi_transactiontypes"].ToString();
                        string vStatus = row["ssi_status"].ToString();
                        string vStatusDate = row["ssi_StatusDate"].ToString();
                        string GPAdvoice = row["ssi_GreshamAdvised"].ToString();
                        //  string StatusTime=row["statusTime"].ToString();

                        try
                        {
                            //task objTask = new task();

                            //objTask.activityid = new Key();
                            //objTask.activityid.Value = Guid.NewGuid();

                            //objTask.subject = "test1";

                            //objTask.ssi_status = new Picklist();
                            //objTask.ssi_status.Value = 1;

                            //objTask.scheduledend = new CrmDateTime();
                            //objTask.scheduledend.Value = DateTime.Now.AddDays(1).ToString();

                            //service.Create(objTask);

                            // ssi_transactionrecommendation objTranRecom = new ssi_transactionrecommendation();

                            Microsoft.Xrm.Sdk.Entity objTranRecom = new Microsoft.Xrm.Sdk.Entity("ssi_transactionrecommendation");
                            //transaction reco ID
                            // objTranRecom.ssi_transactionrecommendationid = new Key();
                            // objTranRecom.ssi_transactionrecommendationid.Value = Guid.NewGuid();

                            //HouseHold ID
                            // objTranRecom.ssi_householdid = new Lookup();
                            // objTranRecom.ssi_householdid.type = EntityName.account.ToString();
                            // objTranRecom.ssi_householdid.Value = new Guid(vHouseHold);

                            objTranRecom["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(vHouseHold));

                            //legal Entity ID
                            // objTranRecom.ssi_legalentityid = new Lookup();
                            // objTranRecom.ssi_legalentityid.type = EntityName.ssi_legalentity.ToString();
                            // objTranRecom.ssi_legalentityid.Value = new Guid(vLegalEntityId);

                            objTranRecom["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(vLegalEntityId));

                            //Fund ID
                            // objTranRecom.ssi_fundid = new Lookup();
                            // objTranRecom.ssi_fundid.type = EntityName.ssi_fund.ToString();
                            // objTranRecom.ssi_fundid.Value = new Guid(vFundId);

                            objTranRecom["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(vFundId));

                            //Account ID
                            if (vAccountId != "")
                            {
                                //objTranRecom.ssi_accountid = new Lookup();
                                //objTranRecom.ssi_accountid.type = EntityName.ssi_account.ToString();
                                //objTranRecom.ssi_accountid.Value = new Guid(vAccountId);

                                objTranRecom["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(vAccountId));

                            }
                            else
                            {
                                // objTranRecom.ssi_accountid = new Lookup();
                                // objTranRecom.ssi_accountid.IsNull = true;
                                // objTranRecom.ssi_accountid.IsNullSpecified = true;

                                objTranRecom["ssi_accountid"] = null;
                            }

                            // TransactionType
                            if (vTransactionTypes != "")
                            {
                                // objTranRecom.ssi_transactiontypes = new Picklist();
                                // objTranRecom.ssi_transactiontypes.Value = Convert.ToInt32(vTransactionTypes);

                                objTranRecom["ssi_transactiontypes"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vTransactionTypes));

                            }

                            //Close date investment
                            //if (vCloseDateInvestment != "")
                            //{
                            // objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            // objTranRecom.ssi_closedateinvestment.Value = DateTime.Parse(txtDate.Text).ToString();

                            objTranRecom["ssi_closedateinvestment"] = Convert.ToDateTime(txtDate.Text);


                            //}
                            //else
                            //{
                            //    objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            //    objTranRecom.ssi_closedateinvestment.IsNull = true;
                            //    objTranRecom.ssi_closedateinvestment.IsNullSpecified = true;
                            //}

                            //Praposed Amount
                            if (vproposedamount != "")
                            {
                                // objTranRecom.ssi_proposedamount = new CrmMoney();
                                // objTranRecom.ssi_proposedamount.Value = Convert.ToDecimal(vproposedamount);

                                objTranRecom["ssi_proposedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vproposedamount));
                            }
                            else
                            {
                                // objTranRecom.ssi_proposedamount = new CrmMoney();
                                // objTranRecom.ssi_proposedamount.IsNull = true;
                                // objTranRecom.ssi_proposedamount.IsNullSpecified = true;

                                objTranRecom["ssi_proposedamount"] = null;
                            }

                            ////Conformed Amount
                            if (vconfirmedamount != "")
                            {
                                // objTranRecom.ssi_confirmedamount = new CrmMoney();
                                // objTranRecom.ssi_confirmedamount.Value = Convert.ToDecimal(vconfirmedamount);

                                objTranRecom["ssi_confirmedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vconfirmedamount));
                            }
                            else
                            {
                                //  objTranRecom.ssi_confirmedamount = new CrmMoney();
                                //  objTranRecom.ssi_confirmedamount.IsNull = true;
                                //  objTranRecom.ssi_confirmedamount.IsNullSpecified = true;

                                objTranRecom["ssi_confirmedamount"] = null;
                            }

                            ////class B Fee
                            //if (vClassbfee != "")
                            //{

                            //    objTranRecom.ssi_classbfee = new CrmMoney();
                            //    objTranRecom.ssi_classbfee.Value = Convert.ToDecimal(vClassbfee);
                            //}
                            //else
                            //{
                            //    objTranRecom.ssi_classbfee = new CrmMoney();
                            //    objTranRecom.ssi_classbfee.IsNull = true;
                            //    objTranRecom.ssi_classbfee.IsNullSpecified = true;
                            //}

                            //Status 
                            if (vStatus != "")
                            {
                                // objTranRecom.ssi_status = new Picklist();
                                // objTranRecom.ssi_status.Value = Convert.ToInt16(vStatus);

                                objTranRecom["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt16(vStatus));
                            }
                            //Status date
                            if (vStatusDate != "")
                            {
                                // objTranRecom.ssi_statusdate = new CrmDateTime();
                                // objTranRecom.ssi_statusdate.Value = DateTime.Parse(vStatusDate).ToString();

                                objTranRecom["ssi_statusdate"] = Convert.ToDateTime(vStatusDate);
                            }
                            //Status date
                            //if (StatusTime != "")
                            //{
                            //    objTranRecom.ssi_s = new CrmDateTime();
                            //    objTranRecom.ssi_statusdate.Value = DateTime.Parse(StatusTime).ToString();
                            //}
                            //Response.Write(vStatusDate);

                            //grasham Adviced checks
                            if (GPAdvoice == "1")
                            {
                                // objTranRecom.ssi_greshamadvised = new CrmBoolean();
                                // objTranRecom.ssi_greshamadvised.Value = true;

                                objTranRecom["ssi_greshamadvised"] = true;
                            }
                            // service.Timeout = 1000000;
                            service.Create(objTranRecom);

                            iTotalCount++;
                        }
                        catch (Exception ex)
                        {
                            lblError.Text = "Error in insert reco=" + ex.ToString();
                            lblError.ForeColor = System.Drawing.Color.Green;
                            Response.Write("Error in insert reco=" + ex.ToString());
                        }
                    }

                    count[1] = iTotalCount;
                }
                else
                {
                    lblError.Text = "empty Row";
                    lblError.ForeColor = System.Drawing.Color.Green;
                }

            }
            else
            {
                lblError.Text = "Empty Dataset";
                lblError.ForeColor = System.Drawing.Color.Green;

            }
            //dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat for GDGS Data

            //int GDGScount = InsertDistibuteRecoGDGS(dsData);
            //count.Add(iTotalCount);
            //count.Add(GDGScount);

            return count;
        }
        catch (Exception e)
        {
            lblError.Text = "Capital Call Insert Fail Please Try Again-" + e.ToString();
            lblError.ForeColor = System.Drawing.Color.Green;
            Response.Write(e.Message);
            return count;

        }
    }


    #region not used
    //protected int InsertDistibuteRecoGDGS(DataSet dsData)
    //{
    //    int iTotalCount = 0;
    //    string crmServerUrl = "http://crm-test3/";
    //    //string crmServerURL = "http://server:5555/";

    //    string orgName = "GreshamPartners";
    //    //string orgName = "Webdev";
    //    bool bProceed = false;
    //    //try
    //    //{
    //    string UserId = GetcurrentUser();
    // //   CrmService service = GetCrmService(crmServerUrl, orgName, UserId); // create Connect

    //    //  DataSet dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat



    //    if (dsData.Tables.Count > 0)
    //    {
    //        if (dsData.Tables[1].Rows.Count > 0)
    //        {
    //            foreach (DataRow row in dsData.Tables[1].Rows)
    //            {
    //                string vAnizianoid = row["ssi_anzianoid"].ToString();
    //                string vHouseHold = row["Ssi_HouseholdId"].ToString();
    //                string vLegalEntityId = row["Ssi_LegalEntityId"].ToString();
    //                string vFundId = row["Ssi_FundId"].ToString();
    //                string vAccountId = row["Ssi_AccountId"].ToString();
    //                string vTransactionTypes = row["ssi_transactiontypes"].ToString();
    //                string vgreshamAdvised = row["ssi_greshamAdvised"].ToString();
    //                string vWithdrewalType = row["ssi_withdrawalType"].ToString();

    //                string vCloseDateInvestment = row["Ssi_CloseDateInvestment"].ToString();
    //                string vproposedamount = row["ssi_proposedamount"].ToString();
    //                string vconfirmedamount = row["ssi_confirmedamount"].ToString();

    //                string vClassbfee = row["ssi_classbfee"].ToString();

    //                string vStatus = row["Ssi_Status"].ToString();
    //                string vStatusDate = row["Ssi_StatusDate"].ToString();
    //                string vNotes = row["Ssi_Notes"].ToString();

    //                try
    //                {


    //                    ssi_transactionrecommendation objTranRecom = new ssi_transactionrecommendation();
    //                    //transaction reco ID
    //                    objTranRecom.ssi_transactionrecommendationid = new Key();
    //                    objTranRecom.ssi_transactionrecommendationid.Value = Guid.NewGuid();

    //                    //HouseHold ID
    //                    objTranRecom.ssi_householdid = new Lookup();
    //                    objTranRecom.ssi_householdid.type = EntityName.account.ToString();
    //                    objTranRecom.ssi_householdid.Value = new Guid(vHouseHold);

    //                    //legal Entity ID
    //                    objTranRecom.ssi_legalentityid = new Lookup();
    //                    objTranRecom.ssi_legalentityid.type = EntityName.ssi_legalentity.ToString();
    //                    objTranRecom.ssi_legalentityid.Value = new Guid(vLegalEntityId);

    //                    //Fund ID
    //                    objTranRecom.ssi_fundid = new Lookup();
    //                    objTranRecom.ssi_fundid.type = EntityName.ssi_fund.ToString();
    //                    objTranRecom.ssi_fundid.Value = new Guid(vFundId);

    //                    //Account ID
    //                    if (vAccountId != "")
    //                    {
    //                        objTranRecom.ssi_accountid = new Lookup();
    //                        objTranRecom.ssi_accountid.type = EntityName.ssi_account.ToString();
    //                        objTranRecom.ssi_accountid.Value = new Guid(vAccountId);
    //                    }
    //                    else
    //                    {
    //                        objTranRecom.ssi_accountid = new Lookup();
    //                        objTranRecom.ssi_accountid.IsNull = true;
    //                        objTranRecom.ssi_accountid.IsNullSpecified = true;
    //                    }

    //                    //TransactionType
    //                    objTranRecom.ssi_transactiontypes = new Picklist();
    //                    objTranRecom.ssi_transactiontypes.Value = Convert.ToInt32(vTransactionTypes);

    //                    //grasham Adviced checks
    //                    objTranRecom.ssi_greshamadvised = new CrmBoolean();
    //                    objTranRecom.ssi_greshamadvised.Value = true;


    //                    //Close date investment
    //                    if (vCloseDateInvestment != "")
    //                    {
    //                        objTranRecom.ssi_closedateinvestment = new CrmDateTime();
    //                        objTranRecom.ssi_closedateinvestment.Value = DateTime.Parse(vCloseDateInvestment).ToString();
    //                    }
    //                    else
    //                    {
    //                        objTranRecom.ssi_closedateinvestment = new CrmDateTime();
    //                        objTranRecom.ssi_closedateinvestment.IsNull = true;
    //                        objTranRecom.ssi_closedateinvestment.IsNullSpecified = true;
    //                    }

    //                    //Praposed Amount
    //                    if (vproposedamount != "")
    //                    {
    //                        objTranRecom.ssi_proposedamount = new CrmMoney();
    //                        objTranRecom.ssi_proposedamount.Value = Convert.ToDecimal(vproposedamount);
    //                    }
    //                    else
    //                    {
    //                        objTranRecom.ssi_proposedamount = new CrmMoney();
    //                        objTranRecom.ssi_proposedamount.IsNull = true;
    //                        objTranRecom.ssi_proposedamount.IsNullSpecified = true;
    //                    }

    //                    //Conformed Amount
    //                    if (vconfirmedamount != "")
    //                    {
    //                        objTranRecom.ssi_confirmedamount = new CrmMoney();
    //                        objTranRecom.ssi_confirmedamount.Value = Convert.ToDecimal(vconfirmedamount);
    //                    }
    //                    else
    //                    {
    //                        objTranRecom.ssi_confirmedamount = new CrmMoney();
    //                        objTranRecom.ssi_confirmedamount.IsNull = true;
    //                        objTranRecom.ssi_confirmedamount.IsNullSpecified = true;
    //                    }

    //                    //class B Fee
    //                    if (vClassbfee != "")
    //                    {
    //                        objTranRecom.ssi_classbfee = new CrmMoney();
    //                        objTranRecom.ssi_classbfee.Value = Convert.ToDecimal(vClassbfee);

    //                    }
    //                    else
    //                    {
    //                        objTranRecom.ssi_classbfee = new CrmMoney();
    //                        objTranRecom.ssi_classbfee.IsNull = true;
    //                        objTranRecom.ssi_classbfee.IsNullSpecified = true;
    //                    }

    //                    //Status 
    //                    objTranRecom.ssi_status = new Picklist();
    //                    objTranRecom.ssi_status.Value = Convert.ToInt16(vStatus);

    //                    //Status date
    //                    objTranRecom.ssi_statusdate = new CrmDateTime();
    //                    objTranRecom.ssi_statusdate.Value = DateTime.Parse(vStatusDate).ToString();
    //                    //Response.Write(vStatusDate);


    //                    // withdrewal type 
    //                    objTranRecom.ssi_withdrawaltype = new Picklist();
    //                    objTranRecom.ssi_withdrawaltype.Value = Convert.ToInt16(vWithdrewalType);

    //                    objTranRecom.ssi_notes = vNotes;

    //                    service.Create(objTranRecom);

    //                    iTotalCount++;
    //                }
    //                catch (Exception ex)
    //                {
    //                    lblError.Text = "Error in insert reco=" + ex.ToString();
    //                    lblError.ForeColor = System.Drawing.Color.Green;
    //                }
    //            }
    //        }
    //        //else
    //        //{
    //        //    lblError.Text = "empty Row";
    //        //    lblError.ForeColor = System.Drawing.Color.Red;
    //        //}

    //    }
    //    //else
    //    //{
    //    //    lblError.Text = "Empty DS";
    //    //    lblError.ForeColor = System.Drawing.Color.Red;

    //    //}


    //    return iTotalCount;
    //}
    #endregion

    public DataSet GetData(string vFundName, string vStartDate)
    {
        DataSet ds = new DataSet();
        try
        {

            SqlConnection conn = null;
            string vQuery = "EXEC SP_S_DISTRIBUTION_RECOMMENDATION @Fund_Name ='" + vFundName + "',@Dist_Date = '" + vStartDate + "'";
            try
            {
                string ConnString = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
                //   conn = new SqlConnection(AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring));
                conn = new SqlConnection(ConnString);

                SqlDataAdapter da = new SqlDataAdapter(vQuery, conn);
                da.SelectCommand.CommandTimeout = 2400;

                conn.Open();
                da.Fill(ds);
                da.Dispose();

            }
            finally
            {
                if (conn != null)
                    conn.Close();
            }
        }
        catch
        {
            lblError.Text = "File read Fail";
            lblError.ForeColor = System.Drawing.Color.Green;
        }
        return ds;
    }
    public DataSet getDataSet(string vSqlQuery)
    {
        DataSet ds = new DataSet();
        try
        {
            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter dagersham = new SqlDataAdapter();
            DataSet ds_gresham = new DataSet();

            dagersham = new SqlDataAdapter(vSqlQuery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds);//Fill Dataset 
        }
        catch (Exception ex)
        {

            //  lblMessage.ForeColor = System.Drawing.Color.Red;
            //lblMessage.Text = "Error in getting dataset value" + ex.Message;
        }
        return ds;

    }

    public void InsertCapitalCallGDGS()
    {
        #region CRM Connection
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        // CrmService service = null;
        IOrganizationService service = null;

        try
        {
            string UserId = GetcurrentUser();

            service = clsGM.GetCrmService();
            ///  strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            //  bProceed = false;
            //strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            //lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            //bProceed = false;
            //strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            //lblMessage.Text = strDescription;
        }

        #endregion

        DataSet dsData = getDataSet("SP_S_CAPITALCALLLETTER_GDGS @GDGSFlg = 1,@Asofdate='" + txtDate.Text + "'");
        int DeleteRecord = 0, insertRecord = 0;

        if (dsData.Tables[1].Rows.Count > 0)
        {


        }
        try
        {
            //  delete transactionrecommendation
            if (dsData.Tables.Count > 0)
            {
                if (dsData.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in dsData.Tables[1].Rows)
                    {

                        string status = row["status"].ToString();
                        string ssi_transactionrecommendationid = row["ssi_transactionrecommendationid"].ToString();

                        Guid gUId = new Guid(ssi_transactionrecommendationid);
                        // service.Timeout = 1000000;
                        service.Delete("ssi_transactionrecommendation", gUId);
                        DeleteRecord++;
                    }
                }
            }
            int iTotalCount = 0;


            //// insert GDGS 

            if (dsData.Tables[2].Rows.Count > 0)
            {
                foreach (DataRow row in dsData.Tables[2].Rows)
                {
                    string vHouseHold = row["AccountID"].ToString();
                    string vLegalEntityId = row["Ssi_LegalEntityId"].ToString();
                    string vFundId = row["ssi_fundid"].ToString();

                    string vAccountId = row["Ssi_AccountId"].ToString();
                    string vproposedamount = row["proposedamount"].ToString();
                    string vconfirmedamount = row["confirmedamount"].ToString();

                    string vTransactionTypes = row["ssi_transactiontypes"].ToString();
                    string vStatus = row["ssi_status"].ToString();
                    string vStatusDate = row["ssi_StatusDate"].ToString();
                    string GPAdvoice = row["ssi_GreshamAdvised"].ToString();
                    //  string StatusTime=row["statusTime"].ToString();
                    string Ssi_WithdrawalType = row["Ssi_WithdrawalType"].ToString();

                    string fund_shot_name = row["fund_short_name"].ToString();
                    string notes = "";
                    if (vconfirmedamount != "")
                    {
                        notes = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(vconfirmedamount));
                        notes = notes + ";" + fund_shot_name;
                    }
                    string ssi_closedateInvestment = row["ssi_closedateInvestment"].ToString();
                    try
                    {


                        // ssi_transactionrecommendation objTranRecom = new ssi_transactionrecommendation();
                        Microsoft.Xrm.Sdk.Entity objTranRecom = new Microsoft.Xrm.Sdk.Entity("ssi_transactionrecommendation");

                        //transaction reco ID
                        // objTranRecom.ssi_transactionrecommendationid = new Key();
                        // objTranRecom.ssi_transactionrecommendationid.Value = Guid.NewGuid();

                        //HouseHold ID
                        // objTranRecom.ssi_householdid = new Lookup();
                        // objTranRecom.ssi_householdid.type = EntityName.account.ToString();
                        // objTranRecom.ssi_householdid.Value = new Guid(vHouseHold);

                        objTranRecom["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(vHouseHold));

                        //legal Entity ID
                        // objTranRecom.ssi_legalentityid = new Lookup();
                        // objTranRecom.ssi_legalentityid.type = EntityName.ssi_legalentity.ToString();
                        // objTranRecom.ssi_legalentityid.Value = new Guid(vLegalEntityId);

                        objTranRecom["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(vLegalEntityId));

                        //Fund ID
                        // objTranRecom.ssi_fundid = new Lookup();
                        // objTranRecom.ssi_fundid.type = EntityName.ssi_fund.ToString();
                        // objTranRecom.ssi_fundid.Value = new Guid(vFundId);

                        objTranRecom["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(vFundId));

                        //Account ID
                        if (vAccountId != "")
                        {
                            // objTranRecom.ssi_accountid = new Lookup();
                            // objTranRecom.ssi_accountid.type = EntityName.ssi_account.ToString();
                            // objTranRecom.ssi_accountid.Value = new Guid(vAccountId);

                            objTranRecom["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(vAccountId));
                        }
                        else
                        {
                            // objTranRecom.ssi_accountid = new Lookup();
                            // objTranRecom.ssi_accountid.IsNull = true;
                            // objTranRecom.ssi_accountid.IsNullSpecified = true;

                            objTranRecom["ssi_accountid"] = null;
                        }

                        // TransactionType
                        if (vTransactionTypes != "")
                        {
                            // objTranRecom.ssi_transactiontypes = new Picklist();
                            // objTranRecom.ssi_transactiontypes.Value = Convert.ToInt32(vTransactionTypes);

                            objTranRecom["ssi_transactiontypes"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vTransactionTypes));
                        }

                        //Close date investment
                        if (ssi_closedateInvestment != "")
                        {
                            // objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            // objTranRecom.ssi_closedateinvestment.Value = DateTime.Parse(ssi_closedateInvestment).ToString();

                            objTranRecom["ssi_closedateinvestment"] = Convert.ToDateTime(ssi_closedateInvestment);
                        }
                        else
                        {
                            // objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            // objTranRecom.ssi_closedateinvestment.IsNull = true;
                            // objTranRecom.ssi_closedateinvestment.IsNullSpecified = true;

                            objTranRecom["ssi_closedateinvestment"] = null;
                        }

                        //Praposed Amount
                        if (vproposedamount != "")
                        {
                            //  objTranRecom.ssi_proposedamount = new CrmMoney();
                            //  objTranRecom.ssi_proposedamount.Value = Convert.ToDecimal(vproposedamount);

                            objTranRecom["ssi_proposedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vproposedamount));
                        }
                        else
                        {
                            // objTranRecom.ssi_proposedamount = new CrmMoney();
                            // objTranRecom.ssi_proposedamount.IsNull = true;
                            // objTranRecom.ssi_proposedamount.IsNullSpecified = true;

                            objTranRecom["ssi_proposedamount"] = null;
                        }

                        ////Conformed Amount
                        if (vconfirmedamount != "")
                        {
                            // objTranRecom.ssi_confirmedamount = new CrmMoney();
                            // objTranRecom.ssi_confirmedamount.Value = Convert.ToDecimal(vconfirmedamount);

                            objTranRecom["ssi_confirmedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vconfirmedamount));
                        }
                        else
                        {
                            // objTranRecom.ssi_confirmedamount = new CrmMoney();
                            // objTranRecom.ssi_confirmedamount.IsNull = true;
                            // objTranRecom.ssi_confirmedamount.IsNullSpecified = true;

                            objTranRecom["ssi_confirmedamount"] = null;
                        }


                        //Status 
                        if (vStatus != "")
                        {
                            // objTranRecom.ssi_status = new Picklist();
                            // objTranRecom.ssi_status.Value = Convert.ToInt16(vStatus);

                            objTranRecom["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vStatus));

                        }
                        //Status date
                        if (vStatusDate != "")
                        {
                            // objTranRecom.ssi_statusdate = new CrmDateTime();
                            // objTranRecom.ssi_statusdate.Value = DateTime.Parse(vStatusDate).ToString();

                            objTranRecom["ssi_statusdate"] = Convert.ToDateTime(vStatusDate);
                        }
                        //Status date
                        //if (StatusTime != "")
                        //{
                        //    objTranRecom.ssi_s = new CrmDateTime();
                        //    objTranRecom.ssi_statusdate.Value = DateTime.Parse(StatusTime).ToString();
                        //}
                        //Response.Write(vStatusDate);

                        //grasham Adviced checks
                        if (GPAdvoice == "1")
                        {
                            //objTranRecom.ssi_greshamadvised = new CrmBoolean();
                            //objTranRecom.ssi_greshamadvised.Value = true;

                            objTranRecom["ssi_greshamadvised"] = Convert.ToBoolean(true);
                        }
                        else
                        {
                            // objTranRecom.ssi_greshamadvised = new CrmBoolean();
                            // objTranRecom.ssi_greshamadvised.Value = false;

                            objTranRecom["ssi_greshamadvised"] = false;

                        }

                        //   Ssi_WithdrawalType
                        if (Ssi_WithdrawalType != "")
                        {
                            // objTranRecom.ssi_withdrawaltype = new Picklist();
                            // objTranRecom.ssi_withdrawaltype.Value = Convert.ToInt32(Ssi_WithdrawalType);


                            objTranRecom["ssi_withdrawaltype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(Ssi_WithdrawalType));
                        }
                        if (notes != "")
                        {
                            //   objTranRecom.ssi_notes = new string();
                            objTranRecom["ssi_notes"] = notes;
                        }
                        //service.Timeout = 1000000;
                        service.Create(objTranRecom);

                        iTotalCount++;
                    }
                    catch (Exception ex)
                    {
                        lblError.Text = "Error in insert reco=" + ex.ToString();
                        lblError.ForeColor = System.Drawing.Color.Green;
                    }
                }

            }

            //string data = "Delete Record -" + DeleteRecord + "<br />Insert Record -" + iTotalCount;

            string data = "<br/> Fund : GDGS" + "  Record Insert= " + iTotalCount + "  Record Delete= " + DeleteRecord;

            string msg1 = Convert.ToString(ViewState["msg"]);

            string error = Convert.ToString(ViewState["error"]);


            string finalresult = "File Uploaded Sucessfully" + "<br>" + msg1 + "<br/>" + data;

            //Response.Write(error);

            //Response.Write(data);

            lblError.Text = finalresult;
            lblError.ForeColor = System.Drawing.Color.Green;

            lblError1.Text = error;

            //  Button2.Visible = false;
        }
        catch (Exception e)
        {
            lblError.Text = "Capital Call GDGS Insert Fail" + e.ToString();
            lblError.ForeColor = System.Drawing.Color.Green;
        }
        //   Button2.Visible = false;




    }
    public void InsertDistributionGDGS()
    {
        #region CRM Connection
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
                                                                                   //string crmServerURL = "http://server:5555/";
                                                                                   // string orgName = "GreshamPartners";
                                                                                   //string orgName = "Webdev";
                                                                                   //  CrmService service = null;
        IOrganizationService service = null;

        try
        {
            string UserId = GetcurrentUser();

            service = clsGM.GetCrmService();
            ///  strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            //  bProceed = false;
            //strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            //lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            //bProceed = false;
            //strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            //lblMessage.Text = strDescription;
        }

        #endregion

        DataSet dsData = getDataSet("SP_S_DISTRIBUTION_GDGS @GDGSFlg = 1,@Asofdate='" + txtDate.Text + "'");
        int DeleteRecord = 0, insertRecord = 0;
        try
        {
            //  delete transactionrecommendation
            if (dsData.Tables.Count > 0)
            {
                if (dsData.Tables[1].Rows.Count > 0)
                {
                    foreach (DataRow row in dsData.Tables[1].Rows)
                    {

                        string status = row["status"].ToString();
                        string ssi_transactionrecommendationid = row["ssi_transactionrecommendationid"].ToString();

                        Guid gUId = new Guid(ssi_transactionrecommendationid);
                        // service.Timeout = 1000000;
                        service.Delete("ssi_transactionrecommendation", gUId);
                        DeleteRecord++;
                    }
                }
            }
            int iTotalCount = 0;


            //// insert GDGS 

            if (dsData.Tables[2].Rows.Count > 0)
            {
                foreach (DataRow row in dsData.Tables[2].Rows)
                {
                    string vHouseHold = row["AccountID"].ToString();
                    string vLegalEntityId = row["Ssi_LegalEntityId"].ToString();
                    string vFundId = row["ssi_fundid"].ToString();

                    string vAccountId = row["Ssi_AccountId"].ToString();
                    string vproposedamount = row["proposedamount"].ToString();
                    string vconfirmedamount = row["confirmedamount"].ToString();
                    string vClassbfee = row["ssi_classbfee"].ToString();

                    string vTransactionTypes = row["ssi_transactiontypes"].ToString();
                    string vStatus = row["ssi_status"].ToString();
                    string Ssi_WithdrawalType = row["Ssi_WithdrawalType"].ToString();

                    string vStatusDate = row["ssi_StatusDate"].ToString();
                    string GPAdvoice = row["ssi_GreshamAdvised"].ToString();
                    //  string StatusTime=row["statusTime"].ToString();
                    string ssi_closedateInvestment = row["ssi_closedateInvestment"].ToString();

                    string ssi_GreshamAdvised = row["ssi_GreshamAdvised"].ToString();
                    string fund_shot_name = row["fund_short_name"].ToString();


                    string notes = "";
                    if (vconfirmedamount != "")
                    {
                        notes = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C2}", Convert.ToDecimal(vconfirmedamount));
                        notes = notes + ";" + fund_shot_name;
                    }

                    try
                    {

                        Microsoft.Xrm.Sdk.Entity objTranRecom = new Microsoft.Xrm.Sdk.Entity("ssi_transactionrecommendation");

                        //transaction reco ID
                        // objTranRecom.ssi_transactionrecommendationid = new Key();
                        // objTranRecom.ssi_transactionrecommendationid.Value = Guid.NewGuid();

                        //HouseHold ID
                        // objTranRecom.ssi_householdid = new Lookup();
                        // objTranRecom.ssi_householdid.type = EntityName.account.ToString();
                        // objTranRecom.ssi_householdid.Value = new Guid(vHouseHold);

                        objTranRecom["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(vHouseHold));

                        //legal Entity ID
                        // objTranRecom.ssi_legalentityid = new Lookup();
                        // objTranRecom.ssi_legalentityid.type = EntityName.ssi_legalentity.ToString();
                        // objTranRecom.ssi_legalentityid.Value = new Guid(vLegalEntityId);

                        objTranRecom["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(vLegalEntityId));

                        //Fund ID
                        // objTranRecom.ssi_fundid = new Lookup();
                        // objTranRecom.ssi_fundid.type = EntityName.ssi_fund.ToString();
                        // objTranRecom.ssi_fundid.Value = new Guid(vFundId);

                        objTranRecom["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(vFundId));

                        //Account ID
                        if (vAccountId != "")
                        {
                            // objTranRecom.ssi_accountid = new Lookup();
                            // objTranRecom.ssi_accountid.type = EntityName.ssi_account.ToString();
                            // objTranRecom.ssi_accountid.Value = new Guid(vAccountId);

                            objTranRecom["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(vAccountId));
                        }
                        else
                        {
                            // objTranRecom.ssi_accountid = new Lookup();
                            // objTranRecom.ssi_accountid.IsNull = true;
                            // objTranRecom.ssi_accountid.IsNullSpecified = true;

                            objTranRecom["ssi_accountid"] = null;
                        }

                        // TransactionType
                        if (vTransactionTypes != "")
                        {
                            // objTranRecom.ssi_transactiontypes = new Picklist();
                            // objTranRecom.ssi_transactiontypes.Value = Convert.ToInt32(vTransactionTypes);

                            objTranRecom["ssi_transactiontypes"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vTransactionTypes));
                        }

                        //Close date investment
                        if (ssi_closedateInvestment != "")
                        {
                            // objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            // objTranRecom.ssi_closedateinvestment.Value = DateTime.Parse(ssi_closedateInvestment).ToString();

                            objTranRecom["ssi_closedateinvestment"] = Convert.ToDateTime(ssi_closedateInvestment);
                        }
                        else
                        {
                            // objTranRecom.ssi_closedateinvestment = new CrmDateTime();
                            // objTranRecom.ssi_closedateinvestment.IsNull = true;
                            // objTranRecom.ssi_closedateinvestment.IsNullSpecified = true;

                            objTranRecom["ssi_closedateinvestment"] = null;
                        }

                        //Praposed Amount
                        if (vproposedamount != "")
                        {
                            //  objTranRecom.ssi_proposedamount = new CrmMoney();
                            //  objTranRecom.ssi_proposedamount.Value = Convert.ToDecimal(vproposedamount);

                            objTranRecom["ssi_proposedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vproposedamount));
                        }
                        else
                        {
                            // objTranRecom.ssi_proposedamount = new CrmMoney();
                            // objTranRecom.ssi_proposedamount.IsNull = true;
                            // objTranRecom.ssi_proposedamount.IsNullSpecified = true;

                            objTranRecom["ssi_proposedamount"] = null;
                        }

                        ////Conformed Amount
                        if (vconfirmedamount != "")
                        {
                            // objTranRecom.ssi_confirmedamount = new CrmMoney();
                            // objTranRecom.ssi_confirmedamount.Value = Convert.ToDecimal(vconfirmedamount);

                            objTranRecom["ssi_confirmedamount"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vconfirmedamount));
                        }
                        else
                        {
                            // objTranRecom.ssi_confirmedamount = new CrmMoney();
                            // objTranRecom.ssi_confirmedamount.IsNull = true;
                            // objTranRecom.ssi_confirmedamount.IsNullSpecified = true;

                            objTranRecom["ssi_confirmedamount"] = null;
                        }


                        //Status 
                        if (vStatus != "")
                        {
                            // objTranRecom.ssi_status = new Picklist();
                            // objTranRecom.ssi_status.Value = Convert.ToInt16(vStatus);

                            objTranRecom["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(vStatus));

                        }
                        //Status date
                        if (vStatusDate != "")
                        {
                            // objTranRecom.ssi_statusdate = new CrmDateTime();
                            // objTranRecom.ssi_statusdate.Value = DateTime.Parse(vStatusDate).ToString();

                            objTranRecom["ssi_statusdate"] = Convert.ToDateTime(vStatusDate);
                        }
                        //Status date
                        //if (StatusTime != "")
                        //{
                        //    objTranRecom.ssi_s = new CrmDateTime();
                        //    objTranRecom.ssi_statusdate.Value = DateTime.Parse(StatusTime).ToString();
                        //}
                        //Response.Write(vStatusDate);

                        //grasham Adviced checks
                        if (GPAdvoice == "1")
                        {
                            //objTranRecom.ssi_greshamadvised = new CrmBoolean();
                            //objTranRecom.ssi_greshamadvised.Value = true;

                            objTranRecom["ssi_greshamadvised"] = Convert.ToBoolean(true);
                        }
                        else
                        {
                            // objTranRecom.ssi_greshamadvised = new CrmBoolean();
                            // objTranRecom.ssi_greshamadvised.Value = false;

                            objTranRecom["ssi_greshamadvised"] = false;

                        }

                        //   Ssi_WithdrawalType
                        if (Ssi_WithdrawalType != "")
                        {
                            // objTranRecom.ssi_withdrawaltype = new Picklist();
                            // objTranRecom.ssi_withdrawaltype.Value = Convert.ToInt32(Ssi_WithdrawalType);


                            objTranRecom["ssi_withdrawaltype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(Ssi_WithdrawalType));
                        }

                        if (vClassbfee != "")
                        {
                            // objTranRecom.ssi_classbfee = new CrmMoney();
                            // objTranRecom.ssi_classbfee.Value = Convert.ToDecimal(vClassbfee);

                            objTranRecom["ssi_classbfee"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(vClassbfee));

                        }
                        else
                        {
                            // objTranRecom.ssi_classbfee = new CrmMoney();
                            // objTranRecom.ssi_classbfee.IsNull = true;
                            // objTranRecom.ssi_classbfee.IsNullSpecified = true;

                            objTranRecom["ssi_classbfee"] = null;
                        }


                        if (notes != "")
                        {
                            //   objTranRecom.ssi_notes = new string();
                            objTranRecom["ssi_notes"] = notes;
                        }
                        // service.Timeout = 1000000;
                        service.Create(objTranRecom);

                        iTotalCount++;
                    }
                    catch (Exception ex)
                    {
                        lblError.Text = "Error in insert reco=" + ex.ToString();
                        lblError.ForeColor = System.Drawing.Color.Green;
                    }
                }

            }


            string data = @"<br/> Fund : GDGS" + "  Record Insert= " + iTotalCount + "  Record Delete= " + DeleteRecord;

            //string data = "<br/> Fund : GDGS" + "  Record Insert= " + iTotalCount + "  Record Delete= " + DeleteRecord;

            string msg1 = Convert.ToString(ViewState["msg"]);

            string error = Convert.ToString(ViewState["error"]);


            string finalresult = "File uploaded Successfully" + msg1 + "<br/>" + data;

            //Response.Write(error);

            //Response.Write(data);

            lblError.Text = finalresult;
            lblError.ForeColor = System.Drawing.Color.Green;

            lblError1.Text = error;

            lblError.ForeColor = System.Drawing.Color.Green;

            //   Button2.Visible = false;
        }
        catch (Exception e)
        {
            lblError.Text = "Distribution GDGS Insert Fail" + e.ToString();
            lblError.ForeColor = System.Drawing.Color.Green;
        }


        //  Button2.Visible = false;


    }

    //public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    // string username = WindowsIdentity.GetCurrent().Name;

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

    //    //////////////////////////// impersonate service to crm user /////////////////////////////

    //    // WhoAmIRequest userRequest = new WhoAmIRequest();
    //    // Execute the request.
    //    // WhoAmIResponse user = (WhoAmIResponse)service.Execute(userRequest);
    //    // string currentuser = user.UserId.ToString();


    //    //string currentuser = "62DE1F95-8203-DE11-A38C-001D09665E8F";
    //    //token.CallerId = new Guid(currentuser);

    //    return service;
    //}


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

        Response.Write("claimsIdentity.Name : "+ claimsIdentity.Name);
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

    public int ProcessZip(string TempPath, string dateTime)
    {
        int count = 0;
        try
        {
            List<string> folderlist = new List<string>();
            if (fuDist.HasFile == true)
            {

                if (System.IO.Path.GetExtension(fuDist.FileName) == ".zip")
                {
                    //  int count = 0;


                    string ZipFileNamewithExtension = fuDist.FileName;
                    string ZipFileNamewithoutExtension = Path.GetFileNameWithoutExtension(fuDist.FileName);
                    string BackupFileName = ZipFileNamewithoutExtension + dateTime + ".zip";

                    string ZipBackupPath = Server.MapPath("") + @"\ExcelTemplate\" + BackupFileName;
                    string TempFilePath = TempPath + ZipFileNamewithExtension;

                    if (!Directory.Exists(TempPath))
                    {
                        Directory.CreateDirectory(TempPath);
                    }

                    //Backup Zip File                           

                    fuDist.PostedFile.SaveAs(ZipBackupPath);// Copy File To Backup Folder.
                    fuDist.PostedFile.SaveAs(TempPath + ZipFileNamewithExtension);//Copy File To Temp Folder For Processing

                    //  ZipFile.ExtractToDirectory(ZipBackupPath, TempPath);
                    ZipFile.ExtractToDirectory(ZipBackupPath, TempPath);
                    string DirectoryName1 = Path.GetFileNameWithoutExtension(TempPath + ZipFileNamewithExtension);

                    string[] Folderindirectory = Directory.GetDirectories(TempPath);
                    foreach (string subdir in Folderindirectory)
                    {

                        string[] filesindirectory = Directory.GetFiles(subdir);
                        foreach (string FileinFolder in filesindirectory)
                        {
                            string Fileextension = Path.GetExtension(FileinFolder);//Extension of the File Inside The Folder of The uploaded Zip.
                            if (Fileextension == ".xlsx")
                            {
                                string DirectoryName = Path.GetFileNameWithoutExtension(subdir); // Name of The Fund - FolderName inside the Zip.
                                File.Copy(FileinFolder, TempPath + DirectoryName + Fileextension, true); // Copy files in the TempFolder
                                lstFile.Add(TempPath + DirectoryName + Fileextension, DirectoryName);
                                count++;
                            }
                        }

                    }
                }
                else
                {
                    lblError.Text = "Please Upload Zip File";
                }
            }


            ViewState["lstFile"] = lstFile;
        }
        catch (Exception ex)
        {
            lblError.Text = "ERROR : " + ex.Message.ToString();
            bProceed = false;

        }
        return count;
    }

    public void Proces_ALL_FUND(string TempPath, IOrganizationService service, bool checkcapitalcall, bool checkdistribution)
    {
        clDB = new DB();
        //foreach (System.Web.UI.WebControls.ListItem li in lstFund.Items)
        string user = AppLogic.GetParam(AppLogic.ConfigParam.UserName).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();

        foreach (KeyValuePair<string, string> pair in lstFile)
        {
            int retVal = 1;
            // string FundShortName = string.Empty;
            string FilePath = string.Empty;

            string SelectedFund = string.Empty;
            string SelectedFundValue = string.Empty;
            string FundShortName = string.Empty;
            string strFileName = string.Empty;
            string SheetName = string.Empty;
            string TypeId = string.Empty;
            bool bProcess = false;

            if (rbCapitalCall.Checked)
            {
                strFileName = clDB.FileName("DistToRecoCapitalcall");
                SheetName = "Capital Call";
                TypeId = "13";
            }
            else if (rbDistribution.Checked)//Fund Distribution Letter - Fund Distribution  
            {
                strFileName = clDB.FileName("DistToRecoDistribution");
                SheetName = "Distribution";
                TypeId = "12";
            }



            FilePath = pair.Key.ToString(); // FilePath of Each File in Zip
            FundShortName = pair.Value.ToString(); // fundShortName of each File Associated to


            string Ssi_ShortName = string.Empty;
            DataTable dtFund = (DataTable)ViewState["dtFund"];
            for (int i = 0; i < dtFund.Rows.Count; i++)
            {
                Ssi_ShortName = Convert.ToString(dtFund.Rows[i]["FundShortName"]);
                SelectedFund = Convert.ToString(dtFund.Rows[i]["FundName"]);
                SelectedFundValue = Convert.ToString(dtFund.Rows[i]["FundId"]);
                //if (SelectedFund == ssi_name)
                //{
                //    break;
                //}
                if (Ssi_ShortName == FundShortName)
                {
                    bProcess = true;
                    break;
                }
                else
                {
                    bProcess = false;
                }
            }
            try
            {


                if (bProcess)
                {
                    bool bCopy = ReadAlPsFile(FilePath, SheetName, TempPath + strFileName);//CleanUp the File from Alps
                    if (bCopy)
                    {

                        string FileType = Convert.ToString(ViewState["FileType"]);

                        if (checkcapitalcall == true && FileType.ToLower() == "capital call")
                        {

                            using (new Impersonation("corp", user, Pass))
                            {
                                if (File.Exists(DTSFilePath + strFileName))
                                {
                                    File.Delete(DTSFilePath + strFileName); // Delete File from DTS PAth

                                    File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                                }
                                else
                                {
                                    File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                                }
                            }

                            #region JOBCALL
                            try
                            {

                                SqlConnection Gresham_con = new SqlConnection(con);

                                SqlCommand cmd = new SqlCommand();
                                SqlDataAdapter dagersham = new SqlDataAdapter();

                                DataSet ds_gresham = new DataSet();
                                DataSet ds = new DataSet();

                                Gresham_con = new SqlConnection(con);
                                Gresham_con.Open();
                                string greshamquery = "SP_S_ExecuteJobs @TypeId = " + TypeId;
                                cmd = new SqlCommand();
                                cmd.Connection = Gresham_con;
                                cmd.CommandText = greshamquery;

                                cmd.ExecuteNonQuery();
                                retVal = 0;
                                jobsucesscnt++;
                            }
                            catch (Exception exception3)
                            {
                                retVal = 1;
                                //  lblError.Text = lblError.Text + "<br/>Error Occurred in File for Fund :" + SelectedFund + exception3.ToString();// + "<br/>" + exception3.Message.ToString();

                                ErrorOccured = ErrorOccured + "<br/>File upload fail for fund:" + SelectedFund;

                                //lblError1.Text = ErrorOccured;
                                bpopup = false;
                                jobfailcont++;
                            }

                            #endregion
                        }


                        else if (checkdistribution == true && FileType.ToLower() == "distribution")
                        {
                            using (new Impersonation("corp", user, Pass))
                            {
                                if (File.Exists(DTSFilePath + strFileName))
                                {
                                    File.Delete(DTSFilePath + strFileName); // Delete File from DTS PAth

                                    File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                                }
                                else
                                {
                                    File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                                }
                            }

                            #region JOBCALL
                            try
                            {

                                SqlConnection Gresham_con = new SqlConnection(con);

                                SqlCommand cmd = new SqlCommand();
                                SqlDataAdapter dagersham = new SqlDataAdapter();

                                DataSet ds_gresham = new DataSet();
                                DataSet ds = new DataSet();

                                Gresham_con = new SqlConnection(con);
                                Gresham_con.Open();
                                string greshamquery = "SP_S_ExecuteJobs @TypeId = " + TypeId;
                                cmd = new SqlCommand();
                                cmd.Connection = Gresham_con;
                                cmd.CommandText = greshamquery;

                                cmd.ExecuteNonQuery();
                                retVal = 0;
                                jobsucesscnt++;
                            }
                            catch (Exception exception3)
                            {
                                retVal = 1;
                                //  lblError.Text = lblError.Text + "<br/>Error Occurred in File for Fund :" + SelectedFund + exception3.ToString();// + "<br/>" + exception3.Message.ToString();

                                ErrorOccured = ErrorOccured + "<br/>File upload fail for fund:" + SelectedFund;

                                //lblError1.Text = ErrorOccured;
                                bpopup = false;
                                jobfailcont++;
                            }

                            #endregion
                        }

                        else
                        {
                            ErrorOccured = ErrorOccured + "<br>" + "File process failed please check the file for fund: " + SelectedFund;
                            jobfailcont++;
                        }

                        if (retVal == 0 && checkcapitalcall == true)
                        {

                            DataSet dsData = getDataSet("SP_S_CAPITALCALLLETTER_GDGS @GDGSFlg = 0, @AsOfDate='" + txtDate.Text + "',@FundId='" + SelectedFundValue + "'");
                            //DataSet dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat
                            //lblError1.Text = "";
                            if (dsData.Tables[1].Rows.Count > 0)
                            {
                                Button3.Visible = true;
                                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                Type tp = this.GetType();
                                sb.Append("\n<script type=text/javascript>\n");
                                sb.Append("var bt = window.document.getElementById('Button3');\n");
                                sb.Append(@"if(confirm('There are existing records for the close date and fund(s) selected. Do you want to delete and reload?')){");
                                sb.Append("\nwindow.document.getElementById('Hidden1').value='1';");
                                sb.Append("\n txtDateClear();");
                                sb.Append(("\n bt.click();\n"));
                                sb.Append("\n}");
                                sb.Append("else\n{");
                                sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
                                sb.Append("\n txtDateClear();");
                                sb.Append("\n}");
                                sb.Append("</script>");
                                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                            }
                            else
                            {
                                count = InsertCapitalCall(SelectedFundValue);
                                msg = msg + @"<br/> Fund Name :" + FundShortName + "  Record Insert= " + count[1].ToString() + "  Record Delete= " + count[0].ToString();
                                bpopup = true;
                                // lblError.Text = @"Capital Call File Upload Sucessfully <br /> Fund Name :" + FundShortName + @"<br/> Record Insert=" + count[1].ToString() + @"<br /> Record Delete=" + count[0].ToString();
                                lblError.ForeColor = System.Drawing.Color.Green;
                                countpopup++;
                            }

                        }

                        else if (retVal == 0 && checkdistribution == true)
                        {
                            DataSet dsData = getDataSet("SP_S_DISTRIBUTION_GDGS @GDGSFlg = 0, @AsOfDate='" + txtDate.Text + "',@FundId='" + SelectedFundValue + "'");
                            //DataSet dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat

                            if (dsData.Tables[1].Rows.Count > 0)
                            {
                                Button3.Visible = true;
                                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                Type tp = this.GetType();
                                sb.Append("\n<script type=text/javascript>\n");
                                sb.Append("var bt = window.document.getElementById('Button3');\n");
                                sb.Append(@"if(confirm('There are existing records for the close date and fund(s) selected. Do you want to delete and reload?')){");
                                sb.Append("\nwindow.document.getElementById('Hidden1').value='1';");
                                sb.Append("\n txtDateClear();");
                                sb.Append(("\n bt.click();\n"));
                                sb.Append("\n}");
                                sb.Append("else\n{");
                                sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
                                sb.Append("\n txtDateClear();");
                                sb.Append("\n}");
                                sb.Append("</script>");
                                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                            }
                            else
                            {
                                count = InsertDistibuteReco(SelectedFundValue);

                                msg = msg + @"<br /> Fund Name :" + FundShortName + "  Record Insert= " + count[1].ToString() + "  Record Delete= " + count[0].ToString();

                                // lblError.Text = "Distribution File Upload Sucessfully" + msg;

                                bpopup = true;
                                //lblError.Text = @"Distribution File Upload Sucessfully <br /> Fund Name :" + FundShortName + @"<br/> Record Insert=" + count[1].ToString() + @"<br /> Record Delete=" + count[0].ToString();
                                lblError.ForeColor = System.Drawing.Color.Green;
                                countpopup++;
                            }

                        }

                        //break;

                    }
                    else
                    {
                        //lblError1.Text = "Error Occured in File Process";

                        ErrorOccured = ErrorOccured + "Error Occured in File Process";

                        // lblError1.Text = ErrorOccured;

                        bpopup = false;
                    }
                }
                else
                {
                    // strErrorOccured = strErrorOccured + "<br/>Fund Not Found :" + FundShortName;

                    ErrorOccured = ErrorOccured + "<br/>Fund Not Found :" + FundShortName;
                    // lblError1.Text = ErrorOccured;

                    // lblError1.Text = lblError1.Text + "<br/>Fund Not Found :" + FundShortName;
                    bpopup = false;
                }


            }
            catch (System.Web.Services.Protocols.SoapException exc1)
            {

                Response.Write("<br/>Exception: " + exc1.Detail.InnerText);

            }
            catch (Exception exc)
            {
                Response.Write(exc.Message + exc.StackTrace);
            }
            finally
            {
                //if (sqlconn != null)
                //    if (sqlconn.State != System.Data.ConnectionState.Open)
                //        sqlconn.Close();
            }

        }
    }


    public void Proces_ALL_FUND_ReInsert(string TempPath, IOrganizationService service, bool checkcapitalcall, bool checkdistribution)
    {
        clDB = new DB();
        //foreach (System.Web.UI.WebControls.ListItem li in lstFund.Items)
        string user = AppLogic.GetParam(AppLogic.ConfigParam.UserName).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        Dictionary<string, string> lstFile1 = (Dictionary<String, String>)ViewState["lstFile"];
        foreach (KeyValuePair<string, string> pair in lstFile1)
        {
            int retVal = 1;
            // string FundShortName = string.Empty;
            string FilePath = string.Empty;

            string SelectedFund = string.Empty;
            string SelectedFundValue = string.Empty;
            string FundShortName = string.Empty;
            string strFileName = string.Empty;
            string SheetName = string.Empty;
            string TypeId = string.Empty;
            bool bProcess = false;

            if (rbCapitalCall.Checked)
            {
                strFileName = clDB.FileName("DistToRecoCapitalcall");
                SheetName = "Capital Call";
                TypeId = "13";
            }
            else if (rbDistribution.Checked)//Fund Distribution Letter - Fund Distribution  
            {
                strFileName = clDB.FileName("DistToRecoDistribution");
                SheetName = "Distribution";
                TypeId = "12";
            }



            FilePath = pair.Key.ToString(); // FilePath of Each File in Zip
            FundShortName = pair.Value.ToString(); // fundShortName of each File Associated to


            string Ssi_ShortName = string.Empty;
            DataTable dtFund = (DataTable)ViewState["dtFund"];
            for (int i = 0; i < dtFund.Rows.Count; i++)
            {
                Ssi_ShortName = Convert.ToString(dtFund.Rows[i]["FundShortName"]);
                SelectedFund = Convert.ToString(dtFund.Rows[i]["FundName"]);
                SelectedFundValue = Convert.ToString(dtFund.Rows[i]["FundId"]);
                //if (SelectedFund == ssi_name)
                //{
                //    break;
                //}
                if (Ssi_ShortName == FundShortName)
                {
                    bProcess = true;
                    break;
                }
                else
                {
                    bProcess = false;
                }
            }
            try
            {


                if (bProcess)
                {
                    bool bCopy = ReadAlPsFile(FilePath, SheetName, TempPath + strFileName);//CleanUp the File from Alps
                    if (bCopy)
                    {

                        using (new Impersonation("corp", user, Pass))
                        {
                            if (File.Exists(DTSFilePath + strFileName))
                            {
                                File.Delete(DTSFilePath + strFileName); // Delete File from DTS PAth

                                File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                            }
                            else
                            {
                                File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                            }
                        }

                        #region JOBCALL
                        try
                        {

                            SqlConnection Gresham_con = new SqlConnection(con);

                            SqlCommand cmd = new SqlCommand();
                            SqlDataAdapter dagersham = new SqlDataAdapter();

                            DataSet ds_gresham = new DataSet();
                            DataSet ds = new DataSet();

                            Gresham_con = new SqlConnection(con);
                            Gresham_con.Open();
                            string greshamquery = "SP_S_ExecuteJobs @TypeId = " + TypeId;
                            cmd = new SqlCommand();
                            cmd.Connection = Gresham_con;
                            cmd.CommandText = greshamquery;

                            cmd.ExecuteNonQuery();
                            retVal = 0;
                        }
                        catch (Exception exception3)
                        {
                            retVal = 1;
                            lblError.Text = lblError.Text + "<br/>Error Occurred in File for Fund :" + SelectedFund + exception3.ToString();// + "<br/>" + exception3.Message.ToString();
                            bpopup = false;
                        }

                        #endregion








                        if (retVal == 0 && checkcapitalcall == true)
                        {

                            //DataSet dsData = getDataSet("SP_S_CAPITALCALLLETTER_GDGS @GDGSFlg = 0, @AsOfDate='" + txtDate.Text + "',@FundId='" + SelectedFundValue + "'");
                            //DataSet dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat



                            count = InsertCapitalCall(SelectedFundValue);
                            //lblError.Text = @"Capital Call File Upload Sucessfully <br />Record  Insert=" + count[1].ToString() + @"<br /> Record Delete=" + count[0].ToString();

                            //msg = msg + @"<br/> Fund Name :" + FundShortName + @"<br/> Record Insert=" + count[1].ToString() + @"<br/> Record Delete=" + count[0].ToString();

                            msg = msg + @"<br/> Fund Name :" + FundShortName + "  Record Insert= " + count[1].ToString() + "  Record Delete= " + count[0].ToString();
                            // lblError.Text = "Capital Call File Upload Sucessfully " + msg;

                            //  ViewState["msg"] = msg;
                            bpopup = true;
                            // lblError.Text = @"Capital Call File Upload Sucessfully <br /> Fund Name :" + FundShortName + @"<br/> Record Insert=" + count[1].ToString() + @"<br /> Record Delete=" + count[0].ToString();
                            lblError.ForeColor = System.Drawing.Color.Green;

                            countpopup++;




                        }

                        else if (retVal == 0 && checkdistribution == true)
                        {
                            //DataSet dsData = getDataSet("SP_S_DISTRIBUTION_GDGS @GDGSFlg = 0, @AsOfDate='" + txtDate.Text + "',@FundId='" + SelectedFundValue + "'");
                            //DataSet dsData = GetData(vFundName, vStartDate); // get Excel data in DS Fromat


                            count = InsertDistibuteReco(SelectedFundValue);

                            msg = msg + @"<br /> Fund Name :" + FundShortName + "  Record Insert= " + count[1].ToString() + "  Record Delete= " + count[0].ToString();

                            // lblError.Text = "Distribution File Upload Sucessfully" + msg;

                            bpopup = true;
                            //lblError.Text = @"Distribution File Upload Sucessfully <br /> Fund Name :" + FundShortName + @"<br/> Record Insert=" + count[1].ToString() + @"<br /> Record Delete=" + count[0].ToString();
                            lblError.ForeColor = System.Drawing.Color.Green;

                            countpopup++;
                        }

                        //break;

                    }
                    else
                    {
                        // lblError1.Text = "Error Occured in File Process";

                        ErrorOccured = "Error Occured in File Process";

                        //  lblError1.Text = ErrorOccured;
                        bpopup = false;
                    }
                }
                else
                {
                    // strErrorOccured = strErrorOccured + "<br/>Fund Not Found :" + FundShortName;
                    ErrorOccured = ErrorOccured + "<br/>Fund Not Found :" + FundShortName;

                    // lblError1.Text = ErrorOccured;
                    //  lblError1.Text = lblError1.Text + "<br/>Fund Not Found :" + FundShortName;
                    bpopup = false;
                }


            }
            catch (System.Web.Services.Protocols.SoapException exc1)
            {

                Response.Write("<br/>Exception: " + exc1.Detail.InnerText);

            }
            catch (Exception exc)
            {
                Response.Write(exc.Message + exc.StackTrace);
            }
            finally
            {
                //if (sqlconn != null)
                //    if (sqlconn.State != System.Data.ConnectionState.Open)
                //        sqlconn.Close();
            }

        }
    }



    protected void txtDate_TextChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        lblError1.Text = "";

    }
    protected void rbCapitalCall_CheckedChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        lblError1.Text = "";
    }
    protected void rbDistribution_CheckedChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        lblError1.Text = "";
    }
    protected void Button1_Click(object sender, EventArgs e)
    {

        //System.Text.StringBuilder sb = new System.Text.StringBuilder();
        //Type tp = this.GetType();

        //sb.Append("\n<script type=text/javascript>\n");

        //sb.Append("var bt = window.document.getElementById('Button2');\n");
        //sb.Append("if(confirm('There are existing records for the close date and fund selected. Do you want to delete and reload?'))\n{");
        //sb.Append(("\n bt.click();\n"));
        //sb.Append("\n}");
        //sb.Append("else\n{");
        //sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
        //sb.Append("\n}");
        //sb.Append("</script>");
        //ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());


    }

    //protected void Button2_Click(object sender, EventArgs e)
    //{

    //    if (Hidden1.Value == "1")
    //    {
    //        if (txtDate.Text != "")
    //        {
    //            if (rbCapitalCall.Checked || rbDistribution.Checked)
    //            {
    //                if (rbCapitalCall.Checked)
    //                {


    //                    InsertCapitalCallGDGS();
    //                }
    //                else if (rbDistribution.Checked)
    //                {
    //                    InsertDistributionGDGS();

    //                }
    //            }
    //            else
    //            {
    //                lblError.Text = "Please Select File Type";
    //                lblError.ForeColor = System.Drawing.Color.Green;
    //            }
    //        }
    //        else
    //        {
    //            lblError.Text = "Please select Date";
    //            lblError.ForeColor = System.Drawing.Color.Green;
    //        }
    //        Hidden1.Value = "0";
    //    }
    //    Button2.Visible = false;
    //}

    protected void Button2_Click(object sender, EventArgs e)
    {
        if (Hidden3.Value == "1")
        {
            if (txtDate.Text != "")
            {
                if (rbCapitalCall.Checked || rbDistribution.Checked)
                {
                    if (rbCapitalCall.Checked)
                    {
                        InsertCapitalCallGDGS();
                    }
                    else if (rbDistribution.Checked)
                    {
                        InsertDistributionGDGS();
                    }
                }
                else
                {
                    lblError.Text = "Please Select File Type";
                    lblError.ForeColor = System.Drawing.Color.Green;
                }
            }
            else
            {
                lblError.Text = "Please select Date";
                lblError.ForeColor = System.Drawing.Color.Green;
            }
            // Hidden2.Value = "0";

        }

        if (Hidden3.Value == "0")
        {

            string msg = Convert.ToString(ViewState["msg"]);
            string error = Convert.ToString(ViewState["error"]);
            lblError.Text = msg;
            lblError1.Text = error;
        }

        Button2.Visible = false;
        Button3.Visible = false;
        Button1.Visible = false;
    }
    protected void Button1_Click1(object sender, EventArgs e)
    {


        if (Hidden2.Value == "1")
        {
            if (txtDate.Text != "")
            {
                if (rbCapitalCall.Checked || rbDistribution.Checked)
                {
                    if (rbCapitalCall.Checked)
                    {
                        DataSet dsData = getDataSet("SP_S_CAPITALCALLLETTER_GDGS @GDGSFlg = 1,@Asofdate='" + txtDate.Text + "'");
                        if (dsData.Tables[1].Rows.Count > 0)
                        {
                            Button2.Visible = true;
                            System.Text.StringBuilder sb = new System.Text.StringBuilder();
                            Type tp = this.GetType();
                            sb.Append("\n<script type=text/javascript>\n");
                            sb.Append("var bt = window.document.getElementById('Button2');\n");
                            sb.Append("if(!alert('There are existing GDGS records for the close date selected, which will be deleted and reloaded'))\n{");
                            sb.Append("\nwindow.document.getElementById('Hidden3').value='1';");
                            sb.Append("\n txtDateClear();");
                            sb.Append(("\n bt.click();\n"));
                            sb.Append("\n}");
                            sb.Append("else\n{");
                            sb.Append(("\nwindow.document.getElementById('Hidden3').value='0';"));
                            sb.Append(("\n bt.click();\n"));
                            sb.Append("\n txtDateClear();");
                            sb.Append("\n}");
                            sb.Append("</script>");
                            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                        }
                        else
                        {
                            InsertCapitalCallGDGS();
                        }
                    }
                    else if (rbDistribution.Checked)
                    {
                        DataSet dsData = getDataSet("SP_S_DISTRIBUTION_GDGS @GDGSFlg = 1,@Asofdate='" + txtDate.Text + "'");
                        if (dsData.Tables[1].Rows.Count > 0)
                        {
                            Button2.Visible = true;
                            System.Text.StringBuilder sb = new System.Text.StringBuilder();
                            Type tp = this.GetType();
                            sb.Append("\n<script type=text/javascript>\n");
                            sb.Append("var bt = window.document.getElementById('Button2');\n");
                            sb.Append("if(!alert('There are existing GDGS records for the close date selected, which will be deleted and reloaded'))\n{");
                            sb.Append("\nwindow.document.getElementById('Hidden3').value='1';");
                            sb.Append("\n txtDateClear();");
                            sb.Append(("\n bt.click();\n"));
                            sb.Append("\n}");
                            sb.Append("else\n{");
                            sb.Append(("\nwindow.document.getElementById('Hidden3').value='0';"));
                            sb.Append(("\n bt.click();\n"));
                            sb.Append("\n txtDateClear();");
                            sb.Append("\n}");
                            sb.Append("</script>");
                            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                        }
                        else
                        {

                            InsertDistributionGDGS();
                        }

                    }
                }
                else
                {
                    lblError.Text = "Please Select File Type";
                    lblError.ForeColor = System.Drawing.Color.Green;
                }
            }
            else
            {
                lblError.Text = "Please select Date";
                lblError.ForeColor = System.Drawing.Color.Green;
            }
            Hidden1.Value = "0";
        }

        else if (Hidden2.Value == "0")
        {

            string mgs = Convert.ToString(ViewState["msg"]);
            string error = Convert.ToString(ViewState["error"]);
            lblError.Text = "File Uploaded Successfully" + mgs;
            lblError1.Text = error;
        }
        Button1.Visible = false;
    }

    protected void Button3_Click(object sender, EventArgs e)
    {
        List<Int32> count = new List<int>();

        string TempPath1 = Convert.ToString(ViewState["Temppath"]);
        if (countSelectedFund == 0)
        {
            Proces_ALL_FUND_ReInsert(TempPath1, service, rbCapitalCall.Checked, rbDistribution.Checked);

            ViewState["msg"] = msg;

            ViewState["error"] = ErrorOccured;
        }

        if (bpopup || countpopup > 0)
        {
            Button1.Visible = true;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");
            sb.Append("var bt = window.document.getElementById('Button1');\n");
            sb.Append("if(!alert('Done creating fund recommendations, ready to create the GDGS recommendations.  Click OK to continue.'))\n{");
            sb.Append("\nwindow.document.getElementById('Hidden2').value='1';");
            sb.Append("\n txtDateClear();");
            sb.Append(("\n bt.click();\n"));
            sb.Append("\n}");
            sb.Append("else\n{");
            // sb.Append(("\nwindow.document.getElementById('Hidden2').value='0';"));
            sb.Append(("\nwindow.document.getElementById('Hidden2').value='0';"));
            sb.Append(("\n bt.click();\n"));
            sb.Append("\n txtDateClear();");
            sb.Append("\n}");
            sb.Append("</script>");
            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
        }


    }
}




