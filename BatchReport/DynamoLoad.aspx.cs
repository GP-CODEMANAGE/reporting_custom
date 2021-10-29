using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.SqlClient;
using System.Security.Principal;
using System.IO;
using System.Text;
using Spire.Xls;
using System.Xml;

using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Data;
using System.Data.OleDb;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class DynamoLoad : System.Web.UI.Page
{
    bool bProceed = false;
    string sqlstr = string.Empty;
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    Logs lg = new Logs();
    public StreamWriter sw = null;
    public string execType = string.Empty;
    string strDescription = "";
    bool bTransaction = false;
    bool bPosition = false;
    bool bSummaryValuation = false;
    bool bPerformance = false;

    public const string Commitment = "Commitment";
    public const string Summary = "Summary";
    public const string Transaction = "Transaction";
    public const string Performance = "Performance";


    DataSet dsPosition = null;
    DataSet dsTransaction = null;
    DataSet dsSummaryValuation = null;
    DataSet dsPerformance = null;


    bool bchecktransaction = false;
    bool bcheckcommit = false;
    bool bcheckposition = false;
    bool bcheckperformance = false;




    Dictionary<string, string> checksucess = new Dictionary<string, string>();

    Dictionary<string, string> checkLock = new Dictionary<string, string>();

    /******/
    DataSet DSTrans = null;
    DataSet DSCommit = null;
    DataSet DSSummary = null;
    DataSet DSPerformance = new DataSet();
    /**/
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            mvLoad.ActiveViewIndex = 0;
            // BindFundType();
            // lstFundType.SelectedValue = "1";
            //BindFund();
            // BindLegalEntity();
            trDownLoad.Style.Add("display", "none");




        }
    }



    protected void lstFundType_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        // BindFund();
    }





    private void ConfirmUserToRunLoad()
    {

        uplhide.Value = "1";
        StringBuilder sb = new StringBuilder();

        sb.Append("<script language=\"javascript\">");
        sb.Append("var x=confirm(\"Data is locked. Do you want to continue the load?\");");
        sb.Append("if (x == true) {");
        sb.Append(" document.getElementById(\"ConfirmLock\").value = \"1\";");
        sb.Append(" document.getElementById(\"btnLoadRun\").click();");
        sb.Append("}");
        sb.Append("else {");
        sb.Append("document.getElementById(\"ConfirmLock\").value = \"0\";");
        sb.Append(" }");
        sb.Append("</script>");

        ClientScript.RegisterStartupScript(typeof(Page), "confirm", sb.ToString());
    }

    protected void btnLoadRun_Click(object sender, EventArgs e)
    {
        clearlabel();


        #region Log Declaration
        DateTime dtmain = DateTime.Now;
        string LogFileName = string.Empty;
        LogFileName = "Log-" + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        sw = new StreamWriter(LogFileName);
        sw.Close();
        HttpContext.Current.Session["Filename"] = LogFileName;
        // ViewState["Filename"] = LogFileName;

        //  LogFileName = (string)ViewState["Filename"];

        //  Session["Filename"] = LogFileName;

        #endregion

        #region Clear old Backupfiles
        lg.AddinLogFile(Session["Filename"].ToString(), " Clear old Backupfiles(3month older Files) START " + DateTime.Now.ToString());
        if (Directory.Exists(Server.MapPath("") + @"\ExcelTemplate\DynamoBackup"))
        {
            string[] files = Directory.GetFiles(Server.MapPath("") + @"\ExcelTemplate\DynamoBackup");

            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                string FileName = fi.Name;

                if (fi.LastAccessTime < DateTime.Now.AddMonths(-3))
                {
                    fi.Delete();
                    lg.AddinLogFile(Session["Filename"].ToString(), "Deleted " + FileName + " " + DateTime.Now.ToString());
                }
            }
        }

        #endregion

        lblError.Text = "";
        bool DataLoaded = false;

        int intConfirmLock = int.Parse(ConfirmLock.Value);
        if (intConfirmLock == 1)
        {
            ConfirmLock.Value = "0";
            LoadFile();
            if (bProceed)
            {
                if (bPosition || bTransaction || bSummaryValuation || bPerformance)
                {
                    bool newSec = ShowLinkforNewSec();
                    if (newSec == false)
                    {

                        DataLoaded = LoadData();
                        if (DataLoaded == true)
                        {

                            #region Not Used
                            //string msg = "";
                            //bool btrans = (bool)Session["bchecktransaction"];
                            //bool bsumm = (bool)Session["bcheckposition"];
                            //bool bcommit = (bool)Session["bcheckcommit"];

                            //if (btrans)
                            //    msg = "Transaction";
                            //if (bsumm)
                            //    msg = msg + "Summary";
                            //if (bcommit)
                            //    msg = msg + "Commitment";


                            //Session.Remove("bchecktransaction");
                            //Session.Remove("bcheckposition");
                            //Session.Remove("bcheckcommit");



                            //if (btrans || bsumm || bcommit)
                            //{

                            //    lblError.Text = "Load run successfully please check load log for Details.";

                            //    Session.Remove("dsTransaction");
                            //    Session.Remove("dsPosition");
                            //    Session.Remove("dsSummaryValuation");
                            //}

                            #endregion

                            #region File Format Validation
                            //if (bchecktransaction && bcheckposition && bcheckcommit)
                            //    lblError.Text = "Load Completed Succssfully for all";
                            //else if (bchecktransaction && bcheckposition)
                            //    lblError.Text = "load failed for Commitment,sucess for rest.";
                            //else if (bcheckcommit && bcheckposition)
                            //    lblError.Text = "load failed for Transaction,sucess for rest.";
                            //else if (bchecktransaction && bcheckcommit)
                            //    lblError.Text = "load failed for Summaryvaluation,sucess for rest.";
                            //else if (bchecktransaction)
                            //    lblError.Text = "load sucess for Transaction.";
                            //else if (bcheckcommit)
                            //    lblError.Text = "load Sucess for Commitment.";
                            //else if (bcheckposition)
                            //    lblError.Text = "load Sucess for SummaryValution.";
                            //else if (bTransaction == true && bchecktransaction == false)
                            //    lblError.Text = "load failed for Transaction.";
                            //else if (bSummaryValuation == true && bcheckposition == false)
                            //    lblError.Text = "load failed for SummaryValution.";
                            //else if (bPosition == true && bcheckcommit == false)
                            //    lblError.Text = "load failed for Transaction.";

                            //string msg1 = string.Empty;
                            //string msg2 = string.Empty;

                            //foreach (KeyValuePair<string, string> pair in checksucess)
                            //{
                            //    string key = pair.Key.ToString();

                            //    string Value = pair.Value.ToString();

                            //    if (Value.ToLower() == "true")
                            //    {
                            //        msg1 = msg1 + key + ",";

                            //    }

                            //    else
                            //    {

                            //        msg2 = msg2 + key + ",";
                            //    }


                            //}


                            //if (msg1 != "" && msg2 != "")
                            //    lblError.Text = "Load successful for " + msg1 + " " + "load Failed for " + msg2;

                            //else if (msg1 != "" && msg2 == "")
                            //    lblError.Text = "Load successful for " + msg1;

                            //else if (msg2 != "" && msg1 == "")
                            //    lblError.Text = "load Failed for " + msg2;


                            #endregion


                            CheckFileFormat();

                            //if (msg1 != "" && msg2 != "")
                            //    lblError.Text = "Load successful for " + msg1 + " " + "load Failed for " + msg2;

                            //else if (msg1 != "" && msg2 == "")
                            //    lblError.Text = "Load successful for " + msg1;

                            //else if (msg2 != "" && msg1 == "")
                            //    lblError.Text = "load Failed for " + msg2;



                            Session.Remove("dsTransaction");
                            Session.Remove("dsPosition");
                            Session.Remove("dsSummaryValuation");
                            Session.Remove("dsPerformance");

                            Session.Remove("bchecktransaction");
                            Session.Remove("bcheckposition");
                            Session.Remove("bcheckcommit");
                            //  Session.Remove("bcheckperformance");





                        }
                        else
                            lblError.Text = "No records loaded from Dynamo.";
                        trSubmit.Style.Add("display", "inline");
                    }
                }
            }
        }
        else
        {
            LoadFile();
            bool IsLocked = CheckLockStatus();
            if (IsLocked == true)
            {
                ConfirmUserToRunLoad();
            }
            else
            {
                // LoadFile();
                if (bProceed)
                {
                    if (bPosition || bTransaction || bSummaryValuation || bPerformance)
                    {
                        bool newSec = ShowLinkforNewSec();
                        if (newSec == false)
                        {

                            DataLoaded = LoadData();
                            if (DataLoaded == true)
                            {
                                string msg = "";
                                //bool btrans = (bool)Session["bchecktransaction"];
                                //bool bsumm = (bool)Session["bcheckposition"];
                                //bool bcommit = (bool)Session["bcheckcommit"];

                                //if (btrans)
                                //    msg = "Transaction";
                                //if (bsumm)
                                //    msg = msg + "Summary";
                                //if (bcommit)
                                //    msg = msg + "Commitment";


                                //lblError.Text = msg;

                                //Session.Remove("bchecktransaction");
                                //Session.Remove("bcheckposition");
                                //Session.Remove("bcheckcommit");

                                //if (btrans || bsumm || bcommit)
                                //{
                                //    lblError.Text = "Load run successfully please check load log for Details.";
                                //    Session.Remove("dsTransaction");
                                //    Session.Remove("dsPosition");
                                //    Session.Remove("dsSummaryValuation");
                                //}


                                #region Not Used

                                //Session.Remove("dsTransaction");
                                //Session.Remove("dsPosition");
                                //Session.Remove("dsSummaryValuation");
                                ////}

                                //if (bchecktransaction)
                                //    msg = "Transaction,";
                                //if (bcheckposition)
                                //    msg = msg + "Summary,";
                                //if (bcheckcommit)
                                //    msg = msg + "Commitment,";

                                //if (bPosition == true || bTransaction == true || bSummaryValuation == true)
                                //{
                                //    int cnt = 0;
                                //    //if (bTransaction == true && bchecktransaction == false)
                                //    //    msg = "Transaction";
                                //    //if (bPosition == true && bcheckcommit == false)
                                //    //    msg = "Commitment";
                                //    //if (bSummaryValuation == true && bcheckposition == false)
                                //    //    msg = "SummaryValuation";

                                //    if (bTransaction == true && bchecktransaction == false || bPosition == true && bcheckcommit == false || bSummaryValuation == true && bcheckposition == false)
                                //    {
                                //        if (bTransaction)
                                //        {
                                //            msg = "Transaction,";
                                //            cnt++;
                                //        }
                                //        if (bPosition)
                                //        {
                                //            msg = "Commitment,";
                                //            cnt++;
                                //        }
                                //        if (bSummaryValuation)
                                //        {
                                //            msg = "SummaryValuation,";
                                //            cnt++;
                                //        }

                                //        if (cnt > 0)
                                //            msg = msg + "Load Completed Successfully";
                                //        else if (cnt == 1)
                                //            msg = "Load Completed Successfully";
                                //    }


                                //    if (bTransaction == true && bchecktransaction == true || bPosition == true && bcheckcommit == true || bSummaryValuation == true && bcheckposition == true)
                                //    {
                                //        if (bTransaction)
                                //        {
                                //            msg = "Transaction,";
                                //            cnt++;
                                //        }
                                //        if (bPosition)
                                //        {
                                //            msg = "Commitment,";
                                //            cnt++;
                                //        }
                                //        if (bSummaryValuation)
                                //        {
                                //            msg = "SummaryValuation,";
                                //            cnt++;
                                //        }

                                //    }


                                //}



                                #endregion



                                //else


                                //{
                                //    msg = msg + "did not load successfully please contact administrator.";
                                //}

                                #region File Format Validation
                                //if (bchecktransaction && bcheckposition && bcheckcommit)
                                //    lblError.Text = "Load Completed Succssfully for all";
                                //else if (bchecktransaction && bcheckposition)
                                //    lblError.Text = "load failed for Commitment,sucess for rest.";
                                //else if (bcheckcommit && bcheckposition)
                                //    lblError.Text = "load failed for Transaction,sucess for rest.";
                                //else if (bchecktransaction && bcheckcommit)
                                //    lblError.Text = "load failed for Summaryvaluation,sucess for rest.";
                                //else if (bchecktransaction)
                                //    lblError.Text = "load sucess for Transaction.";
                                //else if (bcheckcommit)
                                //    lblError.Text = "load Sucess for Commitment.";
                                //else if (bcheckposition)
                                //    lblError.Text = "load Sucess for SummaryValution.";
                                //else if (bTransaction == true && bchecktransaction == false)
                                //    lblError.Text = "load failed for Transaction.";
                                //else if (bSummaryValuation == true && bcheckposition == false)
                                //    lblError.Text = "load failed for SummaryValution.";
                                //else if (bPosition == true && bcheckcommit == false)
                                //    lblError.Text = "load failed for Transaction.";

                                //string msg1 = string.Empty;
                                //string msg2 = string.Empty;

                                //foreach (KeyValuePair<string, string> pair in checksucess)
                                //{
                                //    string key = pair.Key.ToString();

                                //    string Value = pair.Value.ToString();

                                //    if (Value.ToLower() == "true")
                                //    {
                                //        if (msg1 != "")
                                //            msg1 = msg1 + "," + key;
                                //        else
                                //            msg1 = key;
                                //    }

                                //    else
                                //    {
                                //        if (msg2 != "")
                                //            msg2 = msg2 + key + ",";
                                //        else
                                //            msg2 = key;
                                //    }


                                //}


                                //if (msg1 != "" && msg2 != "")
                                //    lblError.Text = "Load successful for " + msg1 + " " + "load Failed for " + msg2;

                                //else if (msg1 != "" && msg2 == "")
                                //    lblError.Text = "Load successful for " + msg1;

                                //else if (msg2 != "" && msg1 == "")
                                //    lblError.Text = "load Failed for " + msg2;


                                #endregion

                                CheckFileFormat();


                                Session.Remove("dsTransaction");
                                Session.Remove("dsPosition");
                                Session.Remove("dsSummaryValuation");
                                Session.Remove("dsPerformance");


                                Session.Remove("bchecktransaction");
                                Session.Remove("bcheckposition");
                                Session.Remove("bcheckcommit");

                                // lblError.Text = msg;




                            }
                            else
                            {
                                string msg = "";
                                //bool btrans = (bool)Session["bchecktransaction"];
                                //bool bsumm = (bool)Session["bcheckposition"];
                                //bool bcommit = (bool)Session["bcheckcommit"];

                                //if (bchecktransaction)
                                //    msg = "Transaction,";
                                //if (bcheckposition)
                                //    msg = msg + "Summary,";
                                //if (bcheckcommit)
                                //    msg = msg + "Commitment,";

                                //if (bPosition == true || bTransaction == true || bSummaryValuation == true)
                                //{
                                //    if (bTransaction)
                                //        msg = "Transaction";
                                //    msg = msg + "did not load successfully please contact administrator, rest of the load completed successfully ";
                                //}

                                //else
                                //{
                                //    msg = msg + "did not load successfully please contact administrator.";
                                //}

                                //lblError.Text = msg;

                                //Session.Remove("bchecktransaction");
                                //Session.Remove("bcheckposition");
                                //Session.Remove("bcheckcommit");

                                //if (!bchecktransaction && !bcheckposition && !bcheckcommit)
                                //    lblError.Text = "No records loaded from Dynamo.";

                                #region not used

                                //if (bchecktransaction)
                                //    msg = "Transaction,";
                                //if (bcheckposition)
                                //    msg = msg + "Summary,";
                                //if (bcheckcommit)
                                //    msg = msg + "Commitment,";

                                //if (bPosition == true || bTransaction == true || bSummaryValuation == true)
                                //{
                                //    if (bTransaction == true && bchecktransaction == false)
                                //        msg = "Transaction";
                                //    if (bPosition == true && bcheckcommit == false)
                                //        msg = "Commitment";
                                //    if (bSummaryValuation == true && bcheckposition == false)
                                //        msg = "SummaryValuation";
                                //    msg = msg + "did not load successfully please contact administrator, rest of the load completed successfully ";
                                //}


                                #endregion
                                //else
                                //{
                                //    msg = msg + "did not load successfully please contact administrator.";
                                //}

                                #region File Format Validation
                                //if (bchecktransaction && bcheckposition && bcheckcommit)
                                //    lblError.Text = "Load Completed Succssfully for all";
                                //else if (bchecktransaction && bcheckposition)
                                //    lblError.Text = "load failed for Commitment,sucess for rest.";
                                //else if (bcheckcommit && bcheckposition)
                                //    lblError.Text = "load failed for Transaction,sucess for rest.";
                                //else if (bchecktransaction && bcheckcommit)
                                //    lblError.Text = "load failed for Summaryvaluation,sucess for rest.";
                                //else if (bchecktransaction)
                                //    lblError.Text = "load sucess for Transaction.";
                                //else if (bcheckcommit)
                                //    lblError.Text = "load Sucess for Commitment.";
                                //else if (bcheckposition)
                                //    lblError.Text = "load Sucess for SummaryValution.";
                                //else if (bTransaction == true && bchecktransaction == false)
                                //    lblError.Text = "load failed for Transaction.";
                                //else if (bSummaryValuation == true && bcheckposition == false)
                                //    lblError.Text = "load failed for SummaryValution.";
                                //else if (bPosition == true && bcheckcommit == false)
                                //    lblError.Text = "load failed for Transaction.";

                                //string msg1 = string.Empty;
                                //string msg2 = string.Empty;

                                //foreach (KeyValuePair<string, string> pair in checksucess)
                                //{
                                //    string key = pair.Key.ToString();

                                //    string Value = pair.Value.ToString();

                                //    if (Value.ToLower() == "true")
                                //    {
                                //        msg1 = msg1 + key + ",";

                                //    }

                                //    else
                                //    {

                                //        msg2 = msg2 + key + ",";
                                //    }


                                //}


                                //if (msg1 != "" && msg2 != "")
                                //    lblError.Text = "Load successful for " + msg1 + " " + "load Failed for " + msg2;

                                //else if (msg1 != "" && msg2 == "")
                                //    lblError.Text = "Load successful for " + msg1;

                                //else if (msg2 != "" && msg1 == "")
                                //    lblError.Text = "load Failed for " + msg2;


                                #endregion

                                CheckFileFormat();

                                // lblError.Text = msg;

                                Session.Remove("dsTransaction");
                                Session.Remove("dsPosition");
                                Session.Remove("dsSummaryValuation");
                                Session.Remove("dsPerformance");



                                Session.Remove("bchecktransaction");
                                Session.Remove("bcheckposition");
                                Session.Remove("bcheckcommit");

                            }
                            trSubmit.Style.Add("display", "inline");
                        }
                    }
                }
            }
        }
    }
    public void LoadFile()
    {


        DataSet ds = (DataSet)Session["dsTransaction"];

        DataSet ds1 = (DataSet)Session["dsPosition"];

        DataSet ds2 = (DataSet)Session["dsSummaryValuation"];

        DataSet ds3 = (DataSet)Session["dsPerformance"];

        //DataSet ds = (DataSet)ViewState["dsTransaction"];

        //DataSet ds1 = (DataSet)ViewState["dsPosition"];

        //DataSet ds2 = (DataSet)ViewState["dsSummaryValuation"];



        if (!Directory.Exists(Server.MapPath("") + @"\ExcelTemplate\TempFolder"))
        {
            Directory.CreateDirectory(Server.MapPath("") + @"\ExcelTemplate\TempFolder");
        }
        if (!Directory.Exists(Server.MapPath("") + @"\ExcelTemplate\DynamoBackup"))
        {
            Directory.CreateDirectory(Server.MapPath("") + @"\ExcelTemplate\DynamoBackup");
        }

        if (UplTransaction.HasFile)
        {

            UploadTransaction();
        }

        else if (ds != null)
        {
            bProceed = true;
            bTransaction = true;
            dsTransaction = ds.Copy();
        }


        if (UPLPosition.HasFile)
        {
            UploadCommit();
        }


        else if (ds1 != null)
        {
            bProceed = true;
            bPosition = true;
            dsPosition = ds1.Copy();
        }


        if (UplSummaryValuation.HasFile)
        {
            UploadSummary();
        }

        else if (ds2 != null)
        {
            bProceed = true;
            bSummaryValuation = true;
            dsSummaryValuation = ds2.Copy();
        }


        if (UPLPerf.HasFile)
        {
            UploadPerformance();
        }


        else if (ds3 != null)
        {
            bProceed = true;
            bPerformance = true;
            dsPerformance = ds3.Copy();
        }

        //  return bProceed;
    }

    public void CheckFileFormat()
    {
        #region File Format Validation


        //string msg1 = string.Empty;
        //string msg2 = string.Empty;

        //foreach (KeyValuePair<string, string> pair in checksucess)
        //{
        //    string key = pair.Key.ToString();

        //    string Value = pair.Value.ToString();

        //    if (Value.ToLower() == "true")
        //    {
        //        if (msg1 != "")
        //            msg1 = msg1 + "," + key;
        //        else
        //            msg1 = key;
        //    }

        //    else
        //    {
        //        if (msg2 != "")
        //            msg2 = msg2 + "," + key;
        //        else
        //            msg2 = key;
        //    }


        //}


        //if (msg1 != "" && msg2 != "")
        //    lblError.Text = "Load successful for " + msg1 + " " + "Load Failed for " + msg2;

        //else if (msg1 != "" && msg2 == "")
        //    lblError.Text = "Load successful for " + msg1;

        //else if (msg2 != "" && msg1 == "")
        //    lblError.Text = "Load Failed for " + msg2;

        // lblError.Text = "Transaction   <span style='color:green'>Sucess</span> ";

        string msg1 = string.Empty;
        string msg2 = string.Empty;

        foreach (KeyValuePair<string, string> pair in checksucess)
        {
            string key = pair.Key.ToString();

            string Value = pair.Value.ToString();

            if (Value.ToLower() == "true")
            {
                if (key.ToLower() == "transaction")
                    lbltrans.Text = "Transaction        Success";
                else if (key.ToLower() == "commitment")
                    lblcommit.Text = "Commitment        Success";
                else if (key.ToLower() == "summaryvaluation")
                    lblsummary.Text = "SummaryValuation Success";
                else if (key.ToLower() == "performance")
                    lblperf.Text = "Performance         Success";
            }

            else
            {
                if (key.ToLower() == "transaction")
                    lbltrans.Text = "Transaction         Failed";
                else if (key.ToLower() == "commitment")
                    lblcommit.Text = "Commitment         Failed";
                else if (key.ToLower() == "summaryvaluation")
                    lblsummary.Text = "SummaryValuation  Failed";
                else if (key.ToLower() == "performance")
                    lblperf.Text = "Performance          Failed";
            }


        }


        //if (msg1 != "" && msg2 != "")
        //    lblError.Text = "Load successful for " + msg1 + " " + "Load Failed for " + msg2;

        //else if (msg1 != "" && msg2 == "")
        //    lblError.Text = "Load successful for " + msg1;

        //else if (msg2 != "" && msg1 == "")
        //    lblError.Text = "Load Failed for " + msg2;


        #endregion
    }

    public void ProcessData(DataSet ds)
    {

        //for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
        //{
        //    if (Convert.ToString(ds.Tables[0].Rows[i]["Fund Admin ID"]) == "")
        //        //ds.Tables[0].Rows[i].Delete();
        //}

        try
        {

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string dDate = Convert.ToString(ds.Tables[0].Rows[i]["Effective Date"]);
                //  string strTxnNum = Convert.ToString(ds.Tables[0].Rows[i]["Txn Num"]);

                try
                {
                    DateTime Effective = Convert.ToDateTime(ds.Tables[0].Rows[i]["Effective Date"]);
                }
                catch (Exception ex)
                {
                    ds.Tables[0].Rows[i]["Effective Date"] = DBNull.Value;
                }
                try
                {
                    string strTxnNum = Convert.ToString(ds.Tables[0].Rows[i]["Txn Num"]);
                    //   ds.Tables[0].Rows[i]["Txn Num"] = new Guid(Convert.ToString(strTxnNum));
                }
                catch (Exception ex)
                {
                    lblError.Text = "Error converting Txn Num to Guid";
                    bProceed = false;
                }
            }
        }
        catch (Exception ex)
        {

        }

    }

    public void UploadTransaction()
    {
        try
        {

            string filename = Path.GetFileName(UplTransaction.FileName);
            string strextension = Path.GetExtension(UplTransaction.FileName);
            if (strextension == ".xlsx")
            {
                //ViewState["FilePath"] = filename;
                string FilePath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + filename;
                if (File.Exists(FilePath))
                    File.Delete(FilePath);
                ViewState["FilePath"] = FilePath;
                UplTransaction.SaveAs(FilePath);

                #region Backup

                string DateTimenow = DateTime.Now.ToString();
                DateTimenow = DateTimenow.Replace(":", "-");
                DateTimenow = DateTimenow.Replace("/", "-");

                strextension = Path.GetExtension(UplTransaction.FileName);
                string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(UplTransaction.FileName);
                string BackupFilePath = Server.MapPath("") + @"\ExcelTemplate\DynamoBackup\" + FileNameWithoutExtension + DateTimenow + strextension;
                UplTransaction.SaveAs(BackupFilePath);
                #endregion

                //Read Excel File 
                dsTransaction = GetDataTableFromExcel(FilePath);

                DSTrans = dsTransaction.Copy();
                Session["dsTransaction"] = dsTransaction;

                //  ViewState["dsTransaction"] = dsTransaction;

                if (dsTransaction.Tables.Count > 0)
                {
                    ProcessData(dsTransaction);
                    File.Delete(FilePath);
                    bProceed = true;
                    bTransaction = true;
                }
            }
            else
            {
                bProceed = false;
                string Desc = "Please Enter Correct File";
                lblError.Text = Desc;
                lblError.Visible = true;
                lg.AddinLogFile(Session["Filename"].ToString(), Desc + "   " + strextension + DateTime.Now.ToString());
            }
        }
        catch (Exception ex)
        {
            string Desc = "ERROR: Incorrect File format ";// + ex.Message;
            lblError.Text = Desc;
            lblError.Visible = true;
            lg.AddinLogFile(Session["Filename"].ToString(), Desc + DateTime.Now.ToString());
            lg.AddinLogFile(Session["Filename"].ToString(), "ERROR: " + ex.Message.ToString() + "  " + DateTime.Now.ToString());
            // bProceed = false;
        }
    }

    public void UploadCommit()
    {
        try
        {

            string filename = Path.GetFileName(UPLPosition.FileName);
            string strextension = Path.GetExtension(UPLPosition.FileName);
            if (strextension == ".xlsx")
            {

                //ViewState["FilePath"] = filename;
                string FilePath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + filename;
                if (File.Exists(FilePath))
                    File.Delete(FilePath);
                ViewState["FilePath"] = FilePath;
                UPLPosition.SaveAs(FilePath);

                #region Backup

                string DateTimenow = DateTime.Now.ToString();
                DateTimenow = DateTimenow.Replace(":", "-");
                DateTimenow = DateTimenow.Replace("/", "-");

                strextension = Path.GetExtension(UPLPosition.FileName);
                string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(UPLPosition.FileName);
                string BackupFilePath = Server.MapPath("") + @"\ExcelTemplate\DynamoBackup\" + FileNameWithoutExtension + DateTimenow + strextension;
                UPLPosition.SaveAs(BackupFilePath);
                #endregion

                //Read Excel File 
                dsPosition = GetDataTableFromExcel(FilePath);

                DSCommit = dsPosition.Copy();

                Session["dsPosition"] = dsPosition;

                //ViewState["dsPosition"] = dsPosition;

                if (dsPosition.Tables.Count > 0)
                {
                    //  ProcessData(dsPosition);
                    File.Delete(FilePath);
                    bProceed = true;
                    bPosition = true;
                }
            }
            else
            {
                bProceed = false;
                string Desc = "Please Enter Correct File";
                lblError.Text = Desc;
                lblError.Visible = true;
                lg.AddinLogFile(Session["Filename"].ToString(), Desc + "   " + strextension + DateTime.Now.ToString());
            }
        }
        catch (Exception ex)
        {
            string Desc = "ERROR: Incorrect File format ";// + ex.Message;
            lblError.Text = Desc;
            lblError.Visible = true;
            lg.AddinLogFile(Session["Filename"].ToString(), Desc + DateTime.Now.ToString());
            lg.AddinLogFile(Session["Filename"].ToString(), "ERROR: " + ex.Message.ToString() + "  " + DateTime.Now.ToString());
            // bProceed = false;
        }
    }

    public void UploadSummary()
    {
        try
        {
            string filename = Path.GetFileName(UplSummaryValuation.FileName);
            string strextension = Path.GetExtension(UplSummaryValuation.FileName);
            if (strextension == ".xlsx")
            {

                //ViewState["FilePath"] = filename;
                string FilePath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + filename;
                if (File.Exists(FilePath))
                    File.Delete(FilePath);
                ViewState["FilePath"] = FilePath;
                UplSummaryValuation.SaveAs(FilePath);

                #region Backup
                string DateTimenow = DateTime.Now.ToString();
                DateTimenow = DateTimenow.Replace(":", "-");
                DateTimenow = DateTimenow.Replace("/", "-");

                strextension = Path.GetExtension(UplSummaryValuation.FileName);
                string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(UplSummaryValuation.FileName);
                string BackupFilePath = Server.MapPath("") + @"\ExcelTemplate\DynamoBackup\" + FileNameWithoutExtension + DateTimenow + strextension;
                UplSummaryValuation.SaveAs(BackupFilePath);
                #endregion

                //Read Excel File 
                dsSummaryValuation = GetDataTableFromExcel(FilePath);

                DSSummary = dsSummaryValuation.Copy();

                Session["dsSummaryValuation"] = dsSummaryValuation;

                //ViewState["dsSummaryValuation"] = dsSummaryValuation;

                if (dsSummaryValuation.Tables.Count > 0)
                {
                    //  ProcessData(dsSummaryValuation);
                    File.Delete(FilePath);
                    bProceed = true;
                    bSummaryValuation = true;
                }
            }
            else
            {
                bProceed = false;
                string Desc = "Please Enter Correct File";
                lblError.Text = Desc;
                lblError.Visible = true;
                lg.AddinLogFile(Session["Filename"].ToString(), Desc + "   " + strextension + DateTime.Now.ToString());
            }
        }
        catch (Exception ex)
        {
            string Desc = "ERROR: Incorrect File format ";// + ex.Message;
            lblError.Text = Desc;
            lblError.Visible = true;
            lg.AddinLogFile(Session["Filename"].ToString(), Desc + DateTime.Now.ToString());
            lg.AddinLogFile(Session["Filename"].ToString(), "ERROR: " + ex.Message.ToString() + "  " + DateTime.Now.ToString());
            // bProceed = false;
        }
    }

    public void cleansessionvalue()
    {
        Session.Remove("dsTransaction");
        Session.Remove("dsPosition");
        Session.Remove("dsSummaryValuation");
        Session.Remove("dsPerformance");
    }


    public void UploadPerformance()
    {
        try
        {
            string filename = Path.GetFileName(UPLPerf.FileName);
            string strextension = Path.GetExtension(UPLPerf.FileName);
            if (strextension == ".xlsx")
            {

                //ViewState["FilePath"] = filename;
                string FilePath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + filename;
                if (File.Exists(FilePath))
                    File.Delete(FilePath);
                ViewState["FilePath"] = FilePath;
                UPLPerf.SaveAs(FilePath);

                #region Backup
                string DateTimenow = DateTime.Now.ToString();
                DateTimenow = DateTimenow.Replace(":", "-");
                DateTimenow = DateTimenow.Replace("/", "-");

                strextension = Path.GetExtension(UplSummaryValuation.FileName);
                string FileNameWithoutExtension = Path.GetFileNameWithoutExtension(UplSummaryValuation.FileName);
                string BackupFilePath = Server.MapPath("") + @"\ExcelTemplate\DynamoBackup\" + FileNameWithoutExtension + DateTimenow + strextension;
                UPLPerf.SaveAs(BackupFilePath);
                #endregion

                //Read Excel File 
                dsPerformance = GetDataTableFromExcel1(FilePath);

                // dsPerformance = Parse(FilePath);

                //foreach (DataRow dr in dsPerformance.Tables[0].Rows)
                //{

                //    try
                //    {
                //        // DateTime AsOfDate = Convert.ToDateTime(dr[0].ToString());
                //        string AsOfDate = dr[0].ToString();

                //        if (!AsOfDate.Contains("/"))
                //        {
                //            dr.Delete();
                //        }
                //    }

                //    catch (Exception ex)
                //    {
                //        //dr.Delete();
                //    }

                //    //   Datat
                //}

                #region delete row in dataset
                int Beforecount = dsPerformance.Tables[0].Rows.Count;
                for (int i = 0; i < dsPerformance.Tables[0].Rows.Count; i++)
                {
                    try
                    {

                        string AsOFDate = dsPerformance.Tables[0].Rows[i][0].ToString();

                        //string MTD = dsPerformance.Tables[0].Rows[i]["MTD"].ToString();

                        //string QTD = dsPerformance.Tables[0].Rows[i]["QTD"].ToString();

                        //string YTD = dsPerformance.Tables[0].Rows[i]["YTD"].ToString();

                        //string YR1 = dsPerformance.Tables[0].Rows[i]["1 YR"].ToString();

                        //string YR3 = dsPerformance.Tables[0].Rows[i]["3 YR"].ToString();

                        //string YR5 = dsPerformance.Tables[0].Rows[i]["5 YR"].ToString();


                        //if(MTD =="-" || QTD =="-" || YTD=="-" || YR1=="-" || YR1=="-" || YR3=="-" || YR5=="-")

                        // string AsOFDate = "01-16-2019";

                        // DateTime dt = DateTime.ParseExact(AsOFDate, "dd/mm/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                        bool dateflg = CheckDate(AsOFDate);
                        /* for householdfeed or 1st row delete from dataset */
                        //if (!AsOFDate.Contains("/"))//check "/" for find out if it contains date format 
                        //    dsPerformance.Tables[0].Rows[i].Delete();

                        //else
                        //{


                        //    dsPerformance.Tables[0].Rows[i][0] = Convert.ToDateTime(AsOFDate);
                        //}

                        if (dateflg)
                            dsPerformance.Tables[0].Rows[i][0] = Convert.ToDateTime(AsOFDate);
                        else
                            dsPerformance.Tables[0].Rows[i].Delete();
                    }
                    catch (Exception ex)
                    {

                    }
                }

                dsPerformance.Tables[0].AcceptChanges();

                int Aftercount = dsPerformance.Tables[0].Rows.Count;
                #endregion

                Session["dsPerformance"] = dsPerformance;


                if (dsPerformance.Tables.Count > 0)
                {
                    //  ProcessData(dsSummaryValuation);
                    try
                    {
                        bProceed = true;
                        bPerformance = true;
                        File.Delete(FilePath);

                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            else
            {
                bProceed = false;
                string Desc = "Please Enter Correct File";
                lblError.Text = Desc;
                lblError.Visible = true;
                lg.AddinLogFile(Session["Filename"].ToString(), Desc + "   " + strextension + DateTime.Now.ToString());
            }
        }
        catch (Exception ex)
        {
            string Desc = "ERROR: Incorrect File format ";// + ex.Message;
            lblError.Text = Desc;
            lblError.Visible = true;
            lg.AddinLogFile(Session["Filename"].ToString(), Desc + DateTime.Now.ToString());
            lg.AddinLogFile(Session["Filename"].ToString(), "ERROR: " + ex.Message.ToString() + "  " + DateTime.Now.ToString());
            // bProceed = false;
        }
    }

    public DataSet GetDataTableFromExcel(string path, bool hasHeader = true)
    {
        DataSet ds = new DataSet();
        try
        {
            // DataSet ds = new DataSet();
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                int wscnt = pck.Workbook.Worksheets.Count();

                for (int i = 1; i <= wscnt; i++)
                {
                    // ViewState["SheetName"] = pck.Workbook.Worksheets[i].Name;
                    var ws = pck.Workbook.Worksheets[i];
                    DataTable tbl = new DataTable();
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }
                    var startRow = hasHeader ? 2 : 1;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = tbl.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }
                    // return tbl;
                    ds.Tables.Add(tbl);

                }
                return ds;
            }
        }
        catch (Exception ex)
        {
            string Desc = "ExcelRead Failed. Please Contact Administrator " + ex.Message;
            lblError.Text = lblError.Text + "\n" + Desc;
            lblError.Visible = true;
            lg.AddinLogFile(Session["Filename"].ToString(), Desc + DateTime.Now.ToString());
            return null;
        }

    }

    protected bool CheckDate(String date)
    {
        DateTime Temp;


        if (DateTime.TryParse(date, out Temp) == true)
            return true;
        else
            return false;
    }


    public DataSet GetDataTableFromExcel1(string path, bool hasHeader = true)
    {
        DataSet ds = new DataSet();
        try
        {
            // DataSet ds = new DataSet();
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                int wscnt = pck.Workbook.Worksheets.Count();

                for (int i = 1; i <= wscnt; i++)
                {
                    // ViewState["SheetName"] = pck.Workbook.Worksheets[i].Name;

                    string sheetName = pck.Workbook.Worksheets[i].Name;

                    if (sheetName == "6. Aggregate All")
                    {

                        var ws = pck.Workbook.Worksheets[i];
                        DataTable tbl = new DataTable();
                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                        {
                            tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                        }
                        var startRow = hasHeader ? 2 : 1;
                        for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                            DataRow row = tbl.Rows.Add();
                            foreach (var cell in wsRow)
                            {

                                try
                                {
                                    if (cell.Text.Contains("%"))
                                    {
                                        string value1 = cell.Text;

                                        value1 = value1.Replace("%", "");

                                        double value = Convert.ToDouble(value1);

                                        value = value / 100;

                                        row[cell.Start.Column - 1] = value;
                                    }
                                    else
                                    {
                                        row[cell.Start.Column - 1] = cell.Text;
                                    }
                                }
                                catch(Exception ex)
                                {

                                }
                            }
                        }
                        // return tbl;
                        ds.Tables.Add(tbl);
                    }

                }
                return ds;
            }
        }
        catch (Exception ex)
        {
            string Desc = "ExcelRead Failed. Please Contact Administrator " + ex.Message;
            lblError.Text = lblError.Text + "\n" + Desc;
            lblError.Visible = true;
            lg.AddinLogFile(Session["Filename"].ToString(), Desc + DateTime.Now.ToString());
            return null;
        }

    }


    static DataSet Parse(string fileName)
    {
        bool hasHeaders = false;
        string HDR = hasHeaders ? "Yes" : "No";
        string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"", fileName);

        //   string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=0\"", fileName);


        // string connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=1\"", fileName);
        DataSet data = new DataSet();
        //  OleDbConnection con = new OleDbConnection();

        foreach (var sheetName in GetExcelSheetNames(connectionString))
        {

            using (OleDbConnection con = new OleDbConnection(connectionString))
            {

                var dataTable = new DataTable();
                string query = string.Format("SELECT * FROM [{0}]", sheetName);
                // con.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                adapter.Fill(dataTable);
                data.Tables.Add(dataTable);
                //con.Close();
            }
        }

        return data;
    }

    static string[] GetExcelSheetNames(string connectionString)
    {
        OleDbConnection con = null;
        DataTable dt = null;
        con = new OleDbConnection(connectionString);
        con.Open();
        dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

        if (dt == null)
        {
            return null;
        }

        String[] excelSheetNames = new String[dt.Rows.Count];
        int i = 0;

        foreach (DataRow row in dt.Rows)
        {
            excelSheetNames[i] = row["TABLE_NAME"].ToString();
            i++;
        }

        con.Close();
        return excelSheetNames;

    }



    protected void btnContinue_Click(object sender, EventArgs e)
    {

        LoadFile();

        if (bProceed)
        {
            if (bPosition || bTransaction || bSummaryValuation || bPerformance)
            {
                bool DataLoaded = false;
                DataLoaded = LoadData();
                if (DataLoaded == true)
                {
                    // lblError.Text = "Load run successfully please check load log for Details.";

                    CheckFileFormat();

                    Session.Remove("dsTransaction");
                    Session.Remove("dsPosition");
                    Session.Remove("dsSummaryValuation");
                    Session.Remove("dsPerformance");
                }
                else
                    lblError.Text = "No records loaded from Dynamo.";
                trSubmit.Style.Add("display", "inline");
                trDownLoad.Style.Add("display", "none");
            }
        }
    }


    protected void btnLock_Click(object sender, EventArgs e)
    {
        clearlabel();
        //object FundNameTxt = GetListBoxItem(lstFund) == "All" ? "''" : "'" + GetListBoxItem(lstFund) + "'";
        //object LegalEntityName = GetListBoxItem(lstLE) == "All" ? "''" : "'" + GetListBoxItem(lstLE) + "'";

        #region Log Declaration
        //DateTime dtmain = DateTime.Now;
        //string LogFileName = string.Empty;
        //LogFileName = "Log-" + DateTime.Now;
        //LogFileName = LogFileName.Replace(":", "-");
        //LogFileName = LogFileName.Replace("/", "-");
        //LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        //sw = new StreamWriter(LogFileName);
        //sw.Close();
        //HttpContext.Current.Session["Filename"] = LogFileName;
        // ViewState["Filename"] = LogFileName;

        //  LogFileName = (string)ViewState["Filename"];

        //  Session["Filename"] = LogFileName;

        #endregion

        // object FundType = lstFundType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFundType) + "'";
        // string username = WindowsIdentity.GetCurrent().Name;
        string FilePathTrans = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + "Transaction Blank File.xlsx";
        string FilePathCommit = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + "Commitment Blank File.xlsx";
        string FilePathSummary = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + "SummaryValuation Blank File.xlsx";


        string username = "corp\\crmadmin";




        //DSTrans = (DataSet)Session["dsTransaction"];
        //DSCommit = (DataSet)Session["dsPosition"];

        //DSSummary = (DataSet)Session["dsSummaryValuation"];
        ////  dsPerformance = (DataSet)ViewState["dsTransaction"];


        string LoadType = "";
        if (UplTransaction.HasFile)
        {
            UploadTransaction();
            LoadType = LoadType + ",3";
            checkLock.Add("Transaction", "true");

        }
        if (UPLPosition.HasFile)
        {
            UploadCommit();
            LoadType = LoadType + ",1";
            checkLock.Add("Commitment", "true");
        }
        if (UplSummaryValuation.HasFile)
        {
            UploadSummary();
            LoadType = LoadType + ",2";
            checkLock.Add("SummaryValuation", "true");
        }

        if (UPLPerf.HasFile)
        {
            UploadPerformance();
            LoadType = LoadType + ",4";
            checkLock.Add("Performance", "true");
        }

        LoadType = LoadType.Remove(0, 1);

        //if (DSTrans == null)
        //{
        //    DSTrans = GetDataTableFromExcel(FilePathTrans);
        //}


        //if (DSCommit == null)
        //{
        //    DSCommit = GetDataTableFromExcel(FilePathCommit);
        //}

        //if (DSSummary == null)
        //{
        //    DSSummary = GetDataTableFromExcel(FilePathSummary);
        //}


        //if (DSPerformance == null)
        //{
        //    DSPerformance = GetDataTableFromExcel(FilePathSummary);
        //}



        //sqlstr = "SP_S_TNR_LOAD_LOCK @StartDt='" + txtFrom.Text + "'" +
        //    ",@EndDt='" + txtTo.Text + "'" +
        //    ",@FundTypeIdListTxt=" + FundType + "" +
        //    ",@FundNameTxt=" + FundNameTxt + "" +
        //    ",@LegalEntityNameTxt=" + LegalEntityName + "" +
        //    ",@TypeListTxt='" + LoadType + "'" +
        //    ",@LockedBy='" + username + "'";



        string sqlstr = "SP_S_TNR_LOAD_LOCK_DYNAMO";
        //  DataSet DS = clsDB.getDataSet(sqlstr, 1800);

        DataSet DS = new DataSet();

        //foreach (KeyValuePair<string, string> pair in checksucess)
        //{
        //    string key = pair.Key.ToString();

        //    string Value = pair.Value.ToString();

        //    if (Value.ToLower() == "true")
        //    {
        //        if (msg1 != "")
        //            msg1 = msg1 + "," + key;
        //        else
        //            msg1 = key;
        //    }

        //    else
        //    {
        //        if (msg2 != "")
        //            msg2 = msg2 + "," + key;
        //        else
        //            msg2 = key;
        //    }


        //}



        #region Lock Data
        try
        {
            // DataSet ds = new DataSet("TimeRanges");
            string lsConnectionstring = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
            using (SqlConnection conn = new SqlConnection(lsConnectionstring))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = sqlstr;

                foreach (KeyValuePair<string, string> pair in checkLock)
                {
                    string param = string.Empty;

                    DataTable dt = new DataTable();

                    string key = string.Empty;

                    string value = string.Empty;

                    key = pair.Key;

                    if (key.ToLower() == "transaction")
                    {
                        param = "@Transaction";
                        dt = dsTransaction.Tables[0];
                    }

                    if (key.ToLower() == "commitment")
                    {
                        param = "@Commitment";
                        dt = dsPosition.Tables[0];
                    }

                    if (key.ToLower() == "summaryvaluation")
                    {
                        param = "@SumValuation";
                        dt = dsSummaryValuation.Tables[0];
                    }

                    if (key.ToLower() == "performance")
                    {
                        param = "@Performance";
                        dt = dsPerformance.Tables[0];
                    }

                    cmd.Parameters.AddWithValue(param, dt);
                }


                //if (bTransaction)
                //    cmd.Parameters.AddWithValue("@Transaction", dsTransaction.Tables[0]);
                //if (bPosition)
                //    cmd.Parameters.AddWithValue("@Commitment", dsPosition.Tables[0]);
                //if (bSummaryValuation)
                //    cmd.Parameters.AddWithValue("@SumValuation", dsSummaryValuation.Tables[0]);
                //if (bPerformance)
                //    cmd.Parameters.AddWithValue("@Performance", dsPerformance.Tables[0]);

                cmd.Parameters.AddWithValue("@LockedBy", username);

                cmd.Parameters.AddWithValue("@TypeListTxt", LoadType);

                cmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.SelectCommand.CommandTimeout = 1800;
                da.Fill(DS);
            }

            //return ds;
        }
        catch (Exception ex)
        {
            lblError.Text = sqlstr + " Failed to execute. " + ex.Message;
            lblError.Visible = true;
            // return null;
        }

        #endregion


        // DS = LockData(sqlstr, DSCommit.Tables[0], "@Commitment", DSTrans.Tables[0], "@Transaction", DSSummary.Tables[0], "@SumValuation", LoadType, username);


        if (DS != null)
        {
            DataTable Postn = DS.Tables[0];
            DataTable TRN = DS.Tables[1];
            DataTable PERF = DS.Tables[2];
            if (Postn.Rows.Count > 0 || TRN.Rows.Count > 0 || PERF.Rows.Count > 0)
            {
                LockData(DS);
                lblError.Text = "Data locked successfully.please check load log for details.";
                cleansessionvalue();
            }
            else
                lblError.Text = "No records found to lock. Please first run the load";
        }

        else
        {
            if (bTransaction)
                lblError.Text = "Data Lock Failed for transaction Please Contact to Administartor.";
            if (bPosition)
                lblError.Text = "Data Lock Failed for Commitment Please Contact to Administartor.";
            if (bSummaryValuation)
                lblError.Text = "Data Lock Failed for SummaryValuation Please Contact to Administartor.";
            if (bPerformance)
                lblError.Text = "Data Lock Failed for Performance Please Contact to Administartor.";
        }

    }

    private string GetSqlString(string Type, DataSet ds)
    {
        return GetSqlString(Type);
    }


    public void DataLock()
    {
       // string username = WindowsIdentity.GetCurrent().Name;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        string username = claimsIdentity.Name;
        sqlstr = "SP_S_TNR_LOAD_LOCK @LockedBy='" + username + "'";

        DataSet DS = clsDB.getDataSet(sqlstr, 1800);
        DataTable Postn = DS.Tables[0];
        DataTable TRN = DS.Tables[1];
        if (Postn.Rows.Count > 0 || TRN.Rows.Count > 0)
        {
            LockData(DS);
            lblError.Text = "Data locked successfully.please check load log for details.";
        }
        else
            lblError.Text = "No records found to lock. Please first run the load";
    }
    private string GetSqlString(string Type)
    {
        //        @DeleteFlg
        //1   -- DELETE RECORDS
        //0   -- NOT TO DELETE RECORDS.

        if (Type == Commitment)
        {
            //  sqlstr = "SP_S_Dynamo_COMMITMENT_LOAD @DeleteFlg=" + IsDelete + "";

            sqlstr = "SP_S_TNR_COMMITMENT_LOAD_DYNAMO";
        }
        else if (Type == Summary)
        {
            //sqlstr = "SP_S_Dynamo_SUMMARY_VALUATION_LOAD @StartDt='" + txtFrom.Text + "'" +
            //             ",@EndDt='" + txtTo.Text + "'" +
            //             ",@FundTypeIdListTxt=" + FundType + "" +
            //             ",@FundNames=" + FundNameTxt + "" +
            //             ",@LegalEntityNameTxt=" + LegalEntityName + "" +
            //             ",@DeleteFlg=" + IsDelete + "";

            sqlstr = "SP_S_TNR_SUMMARY_VALUATION_LOAD_DYNAMO";


        }
        else if (Type == Transaction)
        {
            // sqlstr = "SP_S_Dynamo_TRANSACTION_LOAD @DeleteFlg=" + IsDelete + "";
            //sqlstr = "EXEC SP_S_Dynamo_TRANSACTION_LOAD_DYNAMO @Transaction = "+ ds.Tables[0] + " ,@DeleteFlg='" + IsDelete + "'";
            sqlstr = "SP_S_TNR_TRANSACTION_LOAD_DYNAMO";
        }

        else if (Type == Performance)
        {
            // sqlstr = "SP_S_Dynamo_TRANSACTION_LOAD @DeleteFlg=" + IsDelete + "";
            //sqlstr = "EXEC SP_S_Dynamo_TRANSACTION_LOAD_DYNAMO @Transaction = "+ ds.Tables[0] + " ,@DeleteFlg='" + IsDelete + "'";
            sqlstr = "SP_S_TNR_PERFORMANCE_DYNAMO";
        }

        return sqlstr;
    }

    private bool CheckLockStatus()
    {
        bool status = false;
        bool CommP = false;
        bool ValsP = false;
        bool trn = false;
        bool perf = false;

        try
        {
            DataSet DSCommP = new DataSet();
            DataSet DSValsP = new DataSet();
            DataSet DStrn = new DataSet();
            //   DataSet dsPerformance = new DataSet();

            DataSet DSPerf = new DataSet();


            DataTable dtTrn = new DataTable();
            DataTable dtSummr = new DataTable();
            DataTable dtPostn = new DataTable();
            DataTable dtPerf = new DataTable();

            if (bPosition)
            {
                //sqlstr = GetSqlString(Commitment, "0", dsPosition);
                //DSCommP = clsDB.getDataSet(sqlstr, 1800);

                //if (DSCommP.Tables.Count > 0)
                //{
                //    if (DSCommP.Tables[1].Rows.Count > 0)
                //    {
                //        CommP = Convert.ToBoolean(DSCommP.Tables[1].Rows[0]["ReturnNmb"]);//ssi_lock
                //    }
                //    dtPostn = DSCommP.Tables[2];
                //}

                sqlstr = GetSqlString(Commitment);
                DSCommP = InsertData(sqlstr, dsPosition.Tables[0], "@Commitment", "@DeleteFlg", 0);

                if (DSCommP != null)
                {
                    if (DSCommP.Tables.Count > 0)
                    {
                        if (DSCommP.Tables[1].Rows.Count > 0)
                        {
                            CommP = Convert.ToBoolean(DSCommP.Tables[1].Rows[0]["ReturnNmb"]);//ssi_lock
                        }
                        dtPostn = DSCommP.Tables[2];
                        //dtPostn = DStrn.Tables[2];
                        //dtSummr = DStrn.Tables[2];
                    }


                }

            }

            if (bSummaryValuation)
            {
                //sqlstr = GetSqlString(Summary, "0", dsSummaryValuation);
                //DSValsP = clsDB.getDataSet(sqlstr, 1800);

                //if (DSValsP.Tables.Count > 0)
                //{
                //    if (DSValsP.Tables[1].Rows.Count > 0)
                //    {
                //        ValsP = Convert.ToBoolean(DSValsP.Tables[1].Rows[0]["ReturnNmb"]);//ssi_lock
                //    }
                //    dtSummr = DSValsP.Tables[2];
                //}


                sqlstr = GetSqlString(Summary);
                DSValsP = InsertData(sqlstr, dsSummaryValuation.Tables[0], "@SumValuation", "@DeleteFlg", 0);

                if (DSValsP != null)
                {
                    if (DSValsP.Tables.Count > 0)
                    {
                        if (DSValsP.Tables[1].Rows.Count > 0)
                        {
                            ValsP = Convert.ToBoolean(DSValsP.Tables[1].Rows[0]["ReturnNmb"]);//ssi_lock
                        }
                        dtSummr = DSValsP.Tables[2];
                        //dtPostn = DStrn.Tables[2];
                        //dtSummr = DStrn.Tables[2];
                    }

                }

                //else
                //{
                //    bcheckposition = true;
                //    //  bSummaryValuation = false;
                //    Session["bcheckposition"] = bcheckposition;
                //}

            }

            if (bTransaction)
            {
                //sqlstr = GetSqlString(Transaction, "0", dsTransaction);
                //DStrn = clsDB.getDataSet(sqlstr, 1800);
                sqlstr = GetSqlString(Transaction);
                DStrn = InsertData(sqlstr, dsTransaction.Tables[0], "@Transaction", "@DeleteFlg", 0);

                if (DStrn != null)
                {
                    if (DStrn.Tables.Count > 0)
                    {
                        if (DStrn.Tables[1].Rows.Count > 0)
                        {
                            trn = Convert.ToBoolean(DStrn.Tables[1].Rows[0]["ReturnNmb"]);//ssi_lock
                        }
                        dtTrn = DStrn.Tables[2];
                        //dtPostn = DStrn.Tables[2];
                        //dtSummr = DStrn.Tables[2];
                    }
                }
            }

            if (bPerformance)
            {
                //sqlstr = GetSqlString(Transaction, "0", dsTransaction);
                //DStrn = clsDB.getDataSet(sqlstr, 1800);
                sqlstr = GetSqlString(Performance);
                DSPerf = InsertData(sqlstr, dsPerformance.Tables[0], "@Performance", "@DeleteFlg", 0);
                if (DSPerf.Tables.Count > 0)
                {
                    if (DSPerf.Tables[1].Rows.Count > 0)
                    {
                        perf = Convert.ToBoolean(DSPerf.Tables[1].Rows[0]["ReturnNmb"]);//ssi_lock
                    }
                    dtPerf = DSPerf.Tables[2];
                    //dtPostn = DStrn.Tables[2];
                    //dtSummr = DStrn.Tables[2];
                }
            }



            ShowMissingSec(dtSummr, dtPostn, dtTrn, dtPerf);
            //DataTable d1 = new DataTable();
            //DataTable d2 = new DataTable();
            //ShowMissingSec(DSCommP.Tables[2], d1, d2);
            if (CommP == true || ValsP == true || trn == true || perf == true)
                status = true;
        }
        catch (Exception ex)
        {
            lg.AddinLogFile(Session["Filename"].ToString(), "Error " + ex.Message.ToString() + " " + DateTime.Now);
        }
        return status;
    }
    private void ShowMissingSec(DataTable summ, DataTable Comm, DataTable Trn, DataTable Perf)
    {
        try
        {
            if (summ.Rows.Count > 0 || Comm.Rows.Count > 0 || Trn.Rows.Count > 0 || Perf.Rows.Count > 0)
            {
                ViewState["NewSec"] = "true";
            }
            else
            {
                ViewState["NewSec"] = "false";
                return;
            }

            DataSet ds = new DataSet();

            summ.TableName = "Summary";
            summ.AcceptChanges();
            ds.Tables.Add(summ.Copy());


            Comm.TableName = "Commitment";
            Comm.AcceptChanges();
            ds.Tables.Add(Comm.Copy());

            Trn.TableName = "Transaction";
            Trn.AcceptChanges();
            ds.Tables.Add(Trn.Copy());


            Perf.TableName = "Performance";
            Perf.AcceptChanges();
            ds.Tables.Add(Perf.Copy());


            // DataSet ds = new DataSet();
            //ds.Tables.Add(summ.Copy());
            //ds.Tables.Add(Comm.Copy());
            //ds.Tables.Add(Trn.Copy());

            ds.AcceptChanges();
            Session.Add("NewDS", ds);
        }
        catch (Exception ex)
        {
            lblError.Text = lblError.Text + "\n" + "ERROR " + ex.Message.ToString();
        }
    }

    private bool ShowLinkforNewSec()
    {
        uplhide.Value = "1";
        bool status = false;
        if (Convert.ToString(ViewState["NewSec"]) == "true")
        {
            DataSet ds = new DataSet();
            ds = (DataSet)Session["NewDS"];

            if (ds.Tables[0].Rows.Count > 0 || ds.Tables[1].Rows.Count > 0 || ds.Tables[2].Rows.Count > 0 || ds.Tables[3].Rows.Count > 0)
            {
                trDownLoad.Style.Add("display", "inline");
                lnkDownLoad.Visible = true;

                trSubmit.Style.Add("display", "none");

                // btnLoadRun.Visible = false;
                status = true;
            }
        }
        return status;
    }

    private string GetListBoxItem(ListBox lst)
    {
        string text = string.Empty;
        for (int i = 0; i < lst.Items.Count; i++)
        {
            if (lst.Items[i].Selected)
            {
                text = text + "|" + lst.Items[i].Text.Replace("'", "''");
            }
        }
        if (text != "")
            text = text.Remove(0, 1);
        return text;
    }

    private bool LoadData()
    {
        #region Declaration

        #region Commented -10_30_2018
        //string LogFileName = "LogFile " + DateTime.Now;
        //LogFileName = LogFileName.Replace(":", "-");
        //LogFileName = LogFileName.Replace("/", "-");
        //sw = new StreamWriter(Request.PhysicalApplicationPath + "\\Log\\" + LogFileName + ".txt", true);

        #endregion



        //sw = new StreamWriter(LogFileName + ".txt", true);

        // string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
        //string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);

        // string CRM_constring = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";

        //string crmServerUrl = "http://Crm01/";
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        //string crmServerURL = "http://server:5555/";

        //string orgName = "GreshamPartners";
        //string orgName = "Webdev";

        string Userid = GetcurrentUser();
        bool bProceed = true;
        //string strDescription;
        //CrmService service = null;		
        IOrganizationService service = null;

        try
        {
            //service = GetCrmService(crmServerUrl, orgName,Userid);
            service = clsGM.GetCrmService();

            strDescription = "Crm Service starts successfully";
            LogMessage(service, strDescription, 62, "GeneralError");

            // sw.WriteLine("step 1 ");
            lg.AddinLogFile(Session["Filename"].ToString(), "step 1 " + DateTime.Now);
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            LogMessage(service, strDescription, 62, "GeneralError");
            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now);
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            LogMessage(service, strDescription, 62, "GeneralError");
            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now);
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
        int totalCountSummary = 0;
        int totalCountCommitment = 0;
        int totalCountPerformance = 0;

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
                if (bTransaction)
                {
                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    //greshamquery = GetSqlString(Transaction);
                    ////dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                    ////ds_gresham = new DataSet();
                    ////dagersham.SelectCommand.CommandTimeout = 1800;
                    ////dagersham.Fill(ds_gresham);
                    //ds_gresham = InsertData(greshamquery, dsTransaction.Tables[0], "@Transaction", "@DeleteFlg", 1);
                    //totalCount = ds_gresham.Tables[0].Rows.Count;
                    //totalCountTransaction = totalCount;
                    lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Transaction check for file format -------------------" + " " + DateTime.Now);
                    greshamquery = GetSqlString(Transaction);
                    ds_gresham = InsertData(greshamquery, dsTransaction.Tables[0], "@Transaction", "@DeleteFlg", 1);

                    if (ds_gresham != null)
                    {

                        #region Insert Transaction
                        try
                        {
                            // sw.WriteLine("---------------------------- New Transaction Insert Starts -------------------");
                            lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Transaction Insert Starts -------------------" + " " + DateTime.Now);
                            greshamquery = GetSqlString(Transaction);
                            //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            //ds_gresham = new DataSet();
                            //dagersham.SelectCommand.CommandTimeout = 1800;
                            //dagersham.Fill(ds_gresham);

                            ds_gresham = InsertData(greshamquery, dsTransaction.Tables[0], "@Transaction", "@DeleteFlg", 1);

                            //   if (ds_gresham ==null)


                            //#region Transaction Dataset blank if file format not correct

                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            totalCountTransaction = totalCount;
                            //Console.WriteLine("gresham Dataset Built for Transaction Insert ");

                            //sw.WriteLine("---------------------------- Transaction Delete Starts -------------------");
                            lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Transaction Delete Starts -------------------" + " " + DateTime.Now);
                            if (ds_gresham.Tables[2].Rows.Count > 0)
                                successCount = DeleteData(ds_gresham.Tables[2], "ssi_transactionlog", "ssi_transactionlogId", Userid);//Convert.ToInt32(ds_gresham.Tables[2].Rows[0]["DeleteCount"]);

                            strDescription = "Total Transaction Deleted: " + successCount;
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now);
                            LogMessage(service, strDescription, 62, "Dynamo Load");
                            // sw.WriteLine("---------------------------- Transaction Delete Ends  -------------------");
                            lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Transaction Delete Ends  -------------------" + " " + DateTime.Now);
                            successCount = 0;

                            // sw.WriteLine("Insert Starts for New Transaction on: " + DateTime.Now.ToString());
                            lg.AddinLogFile(Session["Filename"].ToString(), "Insert Starts for New Transaction on: " + DateTime.Now);
                        }
                        //catch (System.Web.Services.Protocols.SoapException exc)
                        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                        {
                            bProceed = true;
                            totalCount = 0;
                            transactionSuccess = false;
                            strDescription = "Transaction Insert failed, please contact administrator. Error Detail: " + exc.Detail.Message;
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + DateTime.Now);
                            LogMessage(service, strDescription, 62, "Dynamo Load");
                        }
                        catch (Exception exc)
                        {
                            bProceed = true;
                            totalCount = 0;
                            transactionSuccess = false;
                            strDescription = "Transaction Insert failed, please contact administrator. Error Detail: " + exc.Message;
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + DateTime.Now);
                            LogMessage(service, strDescription, 62, "Dynamo Load");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        // ssi_transactionlog objTransaction = new ssi_transactionlog();
                                        Entity objTransaction = new Entity("ssi_transactionlog");

                                        // objTransaction.ssi_transactionlogid = new Key();
                                        // objTransaction.ssi_transactionlogid.Value = Guid.NewGuid();


                                        //account
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) != "")
                                        {
                                            // objTransaction.ssi_accountid = new Lookup();
                                            // objTransaction.ssi_accountid.type = EntityName.ssi_account.ToString();
                                            // objTransaction.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));

                                            objTransaction["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"])));
                                        }

                                        //Security
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]) != "")
                                        {
                                            // objTransaction.ssi_securityid = new Lookup();
                                            // objTransaction.ssi_securityid.type = EntityName.ssi_security.ToString();
                                            // objTransaction.ssi_securityid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]));

                                            objTransaction["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"])));
                                        }

                                        //Assetclassid 
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]) != "")
                                        {
                                            // objTransaction.ssi_assetclassid = new Lookup();
                                            // objTransaction.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                            // objTransaction.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]));

                                            objTransaction["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"])));
                                        }

                                        //SectorId 
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]) != "")
                                        {
                                            // objTransaction.ssi_sectorid = new Lookup();
                                            // objTransaction.ssi_sectorid.type = EntityName.ssi_sector.ToString();
                                            // objTransaction.ssi_sectorid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]));

                                            objTransaction["ssi_sectorid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_sector", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"])));
                                        }

                                        ////quantity
                                        //if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Shares"]) != "" && Convert.ToString(ds_gresham.Tables[0].Rows[i]["Shares"]) != "0")
                                        //{
                                        //    // objTransaction.ssi_quantity = new CrmDecimal();
                                        //    // objTransaction.ssi_quantity.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Shares"]);

                                        //    objTransaction["ssi_quantity"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Shares"]);
                                        //}
                                        //else
                                        //{
                                        //    // objTransaction.ssi_quantity = new CrmDecimal();
                                        //    // objTransaction.ssi_quantity.IsNull = true;
                                        //    // objTransaction.ssi_quantity.IsNullSpecified = true;

                                        //    objTransaction["ssi_quantity"] = null;
                                        //}

                                        //Trade date
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Date"]) != "")
                                        {
                                            // objTransaction.ssi_tradedate = new CrmDateTime();
                                            // objTransaction.ssi_tradedate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Date"]);

                                            objTransaction["ssi_tradedate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["Transaction Date"]);
                                        }

                                        //Trade Amt
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Fund Amount"]) != "")
                                        {
                                            // objTransaction.ssi_value = new CrmMoney();
                                            // objTransaction.ssi_value.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Fund Amount"]);

                                            objTransaction["ssi_value"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Fund Amount"]));
                                        }
                                        else
                                        {
                                            // objTransaction.ssi_value = new CrmMoney();
                                            // objTransaction.ssi_value.IsNull = true;
                                            // objTransaction.ssi_value.IsNullSpecified = true;

                                            objTransaction["ssi_value"] = null;
                                        }

                                        //transactioncodeid
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_transactioncodeid"]) != "")
                                        {
                                            // objTransaction.ssi_transactioncodeid = new Lookup();
                                            // objTransaction.ssi_transactioncodeid.type = EntityName.ssi_transactionlog.ToString();
                                            // objTransaction.ssi_transactioncodeid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_transactioncodeid"]));

                                            objTransaction["ssi_transactioncodeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_transactionlog", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_transactioncodeid"])));
                                        }

                                        //ssi_lock
                                        if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]) == true)
                                        {
                                            // objTransaction.ssi_lock = new CrmBoolean();
                                            // objTransaction.ssi_lock.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]);

                                            objTransaction["ssi_lock"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lock"]).ToLower());
                                        }

                                        // Txn Num
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Txn Num"]) != "")
                                        {
                                            //objTransaction.ssi_Dynamotransactionnum = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Txn Num"]);
                                            objTransaction["ssi_tnrtransactionnum"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Txn Num"]);
                                        }

                                        // Transaction Type
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Type"]) != "")
                                        {
                                            //objTransaction.ssi_Dynamotransactiontype = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Type"]);
                                            objTransaction["ssi_tnrtransactiontype"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Type"]);
                                        }

                                        // Transaction Name - added _3_27_2019 (Basecamp Request)
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Name"]) != "")
                                        {
                                            //objTransaction.ssi_tnrtransactionname = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Name"]);
                                            objTransaction["ssi_tnrtransactionname"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Name"]);
                                        }

                                        // LastLockDate
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]) != "")
                                        {
                                            // objTransaction.ssi_lastlockdate = new CrmDateTime();
                                            // objTransaction.ssi_lastlockdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);

                                            objTransaction["ssi_lastlockdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);
                                        }

                                        if (Userid != "")
                                        {
                                            // objTransaction.createdby = new Lookup();
                                            // objTransaction.createdby.type = EntityName.systemuser.ToString();
                                            // objTransaction.createdby.Value = new Guid(Userid);

                                            objTransaction["createdby"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                                        }

                                        //SourceCode (Dynamo)
                                        // objTransaction.ssi_datasource = new Picklist();
                                        // objTransaction.ssi_datasource.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_datasourceid"]);

                                        objTransaction["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_datasourceid"]));
                                        //***********New Fields **********************//
                                        // Comment
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Comment"]) != "")
                                        {
                                            //objTransaction.ssi_comment = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Comment"]);
                                            objTransaction["ssi_comment"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Comment"]);
                                        }



                                        //Currency (default to USD)
                                        // objTransaction.transactioncurrencyid = new Lookup();
                                        // objTransaction.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                        // objTransaction.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");

                                        objTransaction["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));

                                        //below 3 cols added on 19 dec 2013
                                        //Assetclassid 
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]) != "")
                                        {
                                            // objTransaction.ssi_subassetclassid = new Lookup();
                                            // objTransaction.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                                            // objTransaction.ssi_subassetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]));

                                            objTransaction["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"])));
                                        }

                                        //greshamadvised (sectorflg)
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]) != "")
                                        {
                                            // objTransaction.ssi_benchmarkid = new Lookup();
                                            // objTransaction.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                                            // objTransaction.ssi_benchmarkid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"])); ;

                                            objTransaction["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"])));
                                        }

                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]) != "")
                                        {
                                            // objTransaction.ssi_grehamadvised = new CrmBoolean();
                                            // objTransaction.ssi_grehamadvised.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["sectorflg"]);

                                            objTransaction["ssi_grehamadvised"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
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
                                   // bProceed = false;
                                    failiureCount = failiureCount + 1;
                                    transactionSuccess = false;
                                    string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", Transaction Name: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Name"]);
                                    strDescription = failiureText + " Error Detail:" + exc.Detail.Message;
                                    lg.AddinLogFile(Session["Filename"].ToString(), strDescription + DateTime.Now);
                                    LogMessage(service, strDescription, 26, "Dynamo Load");
                                }
                                catch (Exception exc)
                                {
                                   // bProceed = false;
                                    transactionSuccess = false;
                                    failiureCount = failiureCount + 1;
                                    string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", Transaction Name: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction Name"]);
                                    strDescription = failiureText + " Error Detail:" + exc.Message;
                                    lg.AddinLogFile(Session["Filename"].ToString(), strDescription + DateTime.Now);
                                    LogMessage(service, strDescription, 26, "Dynamo Load");
                                }
                            }

                        // sw.WriteLine("Insert Ends for Transaction on: " + DateTime.Now.ToString());
                        lg.AddinLogFile(Session["Filename"].ToString(), "Insert Ends for Transaction on: " + DateTime.Now.ToString());
                        strDescription = "Total Transaction failed to insert: " + failiureCount;
                        LogMessage(service, strDescription, 40, "Dynamo Load");



                        strDescription = "Total Transaction inserted: " + successCount;
                        lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                        LogMessage(service, strDescription, 29, "Dynamo Load");
                        // sw.WriteLine("---------------------------- New Transaction Insert Ends -------------------");
                        // sw.WriteLine();
                        lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Transaction Insert Ends -------------------" + DateTime.Now.ToString());

                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        //#endregion


                        #endregion

                        bchecktransaction = true;
                        checksucess.Add("Transaction", "true");
                    }

                    else
                    {
                        bchecktransaction = false;

                        checksucess.Add("Transaction", "false");
                        //  checksucess[0] = "transaction";

                        //  bTransaction = false;
                        Session["bchecktransaction"] = bchecktransaction;
                    }


                }
                ///////////////////////////////// Position ///////////////////////////////////////////


                if (bSummaryValuation)
                {
                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Summary Valuation check file format Starts -------------------" + " " + DateTime.Now);
                        greshamquery = GetSqlString(Summary);
                        //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        //ds_gresham = new DataSet();
                        //dagersham.SelectCommand.CommandTimeout = 1800;
                        //dagersham.Fill(ds_gresham);
                        ds_gresham = InsertData(greshamquery, dsSummaryValuation.Tables[0], "@SumValuation", "@DeleteFlg", 1);

                        if (ds_gresham != null)
                        {


                            ////////////////////////////// Insert Position Starts Summary Valuation ////////////////////////////////////////////
                            #region Insert Position Summary Valuation
                            try
                            {
                                //greshamquery = GetSqlString(Summary, dsSummaryValuation);
                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                //ds_gresham = new DataSet();
                                //dagersham.SelectCommand.CommandTimeout = 1800;
                                //dagersham.Fill(ds_gresham);
                                //totalCount = ds_gresham.Tables[0].Rows.Count;
                                //totalCountSummary = totalCount;


                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Summary Valuation Insert Starts -------------------" + " " + DateTime.Now);
                                greshamquery = GetSqlString(Summary);
                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                //ds_gresham = new DataSet();
                                //dagersham.SelectCommand.CommandTimeout = 1800;
                                //dagersham.Fill(ds_gresham);
                                ds_gresham = InsertData(greshamquery, dsSummaryValuation.Tables[0], "@SumValuation", "@DeleteFlg", 1);
                                totalCount = ds_gresham.Tables[0].Rows.Count;
                                totalCountSummary = totalCount;



                                // sw.WriteLine("---------------------------- Position(Summary) Delete Starts -------------------");
                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Position(Summary) Delete Starts -------------------");
                                if (ds_gresham.Tables[2].Rows.Count > 0)
                                    successCount = DeleteData(ds_gresham.Tables[2], "ssi_position", "ssi_PositionId", Userid);//Convert.ToInt32(ds_gresham.Tables[2].Rows[0]["DeleteCount"]);

                                strDescription = "Total Position(Summary) Deleted: " + successCount;
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                                // sw.WriteLine("---------------------------- Position(Summary) Delete Ends  -------------------");
                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Position(Summary) Delete Ends  -------------------");
                                successCount = 0;

                                // sw.WriteLine("---------------------------- New Position Insert Starts -------------------");
                                // sw.WriteLine("Insert Starts for New Position on: " + DateTime.Now.ToString());
                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Position Insert Starts -------------------");
                                lg.AddinLogFile(Session["Filename"].ToString(), "Insert Starts for New Position on: " + DateTime.Now.ToString());
                            }
                            //catch (System.Web.Services.Protocols.SoapException exc)
                            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Detail.Message;
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Message;
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            }

                            if (bProceed == true)
                                for (int i = 0; i < totalCount; i++)
                                {
                                    try
                                    {
                                        if (bProceed == true)
                                        {
                                            //ssi_position objPosition = new ssi_position();
                                            Entity objPosition = new Entity("ssi_position");

                                            //primary key ssi_positionid
                                            // objPosition.ssi_positionid = new Key();
                                            // objPosition.ssi_positionid.Value = Guid.NewGuid();

                                            //accountid
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) != "")
                                            {
                                                // objPosition.ssi_accountid = new Lookup();
                                                // objPosition.ssi_accountid.type = EntityName.ssi_account.ToString();
                                                // objPosition.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));

                                                objPosition["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"])));
                                            }

                                            //SecurityId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]) != "")
                                            {
                                                // objPosition.ssi_securityid = new Lookup();
                                                // objPosition.ssi_securityid.type = EntityName.ssi_security.ToString();
                                                // objPosition.ssi_securityid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]));

                                                objPosition["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"])));
                                            }

                                            //sas_assetclassid
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]) != "")
                                            {
                                                // objPosition.ssi_assetclassid = new Lookup();
                                                // objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                                // objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]));

                                                objPosition["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"])));
                                            }

                                            //SectorId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]) != "")
                                            {
                                                // objPosition.ssi_sectorid = new Lookup();
                                                // objPosition.ssi_sectorid.type = EntityName.ssi_sector.ToString();
                                                // objPosition.ssi_sectorid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]));

                                                objPosition["ssi_sectorid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_sector", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"])));
                                            }

                                            //FundId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]) != "")
                                            {
                                                // objPosition.ssi_fundid = new Lookup();
                                                // objPosition.ssi_fundid.type = EntityName.ssi_fund.ToString();
                                                // objPosition.ssi_fundid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]));

                                                objPosition["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"])));
                                            }

                                            //Quantity
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Valuation"]) != "")
                                            {
                                                // objPosition.ssi_quantity = new CrmDecimal();
                                                // objPosition.ssi_quantity.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]);

                                                objPosition["ssi_quantity"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_quantity = new CrmDecimal();
                                                // objPosition.ssi_quantity.IsNull = true;
                                                // objPosition.ssi_quantity.IsNullSpecified = true;

                                                objPosition["ssi_quantity"] = null;
                                            }

                                            // AS OF DATE / PriceDate
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]) != "")
                                            {
                                                // objPosition.ssi_asofdate = new CrmDateTime();
                                                // objPosition.ssi_asofdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);

                                                objPosition["ssi_asofdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                            }

                                            //Market Value
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Valuation"]) != "")
                                            {
                                                // objPosition.ssi_marketvalue = new CrmMoney();
                                                // objPosition.ssi_marketvalue.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]);

                                                objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]));
                                            }
                                            else
                                            {
                                                // objPosition.ssi_marketvalue = new CrmMoney();
                                                // objPosition.ssi_marketvalue.IsNull = true;
                                                // objPosition.ssi_marketvalue.IsNullSpecified = true;

                                                objPosition["ssi_marketvalue"] = null;
                                            }

                                            //CommitmentSummaryFlg,  
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]) == true)
                                            {
                                                // objPosition.ssi_commitmentsummaryflg = new CrmBoolean();
                                                // objPosition.ssi_commitmentsummaryflg.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]);

                                                objPosition["ssi_commitmentsummaryflg"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]).ToLower());
                                            }

                                            //ssi_lock
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]) == true)
                                            {
                                                // objPosition.ssi_lock = new CrmBoolean();
                                                // objPosition.ssi_lock.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]);

                                                objPosition["ssi_lock"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lock"]).ToLower());
                                            }

                                            //ssi_committedcapital
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Commited Capital"]) != "")
                                            {
                                                // objPosition.ssi_committedcapital = new CrmDecimal();
                                                // objPosition.ssi_committedcapital.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Commited Capital"]);

                                                objPosition["ssi_committedcapital"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Commited Capital"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_committedcapital = new CrmDecimal();
                                                // objPosition.ssi_committedcapital.IsNull = true;
                                                // objPosition.ssi_committedcapital.IsNullSpecified = true;

                                                objPosition["ssi_committedcapital"] = null;
                                            }

                                            //Remaining Capital
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Remaining Capital"]) != "")
                                            {
                                                // objPosition.ssi_remainingcapital = new CrmDecimal();
                                                // objPosition.ssi_remainingcapital.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Remaining Capital"]);

                                                objPosition["ssi_remainingcapital"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Remaining Capital"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_remainingcapital = new CrmDecimal();
                                                // objPosition.ssi_remainingcapital.IsNull = true;
                                                // objPosition.ssi_remainingcapital.IsNullSpecified = true;

                                                objPosition["ssi_remainingcapital"] = null;
                                            }

                                            //Current Month Paid In
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Current Month Paid In"]) != "")
                                            {
                                                // objPosition.ssi_currentmonthpaidin = new CrmDecimal();
                                                // objPosition.ssi_currentmonthpaidin.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Current Month Paid In"]);

                                                objPosition["ssi_currentmonthpaidin"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Current Month Paid In"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_currentmonthpaidin = new CrmDecimal();
                                                // objPosition.ssi_currentmonthpaidin.IsNull = true;
                                                // objPosition.ssi_currentmonthpaidin.IsNullSpecified = true;

                                                objPosition["ssi_currentmonthpaidin"] = null;
                                            }

                                            //Current Month Distribution
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Current Month Distribution"]) != "")
                                            {
                                                // objPosition.ssi_currentmonthdistribution = new CrmDecimal();
                                                // objPosition.ssi_currentmonthdistribution.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Current Month Distribution"]);

                                                objPosition["ssi_currentmonthdistribution"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Current Month Distribution"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_currentmonthdistribution = new CrmDecimal();
                                                // objPosition.ssi_currentmonthdistribution.IsNull = true;
                                                // objPosition.ssi_currentmonthdistribution.IsNullSpecified = true;

                                                objPosition["ssi_currentmonthdistribution"] = null;
                                            }

                                            // LastLockDate
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]) != "")
                                            {
                                                // objPosition.ssi_lastlockdate = new CrmDateTime();
                                                // objPosition.ssi_lastlockdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);

                                                objPosition["ssi_lastlockdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);
                                            }

                                            if (Userid != "")
                                            {
                                                // objPosition.createdby = new Lookup();
                                                // objPosition.createdby.type = EntityName.systemuser.ToString();
                                                // objPosition.createdby.Value = new Guid(Userid);

                                                objPosition["createdby"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                                            }

                                            //SourceCode (Dynamo)
                                            // objPosition.ssi_datasource = new Picklist();
                                            // objPosition.ssi_datasource.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_datasourceid"]);

                                            objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_datasourceid"]));

                                            // Currency ( default to USD )
                                            // objPosition.transactioncurrencyid = new Lookup();
                                            // objPosition.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                            // objPosition.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");

                                            objPosition["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));

                                            //below 3 col added on 19 dec 2013
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]) != "")
                                            {
                                                // objPosition.ssi_subassetclassid = new Lookup();
                                                // objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                                                // objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]));

                                                objPosition["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"])));
                                            }

                                            //SubAssetclassid 
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]) != "")
                                            {
                                                // objPosition.ssi_benchmarkid = new Lookup();
                                                // objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                                                // objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]));

                                                objPosition["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"])));
                                            }

                                            //greshamadvised (sectorflg)
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]) != "")
                                            {
                                                // objPosition.ssi_greshamadvised = new CrmBoolean();
                                                // objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["sectorflg"]);

                                                objPosition["ssi_greshamadvised"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                            }




                                            //paid in to date
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["PaidInToDate"]) != "")
                                            {

                                                objPosition["ssi_paidintodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["PaidInToDate"]);
                                            }


                                            //distributed to date

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["DistrToDate"]) != "")
                                            {

                                                objPosition["ssi_distributedtodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["DistrToDate"]);
                                            }


                                            //Currency
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Currency"]) != "")
                                            {

                                                // objPosition["ssi_currency"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Currency"]));


                                                objPosition["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Currency"])));
                                            }


                                            //FX Divisor

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["FxAdvisor"]) != "")
                                            {

                                                objPosition["ssi_fxdivisor"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["FxAdvisor"]);
                                            }


                                            //Original Committed capital

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgCommCapital"]) != "")
                                            {

                                                objPosition["ssi_originalcommittedcapital"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgCommCapital"]);
                                            }


                                            //Original Valuation

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgValuation"]) != "")
                                            {

                                                objPosition["ssi_originalvaluation"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgValuation"]);
                                            }


                                            //Original Remaining Currency
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgRemCurrency"]) != "")
                                            {
                                                objPosition["ssi_originalremainingcurrency"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgRemCurrency"]);
                                            }

                                            //Original Paid in To Date
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgPaidInDate"]) != "")
                                            {

                                                objPosition["ssi_originalpaidintodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgPaidInDate"]);
                                            }

                                            //Original Distributions to Date
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgDistrInDate"]) != "")
                                            {

                                                objPosition["ssi_originaldistributionstodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgDistrInDate"]);
                                            }


                                            service.Create(objPosition);
                                            //Thread.Sleep(sleepTime);
                                            successCount = successCount + 1;
                                        }
                                        else
                                            break;
                                    }
                                    //catch (System.Web.Services.Protocols.SoapException exc)
                                    catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                                    {
                                       // bProceed = false;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                        strDescription = failiureText + " Error Detail:" + exc.Detail.Message;  //"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position " + failiureText + " Error Detail: " + exc.Detail.InnerText;
                                        LogMessage(service, strDescription, 27, "Dynamo Load");
                                        lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                    }
                                    catch (Exception exc)
                                    {
                                       // bProceed = false;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                        strDescription = failiureText + " Error Detail: " + exc.Message;//"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position" + failiureText + " Error Detail: " + exc.Message;
                                        LogMessage(service, strDescription, 27, "Dynamo Load");
                                        lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                    }
                                }

                            // sw.WriteLine("Insert Ends for Position(Summary Valuation) on: " + DateTime.Now.ToString());
                            lg.AddinLogFile(Session["Filename"].ToString(), "Insert Ends for Position(Summary Valuation) on: " + " " + DateTime.Now.ToString());
                            strDescription = "Total Position(Summary Valuation) failed to insert: " + failiureCount;
                            LogMessage(service, strDescription, 43, "Dynamo Load");

                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());

                            strDescription = "Total Position(Summary Valuation) inserted: " + successCount;
                            LogMessage(service, strDescription, 28, "Dynamo Load");
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Position Insert Ends -------------------" + " " + DateTime.Now.ToString());
                            // sw.WriteLine("---------------------------- New Position Insert Ends -------------------");
                            // sw.WriteLine();
                            ds_gresham.Dispose();
                            dagersham.Dispose();

                            #endregion
                            ////////////////////////////////// Insert Position  Ends Summary Valuation //////////////////////////////////////////

                            bcheckposition = true;
                            checksucess.Add("SummaryValuation", "true");
                        }
                        else
                        {
                            bcheckposition = false;
                            checksucess.Add("SummaryValuation", "false");
                            Session["bcheckposition"] = bcheckposition;
                        }
                    }
                }
                if (bPosition)
                {
                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Commitment check File Format Starts -------------------" + " " + DateTime.Now);
                        greshamquery = GetSqlString(Commitment);
                        //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        //ds_gresham = new DataSet();
                        //dagersham.SelectCommand.CommandTimeout = 1800;
                        //dagersham.Fill(ds_gresham);
                        ds_gresham = InsertData(greshamquery, dsPosition.Tables[0], "@Commitment", "@DeleteFlg", 1);

                        if (ds_gresham != null)
                        {

                            ////////////////////////////// Insert Position Starts Commitment ////////////////////////////////////////////
                            #region Insert Position Commitment
                            try
                            {
                                //greshamquery = GetSqlString(Commitment, dsPosition);
                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                //ds_gresham = new DataSet();
                                //dagersham.SelectCommand.CommandTimeout = 1800;
                                //dagersham.Fill(ds_gresham);
                                //totalCount = ds_gresham.Tables[0].Rows.Count;
                                //totalCountCommitment = totalCount;


                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Commitment Insert Starts -------------------" + " " + DateTime.Now);
                                greshamquery = GetSqlString(Commitment);
                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                //ds_gresham = new DataSet();
                                //dagersham.SelectCommand.CommandTimeout = 1800;
                                //dagersham.Fill(ds_gresham);
                                ds_gresham = InsertData(greshamquery, dsPosition.Tables[0], "@Commitment", "@DeleteFlg", 1);
                                totalCount = ds_gresham.Tables[0].Rows.Count;
                                totalCountCommitment = totalCount;

                                // sw.WriteLine("---------------------------- Position(Commitment) Delete Starts -------------------");
                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Position(Commitment) Delete Starts -------------------" + " " + DateTime.Now.ToString());
                                if (ds_gresham.Tables[2].Rows.Count > 0)
                                    successCount = DeleteData(ds_gresham.Tables[2], "ssi_position", "ssi_PositionId", Userid);//Convert.ToInt32(ds_gresham.Tables[2].Rows[0]["DeleteCount"]);

                                strDescription = "Total Position(Commitment) Deleted: " + successCount;
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                // sw.WriteLine("---------------------------- Position(Commitment) Delete Ends  -------------------");
                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Position(Commitment) Delete Ends  -------------------" + " " + DateTime.Now.ToString());
                                successCount = 0;


                                // sw.WriteLine("---------------------------- New Position Insert Starts -------------------");
                                // sw.WriteLine("Insert Starts for New Position on: " + DateTime.Now.ToString());

                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Position Insert Starts -------------------" + " " + DateTime.Now.ToString());
                                lg.AddinLogFile(Session["Filename"].ToString(), "Insert Starts for New Position on:" + " " + DateTime.Now.ToString());


                            }
                            //catch (System.Web.Services.Protocols.SoapException exc)
                            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Detail.Message;
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Message;
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                            }

                            if (bProceed == true)
                                for (int i = 0; i < totalCount; i++)
                                {
                                    try
                                    {
                                        if (bProceed == true)
                                        {
                                            // ssi_position objPosition = new ssi_position();
                                            Entity objPosition = new Entity("ssi_position");

                                            //primary key ssi_positionid
                                            // objPosition.ssi_positionid = new Key();
                                            // objPosition.ssi_positionid.Value = Guid.NewGuid();

                                            //accountid
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) != "")
                                            {
                                                // objPosition.ssi_accountid = new Lookup();
                                                // objPosition.ssi_accountid.type = EntityName.ssi_account.ToString();
                                                // objPosition.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));

                                                objPosition["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"])));
                                            }

                                            //SecurityId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]) != "")
                                            {
                                                // objPosition.ssi_securityid = new Lookup();
                                                // objPosition.ssi_securityid.type = EntityName.ssi_security.ToString();
                                                // objPosition.ssi_securityid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]));

                                                objPosition["ssi_securityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_security", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"])));
                                            }

                                            //sas_assetclassid
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]) != "")
                                            {
                                                // objPosition.ssi_assetclassid = new Lookup();
                                                // objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                                // objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]));

                                                objPosition["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"])));
                                            }

                                            //SectorId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]) != "")
                                            {
                                                // objPosition.ssi_sectorid = new Lookup();
                                                // objPosition.ssi_sectorid.type = EntityName.ssi_sector.ToString();
                                                // objPosition.ssi_sectorid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"]));

                                                objPosition["ssi_sectorid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_sector", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_sectorid"])));
                                            }

                                            //FundId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]) != "")
                                            {
                                                // objPosition.ssi_fundid = new Lookup();
                                                // objPosition.ssi_fundid.type = EntityName.ssi_fund.ToString();
                                                // objPosition.ssi_fundid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]));

                                                objPosition["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"])));
                                            }

                                            //Quantity
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Valuation"]) != "")
                                            {
                                                // objPosition.ssi_quantity = new CrmDecimal();
                                                // objPosition.ssi_quantity.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]);

                                                objPosition["ssi_quantity"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_quantity = new CrmDecimal();
                                                // objPosition.ssi_quantity.IsNull = true;
                                                // objPosition.ssi_quantity.IsNullSpecified = true;

                                                objPosition["ssi_quantity"] = null;
                                            }

                                            // AS OF DATE / PriceDate
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]) != "")
                                            {
                                                // objPosition.ssi_asofdate = new CrmDateTime();
                                                // objPosition.ssi_asofdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);

                                                objPosition["ssi_asofdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                            }

                                            //Market Value
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Valuation"]) != "")
                                            {
                                                // objPosition.ssi_marketvalue = new CrmMoney();
                                                // objPosition.ssi_marketvalue.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]);

                                                objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Valuation"]));
                                            }
                                            else
                                            {
                                                // objPosition.ssi_marketvalue = new CrmMoney();
                                                // objPosition.ssi_marketvalue.IsNull = true;
                                                // objPosition.ssi_marketvalue.IsNullSpecified = true;

                                                objPosition["ssi_marketvalue"] = null;
                                            }

                                            //CommitmentSummaryFlg,  
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]) == true)
                                            {
                                                // objPosition.ssi_commitmentsummaryflg = new CrmBoolean();
                                                // objPosition.ssi_commitmentsummaryflg.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]);

                                                objPosition["ssi_commitmentsummaryflg"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]).ToLower());
                                            }

                                            //ssi_lock
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]) == true)
                                            {
                                                // objPosition.ssi_lock = new CrmBoolean();
                                                // objPosition.ssi_lock.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]);

                                                objPosition["ssi_lock"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lock"]).ToLower());
                                            }

                                            //ssi_percentownership
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]) != "")
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPosition["ssi_percentownership"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPosition["ssi_percentownership"] = null;
                                            }

                                            //ssi_percentownership
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Commitment"]) != "")
                                            {
                                                // objPosition.ssi_commitment = new CrmDecimal();
                                                // objPosition.ssi_commitment.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Commitment"]);

                                                objPosition["ssi_commitment"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Commitment"]);
                                            }
                                            else
                                            {
                                                // objPosition.ssi_commitment = new CrmDecimal();
                                                // objPosition.ssi_commitment.IsNull = true;
                                                // objPosition.ssi_commitment.IsNullSpecified = true;

                                                objPosition["ssi_commitment"] = null;
                                            }

                                            // LastLockDate
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]) != "")
                                            {
                                                // objPosition.ssi_lastlockdate = new CrmDateTime();
                                                // objPosition.ssi_lastlockdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);

                                                objPosition["ssi_lastlockdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);
                                            }

                                            if (Userid != "")
                                            {
                                                // objPosition.createdby = new Lookup();
                                                // objPosition.createdby.type = EntityName.systemuser.ToString();
                                                // objPosition.createdby.Value = new Guid(Userid);

                                                objPosition["createdby"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                                            }
                                            //SourceCode (Dynamo)
                                            // objPosition.ssi_datasource = new Picklist();
                                            // objPosition.ssi_datasource.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_datasourceid"]);

                                            objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_datasourceid"]));

                                            // Currency ( default to USD )
                                            // objPosition.transactioncurrencyid = new Lookup();
                                            // objPosition.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                            // objPosition.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                                            objPosition["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));


                                            //below 3 col added on 19 dec 2013
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]) != "")
                                            {
                                                // objPosition.ssi_subassetclassid = new Lookup();
                                                // objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                                                // objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]));

                                                objPosition["ssi_subassetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"])));
                                            }

                                            //SubAssetclassid 
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]) != "")
                                            {
                                                // objPosition.ssi_benchmarkid = new Lookup();
                                                // objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                                                // objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]));

                                                objPosition["ssi_benchmarkid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"])));
                                            }

                                            //greshamadvised (sectorflg)
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]) != "")
                                            {
                                                // objPosition.ssi_greshamadvised = new CrmBoolean();
                                                // objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["sectorflg"]);

                                                objPosition["ssi_greshamadvised"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                            }



                                            //paid in to date
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["PaidInToDate"]) != "")
                                            {

                                                objPosition["ssi_paidintodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["PaidInToDate"]);
                                            }


                                            //distributed to date

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["DistrToDate"]) != "")
                                            {

                                                objPosition["ssi_distributedtodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["DistrToDate"]);
                                            }


                                            //Currency
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Currency"]) != "")
                                            {

                                                // objPosition["ssi_currency"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["Currency"]));


                                                objPosition["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Currency"])));
                                            }


                                            //FX Divisor

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["FxAdvisor"]) != "")
                                            {

                                                objPosition["ssi_fxdivisor"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["FxAdvisor"]);
                                            }


                                            //Original Committed capital

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgCommCapital"]) != "")
                                            {

                                                objPosition["ssi_originalcommittedcapital"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgCommCapital"]);
                                            }


                                            //Original Valuation

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgValuation"]) != "")
                                            {

                                                objPosition["ssi_originalvaluation"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgValuation"]);
                                            }


                                            //Original Remaining Currency
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgRemCurrency"]) != "")
                                            {
                                                objPosition["ssi_originalremainingcurrency"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgRemCurrency"]);
                                            }

                                            //Original Paid in To Date
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgPaidInDate"]) != "")
                                            {

                                                objPosition["ssi_originalpaidintodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgPaidInDate"]);
                                            }

                                            //Original Distributions to Date
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OrgDistrInDate"]) != "")
                                            {

                                                objPosition["ssi_originaldistributionstodate"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OrgDistrInDate"]);
                                            }
                                            //CurrentMonthContributions
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["CurrentMonthContributions"]) != "")
                                            {

                                                objPosition["ssi_currentmonthpaidin"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["CurrentMonthContributions"]);
                                            }

                                            //CurrentMonthContributions
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["RemainingCapital"]) != "")
                                            {

                                                objPosition["ssi_remainingcapital"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["RemainingCapital"]);
                                            }

                                            //CurrentMonthContributions
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["CurrentMonthDistributions"]) != "")
                                            {

                                                objPosition["ssi_currentmonthdistribution"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["CurrentMonthDistributions"]);
                                            }


                                            //OriginalCurrentMonthContributions
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OriginalCurrentMonthContributions"]) != "")
                                            {

                                                objPosition["ssi_originalcurrentmonthcontributions"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OriginalCurrentMonthContributions"]));
                                            }

                                            //ssi_OriginalCurrentMonthDistributions
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OriginalCurrentMonthDistributions"]) != "")
                                            {

                                                objPosition["ssi_originalcurrentmonthdistributions"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["OriginalCurrentMonthDistributions"]));
                                            }

                                            service.Create(objPosition);
                                            //Thread.Sleep(sleepTime);
                                            successCount = successCount + 1;
                                        }
                                        else
                                            break;
                                    }
                                    //catch (System.Web.Services.Protocols.SoapException exc)
                                    catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                                    {
                                        //bProceed = false;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                        strDescription = failiureText + " Error Detail:" + exc.Detail.Message;  //"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position " + failiureText + " Error Detail: " + exc.Detail.InnerText;
                                        lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                        LogMessage(service, strDescription, 27, "Dynamo Load");
                                    }
                                    catch (Exception exc)
                                    {
                                       // bProceed = false;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                        strDescription = failiureText + " Error Detail: " + exc.Message;//"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position" + failiureText + " Error Detail: " + exc.Message;
                                        lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                        LogMessage(service, strDescription, 27, "Dynamo Load");
                                    }
                                }

                            // sw.WriteLine("Insert Ends for Position(Commitment) on: " + DateTime.Now.ToString());
                            lg.AddinLogFile(Session["Filename"].ToString(), "Insert Ends for Position(Commitment) on:" + " " + DateTime.Now.ToString());
                            strDescription = "Total Position(Commitment) failed to insert: " + failiureCount;
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            LogMessage(service, strDescription, 43, "Dynamo Load");

                            strDescription = "Total Position(Commitment) inserted: " + successCount;
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            LogMessage(service, strDescription, 28, "Dynamo Load");
                            // sw.WriteLine("---------------------------- New Position Insert Ends -------------------");
                            // sw.WriteLine();
                            lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Position Insert Ends -------------------" + " " + DateTime.Now.ToString());
                            ds_gresham.Dispose();
                            dagersham.Dispose();

                            #endregion
                            ////////////////////////////////// Insert Position  Ends Commitment//////////////////////////////////////////

                            bcheckcommit = true;
                            checksucess.Add("Commitment", "true");
                        }
                        else
                        {
                            //  bcheckcommit = false;

                            // checksucess[1] = "Commitment";
                            checksucess.Add("Commitment", "false");

                            Session["bcheckcommit"] = bcheckcommit;
                        }
                    }
                }


                ///performance----------------------

                if (bPerformance)
                {
                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New performance check File Format Starts -------------------" + " " + DateTime.Now);
                        greshamquery = GetSqlString(Performance);
                        //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        //ds_gresham = new DataSet();
                        //dagersham.SelectCommand.CommandTimeout = 1800;
                        //dagersham.Fill(ds_gresham);
                        ds_gresham = InsertData(greshamquery, dsPerformance.Tables[0], "@Performance", "@DeleteFlg", 1);

                        if (ds_gresham != null)
                        {

                            ////////////////////////////// Insert Performance Starts  ////////////////////////////////////////////
                            #region Insert Performance
                            try
                            {
                                //greshamquery = GetSqlString(Commitment, dsPosition);
                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                //ds_gresham = new DataSet();
                                //dagersham.SelectCommand.CommandTimeout = 1800;
                                //dagersham.Fill(ds_gresham);
                                //totalCount = ds_gresham.Tables[0].Rows.Count;
                                //totalCountCommitment = totalCount;


                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Performance Insert Starts -------------------" + " " + DateTime.Now);
                                greshamquery = GetSqlString(Performance);
                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                //ds_gresham = new DataSet();
                                //dagersham.SelectCommand.CommandTimeout = 1800;
                                //dagersham.Fill(ds_gresham);
                                ds_gresham = InsertData(greshamquery, dsPerformance.Tables[0], "@Performance", "@DeleteFlg", 1);
                                totalCount = ds_gresham.Tables[0].Rows.Count;
                                totalCountPerformance = totalCount;

                                // sw.WriteLine("---------------------------- Position(Commitment) Delete Starts -------------------");
                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Performance Delete Starts -------------------" + " " + DateTime.Now.ToString());
                                if (ds_gresham.Tables[2].Rows.Count > 0)
                                    successCount = DeleteData(ds_gresham.Tables[2], "sas_publicperformance", "Sas_publicperformanceId", Userid);//Convert.ToInt32(ds_gresham.Tables[2].Rows[0]["DeleteCount"]);

                                strDescription = "Total Performance Deleted: " + successCount;
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                // sw.WriteLine("---------------------------- Position(Commitment) Delete Ends  -------------------");
                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Performance Delete Ends  -------------------" + " " + DateTime.Now.ToString());
                                successCount = 0;


                                // sw.WriteLine("---------------------------- New Position Insert Starts -------------------");
                                // sw.WriteLine("Insert Starts for New Position on: " + DateTime.Now.ToString());

                                lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Performance Insert Starts -------------------" + " " + DateTime.Now.ToString());
                                lg.AddinLogFile(Session["Filename"].ToString(), "Insert Starts for New Position on:" + " " + DateTime.Now.ToString());


                            }
                            //catch (System.Web.Services.Protocols.SoapException exc)
                            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Performance Insert failed, please contact administrator. Error Detail:" + exc.Detail.Message;
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Performance Insert failed, please contact administrator. Error Detail:" + exc.Message;
                                lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                LogMessage(service, strDescription, 62, "Dynamo Load");
                            }

                            if (bProceed == true)
                                for (int i = 0; i < totalCount; i++)
                                {
                                    try
                                    {
                                        if (bProceed == true)
                                        {
                                            // ssi_position objPosition = new ssi_position();
                                            Entity objPerformance = new Entity("sas_publicperformance");

                                            //primary key ssi_positionid
                                            // objPosition.ssi_positionid = new Key();
                                            // objPosition.ssi_positionid.Value = Guid.NewGuid();


                                            //sas_assetclassid
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["AssetClassUUID"]) != "")
                                            {
                                                // objPosition.ssi_assetclassid = new Lookup();
                                                // objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                                // objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]));

                                                objPerformance["ssi_assetclassid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_assetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["AssetClassUUID"])));
                                            }



                                            //FundId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["FundUUID"]) != "")
                                            {
                                                // objPosition.ssi_fundid = new Lookup();
                                                // objPosition.ssi_fundid.type = EntityName.ssi_fund.ToString();
                                                // objPosition.ssi_fundid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]));

                                                objPerformance["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["FundUUID"])));
                                            }


                                            //ssi_householdid
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["HouseholdUUID"]) != "")
                                            {
                                                // objPosition.ssi_fundid = new Lookup();
                                                // objPosition.ssi_fundid.type = EntityName.ssi_fund.ToString();
                                                // objPosition.ssi_fundid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]));

                                                objPerformance["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["HouseholdUUID"])));
                                            }


                                            //LegalEntityUUID
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["LegalEntityUUID"]) != "")
                                            {
                                                // objPosition.ssi_fundid = new Lookup();
                                                // objPosition.ssi_fundid.type = EntityName.ssi_fund.ToString();
                                                // objPosition.ssi_fundid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_fundid"]));

                                                objPerformance["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["LegalEntityUUID"])));
                                            }


                                            // sas_startdate

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["StartDate"]) != "")
                                            {
                                                // objPosition.ssi_asofdate = new CrmDateTime();
                                                // objPosition.ssi_asofdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);

                                                objPerformance["sas_startdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["StartDate"]);
                                            }

                                            else
                                            {
                                                objPerformance["sas_startdate"] = null;
                                            }


                                            // sas_enddate
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["AsOfDate"]) != "")
                                            {
                                                // objPosition.ssi_asofdate = new CrmDateTime();
                                                // objPosition.ssi_asofdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);

                                                objPerformance["sas_enddate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["AsOfDate"]);
                                            }

                                            else
                                            {
                                                objPerformance["sas_enddate"] = null;
                                            }


                                            // InceptionDate
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["InceptionDate"]) != "")
                                            {
                                                // objPosition.ssi_asofdate = new CrmDateTime();
                                                // objPosition.ssi_asofdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);

                                                objPerformance["ssi_inceptiondate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["InceptionDate"]);
                                            }

                                            else
                                            {
                                                objPerformance["ssi_inceptiondate"] = null;
                                            }


                                            //FundType,  
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["FundType"]) == true)
                                            {
                                                // objPosition.ssi_commitmentsummaryflg = new CrmBoolean();
                                                // objPosition.ssi_commitmentsummaryflg.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]);

                                                objPerformance["ssi_greshamfund"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["FundType"]).ToLower());
                                            }

                                            //else
                                            //{
                                            //    objPerformance["ssi_greshamfund"] = null;
                                            //}


                                            //Mature,  
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["Mature"]) == true)
                                            {
                                                // objPosition.ssi_commitmentsummaryflg = new CrmBoolean();
                                                // objPosition.ssi_commitmentsummaryflg.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["CommitmentSummaryFlg"]);

                                                objPerformance["ssi_mature"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Mature"]).ToLower());
                                            }

                                            //else
                                            //{
                                            //    objPerformance["ssi_mature"] = null;
                                            //}

                                            //ssi_lock
                                            if (Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]) == true)
                                            {
                                                // objPosition.ssi_lock = new CrmBoolean();
                                                // objPosition.ssi_lock.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["ssi_lock"]);

                                                objPerformance["ssi_lock"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lock"]).ToLower());
                                            }

                                            //else
                                            //{
                                            //    objPerformance["ssi_lock"] = null;
                                            //}

                                            //sas_performance

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["performance"]) != "")//MTD
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["sas_performance"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["performance"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["sas_performance"] = null;
                                            }


                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["YTD"]) != "")//YTD
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_ytd"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["YTD"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_ytd"] = null;
                                            }


                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["QTD"]) != "")//QTD
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_qtd"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["QTD"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_qtd"] = null;
                                            }

                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ITD"]) != "")//QTD
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_itd"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["ITD"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_itd"] = null;
                                            }



                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["1Yr"]) != "")//1Yr
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_1yr"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["1Yr"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_1yr"] = null;
                                            }


                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["3Yr"]) != "")//3Yr
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_3yr"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["3Yr"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_3yr"] = null;
                                            }



                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["5Yr"]) != "")//5Yr
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_5yr"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["5Yr"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_5yr"] = null;
                                            }



                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["10Yr"]) != "")//10Yr//added on 04/04/2019 
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_10yr"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["10Yr"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_10yr"] = null;
                                            }



                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["TVPI"]) != "")//TVPI--added on 04/04/2019 
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["% Of Total Fund"]);

                                                objPerformance["ssi_tvpi"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["TVPI"]);
                                            }

                                            else
                                            {
                                                // objPosition.ssi_percentownership = new CrmDecimal();
                                                // objPosition.ssi_percentownership.IsNull = true;
                                                // objPosition.ssi_percentownership.IsNullSpecified = true;

                                                objPerformance["ssi_tvpi"] = null;
                                            }





                                            // LastLockDate
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]) != "")
                                            {
                                                // objPosition.ssi_lastlockdate = new CrmDateTime();
                                                // objPosition.ssi_lastlockdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);

                                                objPerformance["ssi_lastlockdate"] = Convert.ToDateTime(ds_gresham.Tables[0].Rows[i]["ssi_lastlockdate"]);
                                            }


                                            // Maturity added on 04/18/2019
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Maturity"]) != "")
                                            {
                                                objPerformance["ssi_maturity"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Maturity"]);
                                            }

                                            else
                                            {
                                                objPerformance["ssi_maturity"] = null;
                                            }

                                            // objBillingInvoice["ssi_adjustmentreason"] = Convert.ToString(txtAdjReason.Text);

                                            if (Userid != "")
                                            {
                                                // objPosition.createdby = new Lookup();
                                                // objPosition.createdby.type = EntityName.systemuser.ToString();
                                                // objPosition.createdby.Value = new Guid(Userid);

                                                objPerformance["createdby"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                                            }
                                            //SourceCode (Dynamo)
                                            // objPosition.ssi_datasource = new Picklist();
                                            // objPosition.ssi_datasource.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["ssi_datasourceid"]);

                                            objPerformance["ssi_source"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(100000002));

                                            // Currency ( default to USD )
                                            // objPosition.transactioncurrencyid = new Lookup();
                                            // objPosition.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                            // objPosition.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                                            //  objPosition["transactioncurrencyid"] = new Microsoft.Xrm.Sdk.EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));

                                            service.Create(objPerformance);
                                            //Thread.Sleep(sleepTime);
                                            successCount = successCount + 1;
                                        }
                                        else
                                            break;
                                    }
                                    //catch (System.Web.Services.Protocols.SoapException exc)
                                    catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                                    {
                                       // bProceed = false;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        // string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                        strDescription = " Error Detail:" + exc.Detail.Message;  //"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position " + failiureText + " Error Detail: " + exc.Detail.InnerText;
                                        lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                        LogMessage(service, strDescription, 27, "Dynamo Load");
                                    }
                                    catch (Exception exc)
                                    {
                                       // bProceed = false;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        //  string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                        strDescription = " Error Detail: " + exc.Message;//"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position" + failiureText + " Error Detail: " + exc.Message;
                                        lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                        LogMessage(service, strDescription, 27, "Dynamo Load");
                                    }
                                }

                            // sw.WriteLine("Insert Ends for Position(Commitment) on: " + DateTime.Now.ToString());
                            lg.AddinLogFile(Session["Filename"].ToString(), "Insert Ends for Performance on:" + " " + DateTime.Now.ToString());
                            strDescription = "Total Performance failed to insert: " + failiureCount;
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            LogMessage(service, strDescription, 43, "Dynamo Load");

                            strDescription = "Total Performance inserted: " + successCount;
                            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            LogMessage(service, strDescription, 28, "Dynamo Load");
                            // sw.WriteLine("---------------------------- New Position Insert Ends -------------------");
                            // sw.WriteLine();
                            lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- New Performance Insert Ends -------------------" + " " + DateTime.Now.ToString());
                            ds_gresham.Dispose();
                            dagersham.Dispose();

                            #endregion
                            ////////////////////////////////// Insert Position  Ends Commitment//////////////////////////////////////////

                            //   bcheckcommit = true;

                            bcheckperformance = true;
                            checksucess.Add("performance", "true");
                        }
                        else
                        {
                            //bcheckcommit = false;
                            checksucess.Add("performance", "false");
                            bcheckperformance = false;

                            //  Session["bcheckcommit"] = bcheckcommit;
                        }
                    }
                }

                ////////////////////////////////// Postion End ///////////////////////////////////////

            }

            // sw.Flush();
            // sw.Close();

            #endregion
            Session.Remove("NewDS");

        }

        if (totalCountTransaction == 0 && totalCountSummary == 0 && totalCountCommitment == 0 && totalCountPerformance == 0)
            return false;
        else
            return true;
    }
    public DataSet InsertData(string vSqlQuery, DataTable dt, string parameter1, string parameter2, int deleteFlg)
    {
        try
        {
            DataSet ds = new DataSet("TimeRanges");
            string lsConnectionstring = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
            using (SqlConnection conn = new SqlConnection(lsConnectionstring))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = vSqlQuery;
                cmd.Parameters.AddWithValue(parameter1, dt);
                cmd.Parameters.AddWithValue(parameter2, deleteFlg);

                //cmd.Parameters.Add("@ERROR", SqlDbType.Char, 500);
                //cmd.Parameters["@ERROR"].Direction = ParameterDirection.Output;  

                cmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.SelectCommand.CommandTimeout = 1800;
                da.Fill(ds);

                //int code = (int)cmd.Parameters["@ERROR"].Value;

                //if (code == 1)
                //{
                //    lblError.Text = "Contact Adminstartor.";

                //}

            }

            return ds;
        }
        catch (Exception ex)
        {
            lblError.Text = lblError.Text + "\n" + vSqlQuery + " Failed to execute. " + ex.Message;
            lblError.Visible = true;
            return null;
        }
    }

    public DataSet LockData(string vSqlQuery, DataTable dt, string parameter1, DataTable dt1, string parameter2, DataTable dt2, string parameter3, DataTable dt3, string parameter4, string Loadtype, string LockedBy)
    {
        try
        {
            DataSet ds = new DataSet("TimeRanges");
            string lsConnectionstring = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
            using (SqlConnection conn = new SqlConnection(lsConnectionstring))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = vSqlQuery;
                cmd.Parameters.AddWithValue(parameter1, dt);
                cmd.Parameters.AddWithValue(parameter2, dt1);
                cmd.Parameters.AddWithValue(parameter3, dt2);
                cmd.Parameters.AddWithValue(parameter4, dt3);

                cmd.Parameters.AddWithValue("@LockedBy", LockedBy);

                cmd.Parameters.AddWithValue("@TypeListTxt", Loadtype);

                cmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.SelectCommand.CommandTimeout = 1800;
                da.Fill(ds);
            }

            return ds;
        }
        catch (Exception ex)
        {
            lblError.Text = vSqlQuery + " Failed to execute. " + ex.Message;
            lblError.Visible = true;
            return null;
        }
    }

    private void LockData(DataSet DS)
    {
        #region Declaration
        //string LogFileName = "LockLogFile " + DateTime.Now;
        //LogFileName = LogFileName.Replace(":", "-");
        //LogFileName = LogFileName.Replace("/", "-");
        //sw = new StreamWriter(Request.PhysicalApplicationPath + "\\Log\\" + LogFileName + ".txt", true);

        // string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
        //string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);

        ///string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);

        //string orgName = "GreshamPartners";
        bool bProceed = true;
        string strDescription;
        string Userid = GetcurrentUser();
        //CrmService service = null;
        IOrganizationService service = null;

        try
        {
            //service = GetCrmService(crmServerUrl, orgName, Userid);
            service = clsGM.GetCrmService();

            strDescription = "Crm Service starts successfully";
            LogMessage(service, strDescription, 62, "GeneralError");
            //lg.AddinLogFile("", strDescription + " " + DateTime.Now.ToString());
            // sw.WriteLine("step 1 ");
            // lg.AddinLogFile("", "step 1 " + " " + DateTime.Now.ToString());
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            LogMessage(service, strDescription, 62, "GeneralError");
            //lg.AddinLogFile("", strDescription + " " + DateTime.Now.ToString());
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            LogMessage(service, strDescription, 62, "GeneralError");
            // lg.AddinLogFile("", strDescription + " " + DateTime.Now.ToString());
        }

        //service.PreAuthenticate = true;

        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);

        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();

        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        int totalCount = 0;
        int successCount = 0;
        int failiureCount = 0;

        ds_gresham = DS;

        #endregion

        if (bProceed == true)
        {
            // using data load from file
            #region Transaction Position Lock Data


            successCount = 0;
            failiureCount = 0;
            // bProceed = true;
            bool positionSuccess = true;
            bool transactionSuccess = true;
            execType = "B";
            if (bProceed == true)
            {
                if (bTransaction)
                {
                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    #region Lock Transaction
                    try
                    {
                        //  sw.WriteLine("---------------------------- Transaction Lock Starts -------------------");
                        lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Transaction Lock Starts -------------------" + " " + DateTime.Now.ToString());
                        //greshamquery = GetSqlString(Transaction);
                        //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        //ds_gresham = new DataSet();
                        //dagersham.SelectCommand.CommandTimeout = 1800;
                        // dagersham.Fill(ds_gresham);
                        totalCount = ds_gresham.Tables[1].Rows.Count;
                        //sw.WriteLine("Transaction Lock Starts on: " + DateTime.Now.ToString());
                        // lg.AddinLogFile(Session["Filename"].ToString(), "Transaction Lock Starts on: " + " " + DateTime.Now.ToString());
                    }
                    //catch (System.Web.Services.Protocols.SoapException exc)
                    catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                    {
                        bProceed = true;
                        totalCount = 0;
                        transactionSuccess = false;
                        strDescription = "Transaction Lock failed, please contact administrator. Error Detail: " + exc.Detail.Message;
                        LogMessage(service, strDescription, 62, "Dynamo Load");
                        // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                    }
                    catch (Exception exc)
                    {
                        bProceed = true;
                        totalCount = 0;
                        transactionSuccess = false;
                        strDescription = "Transaction Lock failed, please contact administrator. Error Detail: " + exc.Message;
                        LogMessage(service, strDescription, 62, "Dynamo Load");
                        // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                    }

                    if (bProceed == true)
                        for (int i = 0; i < totalCount; i++)
                        {
                            try
                            {
                                if (bProceed == true)
                                {
                                    //ssi_transactionlog objTransaction = new ssi_transactionlog();
                                    Entity objTransaction = new Entity("ssi_transactionlog");

                                    //Guid TransactionId = new Guid(Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_transactionlogid"]));

                                    // objTransaction.ssi_transactionlogid = new Key();
                                    // objTransaction.ssi_transactionlogid.Value = TransactionId;

                                    objTransaction["ssi_transactionlogid"] = new Guid(Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_transactionlogid"]));

                                    //Locked
                                    // objTransaction.ssi_lock = new CrmBoolean();
                                    // objTransaction.ssi_lock.Value = Convert.ToBoolean(true);

                                    objTransaction["ssi_lock"] = true;


                                    //LastLock date

                                    // objTransaction.ssi_lastlockdate = new CrmDateTime();
                                    // objTransaction.ssi_lastlockdate.Value = Convert.ToString(DateTime.Now);

                                    objTransaction["ssi_lastlockdate"] = DateTime.Now;

                                    //objTransaction.ssi_lockedBy = Convert.ToString(WindowsIdentity.GetCurrent().Name);

                                    service.Update(objTransaction);
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
                                string failiureText = "TransactionId:" + Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_transactionlogid"]);
                                strDescription = failiureText + " Error Detail:" + exc.Detail.Message;
                                LogMessage(service, strDescription, 26, "Dynamo Load");
                                //  lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                transactionSuccess = false;
                                failiureCount = failiureCount + 1;
                                string failiureText = "TransactionId:" + Convert.ToString(ds_gresham.Tables[1].Rows[i]["ssi_transactionlogid"]);
                                strDescription = failiureText + " Error Detail:" + exc.Message;
                                LogMessage(service, strDescription, 26, "Dynamo Load");
                                // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                            }
                        }

                    //  sw.WriteLine("Lock Ends for Transaction on: " + DateTime.Now.ToString());
                    //lg.AddinLogFile(Session["Filename"].ToString(), "Lock Ends for Transaction on: " + " " + DateTime.Now.ToString());


                    strDescription = "Total Transaction failed to Lock: " + failiureCount;
                    LogMessage(service, strDescription, 40, "Dynamo Load");
                    //lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());

                    strDescription = "Total Transaction Locked: " + successCount;
                    LogMessage(service, strDescription, 29, "Dynamo Load");
                    //lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                    //lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Transaction Lock Ends -------------------" + " " + DateTime.Now.ToString());


                    //  sw.WriteLine("---------------------------- Transaction Lock Ends -------------------");
                    // sw.WriteLine();

                    ds_gresham.Dispose();
                    dagersham.Dispose();

                    #endregion
                }
                ///////////////////////////////// Position ///////////////////////////////////////////

                successCount = 0;
                failiureCount = 0;
                totalCount = 0;

                if (bProceed == true)
                {
                    if (bPosition || bSummaryValuation)
                    {
                        ////////////////////////////// Insert Position Starts ////////////////////////////////////////////
                        #region Lock Position
                        try
                        {
                            //greshamquery = GetSqlString(Summary);
                            //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            //ds_gresham = new DataSet();
                            //dagersham.SelectCommand.CommandTimeout = 1800;
                            //dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            //sw.WriteLine("---------------------------- Position Lock Starts -------------------");
                            // sw.WriteLine("Lock Starts for  Position on: " + DateTime.Now.ToString());
                            // lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Position Lock Starts -------------------" + " " + DateTime.Now.ToString());
                            // lg.AddinLogFile(Session["Filename"].ToString(), "Lock Starts for  Position on: " + " " + DateTime.Now.ToString());
                        }
                        //catch (System.Web.Services.Protocols.SoapException exc)
                        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            positionSuccess = false;
                            strDescription = "Position Lock failed, please contact administrator. Error Detail:" + exc.Detail.Message;
                            LogMessage(service, strDescription, 62, "Dynamo Load");
                            // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            positionSuccess = false;
                            strDescription = "Position Lock failed, please contact administrator. Error Detail:" + exc.Message;
                            LogMessage(service, strDescription, 62, "Dynamo Load");
                            // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        //ssi_position objPosition = new ssi_position();
                                        Entity objPosition = new Entity("ssi_position");

                                        //Guid PositionId = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_positionid"]));

                                        // objPosition.ssi_positionid = new Key();
                                        // objPosition.ssi_positionid.Value = PositionId;

                                        objPosition["ssi_positionid"] = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_positionid"]));

                                        //Locked
                                        // objPosition.ssi_lock = new CrmBoolean();
                                        // objPosition.ssi_lock.Value = Convert.ToBoolean(true);

                                        objPosition["ssi_lock"] = true;


                                        //LastLock date

                                        // objPosition.ssi_lastlockdate = new CrmDateTime();
                                        // objPosition.ssi_lastlockdate.Value = Convert.ToString(DateTime.Now);

                                        objPosition["ssi_lastlockdate"] = DateTime.Now;

                                        //objTransaction.ssi_lockedBy = Convert.ToString(WindowsIdentity.GetCurrent().Name);

                                        service.Update(objPosition);
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
                                    positionSuccess = false;
                                    failiureCount = failiureCount + 1;
                                    string failiureText = "PositionId:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_positionid"]);
                                    strDescription = failiureText + " Error Detail:" + exc.Detail.Message;  //"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position " + failiureText + " Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(service, strDescription, 27, "Dynamo Load");
                                    // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                }
                                catch (Exception exc)
                                {
                                    bProceed = false;
                                    positionSuccess = false;
                                    failiureCount = failiureCount + 1;
                                    string failiureText = "PositionId:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_positionid"]);
                                    strDescription = failiureText + " Error Detail: " + exc.Message;//"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position" + failiureText + " Error Detail: " + exc.Message;
                                    LogMessage(service, strDescription, 27, "Dynamo Load");
                                    // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                }
                            }

                        // sw.WriteLine("Lock Ends for Position on: " + DateTime.Now.ToString());
                        //   lg.AddinLogFile(Session["Filename"].ToString(), "Lock Ends for Position on: " + " " + DateTime.Now.ToString());

                        strDescription = "Total Position failed to Lock: " + failiureCount;
                        LogMessage(service, strDescription, 43, "Dynamo Load");
                        // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());

                        strDescription = "Total Position Locked: " + successCount;
                        LogMessage(service, strDescription, 28, "Dynamo Load");
                        //sw.WriteLine("---------------------------- Position Lock Ends -------------------");
                        // sw.WriteLine();

                        //  lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                        //  lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Position Lock Ends -------------------" + " " + DateTime.Now.ToString());

                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        #endregion
                        ////////////////////////////////// Insert Position  Ends Summary Valuation //////////////////////////////////////////
                    }
                }
                ////////////////////////////////// Postion End ///////////////////////////////////////



                if (bProceed == true)
                {
                    if (bPerformance)
                    {
                        ////////////////////////////// Insert Position Starts ////////////////////////////////////////////
                        #region Lock Performance
                        try
                        {
                            //greshamquery = GetSqlString(Summary);
                            //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            //ds_gresham = new DataSet();
                            //dagersham.SelectCommand.CommandTimeout = 1800;
                            //dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[2].Rows.Count;
                            //sw.WriteLine("---------------------------- Position Lock Starts -------------------");
                            // sw.WriteLine("Lock Starts for  Position on: " + DateTime.Now.ToString());
                            //  lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Performance Lock Starts -------------------" + " " + DateTime.Now.ToString());
                            //  lg.AddinLogFile(Session["Filename"].ToString(), "Lock Starts for  Performance on: " + " " + DateTime.Now.ToString());
                        }
                        //catch (System.Web.Services.Protocols.SoapException exc)
                        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            positionSuccess = false;
                            strDescription = "Performance Lock failed, please contact administrator. Error Detail:" + exc.Detail.Message;
                            LogMessage(service, strDescription, 62, "Dynamo Load");
                            // lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            positionSuccess = false;
                            strDescription = "Performance Lock failed, please contact administrator. Error Detail:" + exc.Message;
                            LogMessage(service, strDescription, 62, "Dynamo Load");
                            //lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        //ssi_position objPosition = new ssi_position();
                                        Entity objPerformance = new Entity("sas_publicperformance");

                                        //Guid PositionId = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_positionid"]));

                                        // objPosition.ssi_positionid = new Key();
                                        // objPosition.ssi_positionid.Value = PositionId;

                                        objPerformance["sas_publicperformanceid"] = new Guid(Convert.ToString(ds_gresham.Tables[2].Rows[i]["Sas_publicperformanceId"]));

                                        //Locked
                                        // objPosition.ssi_lock = new CrmBoolean();
                                        // objPosition.ssi_lock.Value = Convert.ToBoolean(true);

                                        objPerformance["ssi_lock"] = true;


                                        //LastLock date

                                        // objPosition.ssi_lastlockdate = new CrmDateTime();
                                        // objPosition.ssi_lastlockdate.Value = Convert.ToString(DateTime.Now);

                                        objPerformance["ssi_lastlockdate"] = DateTime.Now;

                                        //objTransaction.ssi_lockedBy = Convert.ToString(WindowsIdentity.GetCurrent().Name);

                                        service.Update(objPerformance);
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
                                    positionSuccess = false;
                                    failiureCount = failiureCount + 1;
                                    string failiureText = "sas_publicperformanceId:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Sas_publicperformanceId"]);
                                    strDescription = failiureText + " Error Detail:" + exc.Detail.Message;  //"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position " + failiureText + " Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(service, strDescription, 27, "Dynamo Load");
                                    //   lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                }
                                catch (Exception exc)
                                {
                                    bProceed = false;
                                    positionSuccess = false;
                                    failiureCount = failiureCount + 1;
                                    string failiureText = "Sas_publicperformanceId:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Sas_publicperformanceId"]);
                                    strDescription = failiureText + " Error Detail: " + exc.Message;//"Insert failed for Commitment Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position" + failiureText + " Error Detail: " + exc.Message;
                                    LogMessage(service, strDescription, 27, "Dynamo Load");
                                    //  lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                                }
                            }

                        // sw.WriteLine("Lock Ends for Position on: " + DateTime.Now.ToString());
                        // lg.AddinLogFile(Session["Filename"].ToString(), "Lock Ends for Performance on: " + " " + DateTime.Now.ToString());

                        strDescription = "Total Performance failed to Lock: " + failiureCount;
                        LogMessage(service, strDescription, 43, "Dynamo Load");
                        //  lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());

                        strDescription = "Total Performance Locked: " + successCount;
                        LogMessage(service, strDescription, 28, "Dynamo Load");
                        //sw.WriteLine("---------------------------- Position Lock Ends -------------------");
                        // sw.WriteLine();

                        //  lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
                        //  lg.AddinLogFile(Session["Filename"].ToString(), "---------------------------- Performance Lock Ends -------------------" + " " + DateTime.Now.ToString());

                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        #endregion
                        ////////////////////////////////// Insert Position  Ends Summary Valuation //////////////////////////////////////////
                    }
                }
            }

            //   sw.Flush();
            //   sw.Close();

            #endregion

        }

    }

    // private int DeleteData(DataTable dt, EntityName entityName, string ColumnName, string UserId,StreamWriter sw)
    private int DeleteData(DataTable dt, string entityName, string ColumnName, string UserId)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        //string orgName = "GreshamPartners";
        //CrmService service = null;		
        IOrganizationService service = null;

        int successcount = 0;
        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            //bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
            // sw.WriteLine(strDescription);
            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
        }
        catch (Exception exc)
        {
            //bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
            // sw.WriteLine(strDescription);
            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                Guid UUID = new Guid(Convert.ToString(dt.Rows[j][ColumnName]));
                service.Delete(entityName.ToString(), UUID);
                successcount = successcount + 1;
            }
        }
        catch (Exception ex)
        {
            Response.Write("<br/>" + ex.Message);
            // sw.WriteLine(strDescription);
            lg.AddinLogFile(Session["Filename"].ToString(), strDescription + " " + DateTime.Now.ToString());
            LogMessage(service, strDescription, 62, "Dynamo Load");
        }
        return successcount;
    }

    // /// <summary>
    // /// Set up the CRM Service.
    // /// </summary>
    // /// <param name="organizationName">My Organization</param>
    // /// <returns>CrmService configured with AD Authentication</returns>
    // public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    // {
    // // Get the CRM Users appointments
    // // Setup the Authentication Token
    // CrmAuthenticationToken token = new CrmAuthenticationToken();
    // token.AuthenticationType = 0; // Use Active Directory authentication.
    // token.OrganizationName = organizationName;
    // //string username = WindowsIdentity.GetCurrent().Name;

    // //if (username == "CORP\\gbhagia")
    // //{
    // //    // Use the global user ID of the system user that is to be impersonated.
    // //    token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
    // //    //token.CallerId = new Guid("C42C7E05-8303-DE11-A38C-001D09665E8F");//gary                
    // //}
    // if (CallerId != "")
    // token.CallerId = new Guid(CallerId);
    // CrmService service = new CrmService();

    // if (crmServerUrl != null &&
    // crmServerUrl.Length > 0)
    // {
    // UriBuilder builder = new UriBuilder(crmServerUrl);
    // builder.Path = "//MSCRMServices//2007//CrmService.asmx";
    // service.Url = builder.Uri.ToString();
    // }

    // service.CrmAuthenticationTokenValue = token;
    // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

    // return service;
    // }

    private void LogMessage(IOrganizationService service, string strDescription, int intIssueType, string strFileLoading)
    {
        try
        {

            //ssi_loadlog objLoadLog = new ssi_loadlog();
            Entity objLoadLog = new Entity("ssi_loadlog");
            //objLoadLog.ssi_name =Convert.ToString(DateTime.Today);

            //objLoadLog.ssi_date = new CrmDateTime();
            //objLoadLog.ssi_date.Value = DateTime.Now.ToString();
            objLoadLog["ssi_date"] = DateTime.Now;


            //objLoadLog.ssi_fileloading = strFileLoading;
            objLoadLog["ssi_fileloading"] = strFileLoading;

            //objLoadLog.ssi_descriptionofissue = strDescription;
            objLoadLog["ssi_descriptionofissue"] = strDescription;

            //objLoadLog.ssi_typeofissue = new Picklist();
            //objLoadLog.ssi_typeofissue.Value = intIssueType;
            objLoadLog["ssi_typeofissue"] = new Microsoft.Xrm.Sdk.OptionSetValue(intIssueType);

            service.Create(objLoadLog);
        }
        catch (Exception exc)
        {
            //HttpContext.Current.Response.Write(exc.Message.ToString());
            //  sw.WriteLine(exc.Message);
            lg.AddinLogFile(HttpContext.Current.Session["Filename"].ToString(), exc.Message.ToString());
            //  sw.Flush();
            // sw.Close();
          //  throw;
        }
    }
    protected void lnkDownLoad_Click(object sender, EventArgs e)
    {
        #region Commented OLD CODE
        //String lsFileNamforFinalXls = "NewAccountAndSecurity_" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";

        //string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\NewAccountAndSecurity.xlsx");
        //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);
        //string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls.Replace("xlsx", "xml"));
        //string FilePath = (Server.MapPath("") + @"\ExcelTemplate\NewAccountAndSecurity.xlsx");

        //FileInfo loFile = new FileInfo(strDirectory1);
        //loFile.CopyTo(strDirectory, true);

        //DataSet ds = new DataSet();
        //ds = (DataSet)Session["NewDS"];

        ////export datatable to excel
        //Workbook workbook = new Workbook();
        //workbook.LoadFromFile(strDirectory);
        //for (int i = 0; i < ds.Tables.Count; i++)
        //{
        //    if (i == 2)
        //    {
        //        Worksheet sheet = workbook.Worksheets[i];
        //    workbook.Version = ExcelVersion.Version97to2003;
        //    sheet.InsertDataTable(ds.Tables[i], true, 1, 1, -1, -1);
        //    sheet.Name = ds.Tables[i].TableName;
        //  //  if (ds.Tables[i].Rows.Count > 0)

        //        sheet.AllocatedRange.AutoFitColumns();
        //        sheet.AllocatedRange.AutoFitRows();
        //        //sheet.Rows[0].RowHeight = 20;
        //    }
        //}

        //workbook.SaveAsXml(strDirectory2);
        //workbook = null;
        //XmlDocument xmlDoc = new XmlDocument();
        //xmlDoc.Load(strDirectory2);
        //XmlElement businessEntities = xmlDoc.DocumentElement;
        //XmlNode loNode = businessEntities.LastChild;
        //XmlNode loNode1 = businessEntities.FirstChild;
        //businessEntities.RemoveChild(loNode);


        //xmlDoc.Save(strDirectory2);
        //xmlDoc = null;
        //loFile = null;
        //loFile = new FileInfo(strDirectory);
        //loFile.Delete();
        //loFile = new FileInfo(strDirectory2);
        //loFile.CopyTo(strDirectory, true);
        //loFile = null;
        ////loFile = new FileInfo(strDirectory2);
        ////loFile.Delete();

        //#region New xls to xlsx code
        //Workbook workbook1 = new Workbook();
        //workbook1.LoadFromXml(strDirectory2);
        //workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);


        //loFile = new FileInfo(strDirectory2);
        //loFile.Delete();
        //loFile = null;

        //#region PageSetup
        //Workbook workbook2 = new Workbook();
        //workbook2.LoadFromFile(strDirectory);
        //Worksheet sheet1 = workbook2.Worksheets[0];
        //workbook2.SaveToFile(strDirectory, ExcelVersion.Version2016);

        //var setup = sheet1.PageSetup;
        //setup.FitToPagesWide = 1;
        ////setup.FitToPagesTall = 1;
        //setup.IsFitToPage = true;
        //setup.PaperSize = PaperSizeType.PaperA4;
        //setup.Orientation = PageOrientationType.Landscape;
        //setup.FitToPagesWide = 1;
        //setup.FitToPagesTall = 0;
        //setup.CenterHorizontally = true;
        //setup.CenterVertically = false;
        //workbook2.SaveToFile(strDirectory, ExcelVersion.Version2016);
        //#endregion

        ////  lsFileNamforFinalXls = "/ExcelTemplate/TempFolder/" + lsFileNamforFinalXls;
        //#endregion
        #endregion

        #region Spire License Code
        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion

        String lsFileNamforFinalXls = "NewAccountAndSecurity_" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
        string ExcelSavePath = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);
        DataSet ds = new DataSet();
        ds = (DataSet)Session["NewDS"];

        Workbook book = new Workbook();
        book.CreateEmptySheets(4);
        for (int i = 0; i < ds.Tables.Count; i++)
        {
            Worksheet sheet = book.Worksheets[i];
            sheet.Name = ds.Tables[i].TableName;
            var setup = sheet.PageSetup;
            setup.FitToPagesWide = 1;
            //setup.FitToPagesTall = 1;
            setup.IsFitToPage = true;
            setup.PaperSize = PaperSizeType.PaperA4;
            setup.Orientation = PageOrientationType.Landscape;
            setup.FitToPagesWide = 1;
            setup.FitToPagesTall = 0;
            setup.CenterHorizontally = true;
            setup.CenterVertically = false;
            if (ds.Tables[i].Rows.Count > 0)
            {
                sheet.Range[1, 1, 1, ds.Tables[i].Columns.Count].Style.Font.IsBold = true;

                sheet.InsertDataTable(ds.Tables[i], true, 1, 1);
                sheet.Range[1, 1, ds.Tables[i].Rows.Count + 1, ds.Tables[i].Columns.Count + 1].AutoFitColumns();
                sheet.Range[1, 1, ds.Tables[i].Rows.Count + 1, ds.Tables[i].Columns.Count + 1].Style.HorizontalAlignment = HorizontalAlignType.Center;
            }
            //SheetNo++;
        }

        book.SaveToFile(ExcelSavePath, ExcelVersion.Version2016);

        // lsFileNamforFinalXls = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + lsFileNamforFinalXls);
        Response.TransmitFile(ExcelSavePath);
        Response.End();
    }
    protected void btnCancel_Click(object sender, EventArgs e)
    {
        cleansessionvalue();
        mvLoad.ActiveViewIndex = 0;
        lblError.Text = "";

        trDownLoad.Style.Add("display", "none");
        trSubmit.Style.Add("display", "inline");
    }

    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
      //  string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        string strName = string.Empty;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
         strName = claimsIdentity.Name;



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
    protected void lstLE_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
    }
    protected void lstFund_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
    }

    protected void Button1_Click(object sender, EventArgs e)
    {

    }
    protected void linktransaction_Click(object sender, EventArgs e)
    {
        string SampleFilePath = Server.MapPath("") + @"\ExcelTemplate\" + "Transactions.xlsx";
        Download_File(SampleFilePath, "Transactions.xlsx");
    }
    protected void linksumaary_Click(object sender, EventArgs e)
    {
        string SampleFilePath = Server.MapPath("") + @"\ExcelTemplate\" + "SummaryValuation.xlsx";
        Download_File(SampleFilePath, "SummaryValuation.xlsx");
    }
    protected void linkPostion_Click(object sender, EventArgs e)
    {
        string SampleFilePath = Server.MapPath("") + @"\ExcelTemplate\" + "Commitment.xlsx";
        Download_File(SampleFilePath, "Commitment.xlsx");
    }

    private void Download_File(string FilePath, string FileName)
    {
        Response.ContentType = ContentType;
        Response.AppendHeader("Content-Disposition", "attachment; filename=" + FileName);
        Response.WriteFile(FilePath);
        Response.End();
    }
    protected void linkPerformance_Click(object sender, EventArgs e)
    {
        string SampleFilePath = Server.MapPath("") + @"\ExcelTemplate\" + "Performance.xlsx";
        Download_File(SampleFilePath, "Performance.xlsx");
    }


    public void clearlabel()
    {
        lbltrans.Text = "";
        lblcommit.Text = "";
        lblsummary.Text = "";
        lblperf.Text = "";
    }
}