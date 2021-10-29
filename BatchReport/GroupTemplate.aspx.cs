using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Net;
using System.Collections;
using System.Collections.Generic;
using Spire.Xls;
using System.Drawing;
using System.Data.Common;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
//using CrmSdk;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Threading;
using Microsoft.IdentityModel.Claims;
using OfficeOpenXml;


public partial class _GroupTemplate : System.Web.UI.Page
{
    DB clsDB = null;
    GeneralMethods clsGS = new GeneralMethods();
    Boolean fbCheckExcel = false;
    public StreamWriter sw = null;
    string strDescription = string.Empty;
    bool bProceed = true;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, LegalEntityId, lsFamiliesName, lsDateName, ContactId, Template, FilePath, AUMAsOfDate, HHId, BillingName;

    GeneralMethods GM = new GeneralMethods();
    string RandomNumStr = string.Empty;

    int temp = 0;
    int intResult = 0;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            Bindddl();
        }

        Button1.Attributes.Add("onclick", "javascript:" + Button1.ClientID + ".disabled=true;" + ClientScript.GetPostBackEventReference(Button1, ""));
    }

    private void Bindddl()
    {
        GroupTemplate1();
        GroupTemplate2();
        GroupTemplate3();
        GroupTemplate4();
        GroupTemplate5();
        GroupTemplate6();
        GroupTemplate7();
        GroupTemplate8();
        GroupTemplate9();
        GroupTemplate10();
        GroupTemplate11();
        GroupTemplate12();
        GroupTemplate13();
        GroupTemplate14();
        GroupTemplate15();

    }

    #region BindGridview
    private void GroupTemplate1()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate1, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate2()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate2, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate3()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate3, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate4()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate4, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate5()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate5, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate6()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate6, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate7()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate7, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate8()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate8, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate9()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate9, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate10()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate10, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate11()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate11, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate12()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate12, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate13()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate13, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate14()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate14, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    private void GroupTemplate15()
    {
        clsDB = new DB();
        string sql = "SP_S_GroupTemplate_LIST";
        clsGS.getBindDDL(ddlGroupTemplate15, sql, "GroupTemplate", "ssi_mailidtemp");
    }

    #endregion

    protected void Button1_Click(object sender, EventArgs e)
    {
        lbtnExceptionReport.Visible = false;
        string strGroupTemplate = string.Empty;
        string[] strDuplicate = new string[5];


        string LogFileName = "LogFile " + DateTime.Now;
        string FolderName = "Logs";
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        //  Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        string strPath = Server.MapPath("") + "\\" + FolderName;
        if (!Directory.Exists(strPath))
        {
            Directory.CreateDirectory(strPath);
        }
        sw = new StreamWriter(strPath + "\\" + LogFileName + ".txt", true);

        #region Validate Save not used

        //for (int i = 1; i < 6; i++)
        //{           
        //    Control ddlGroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + i.ToString()));

        //    if (ddlGroupTemplate != null)
        //    {
        //        DropDownList GroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + i.ToString()));

        //        if (GroupTemplate.SelectedValue != "" && GroupTemplate.SelectedValue != "0")
        //        {
        //            strGroupTemplate = strGroupTemplate + "^" + GroupTemplate.SelectedValue;
        //        }
        //    }
        //}


        //strGroupTemplate = strGroupTemplate.Substring(1, strGroupTemplate.Length - 1);
        //strDuplicate = strGroupTemplate.Split('^');

        //for (int j = 1; j < strDuplicate.Length +1; j++)
        //{
        //    System.Text.StringBuilder sb = new System.Text.StringBuilder();
        //    Type tp = this.GetType();

        //    Control ddlGroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

        //    if (ddlGroupTemplate != null)
        //    {
        //        DropDownList GroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

        //        if (GroupTemplate.SelectedValue != "" && GroupTemplate.SelectedValue != "0")
        //        {
        //            if (GroupTemplate.SelectedValue == strDuplicate[j])
        //            {
        //                sb.Append("\n<script type=text/javascript>\n");
        //                sb.Append("\n alert('Fund already selected.Please select some other fund.');");
        //                sb.Append("</script>");
        //                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
        //                return;
        //            }
        //        }
        //    }
        //}


        #endregion


        lblError.Text = "";
        ViewState["UnifyResult"] = 0;


        //Unification Logic for Multiple MAIL IDs
        if (chkUnify.Checked)
        {
            string MailIdList = string.Empty;

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();

            for (int j = 1; j < 16; j++)
            {
                Control GroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

                if (GroupTemplate != null)
                {
                    DropDownList ddlGroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

                    if (ddlGroupTemplate.SelectedValue != "" && ddlGroupTemplate.SelectedValue != "0")
                    {
                        MailIdList = MailIdList + "," + ddlGroupTemplate.SelectedValue;
                    }
                }
            }


            if (MailIdList.Length == 0)
            {
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("\n alert('Please select template.');");
                sb.Append("</script>");
                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                return;
            }

            string strMailId = MailIdList.Substring(1, MailIdList.Length - 1);
            // MailIdList = MailIdList == "" || MailIdList == "0" ? "null" : "'" + MailIdList.Substring(1, MailIdList.Length - 1) + "'";

            ViewState["UnifyResult"] = UpdateMailRecords(strMailId);

        }


        if (RadioButton1.Checked)
        {
            fbCheckExcel = false;
            //mvShowReport.ActiveViewIndex = 1;
        }
        if (RadioButton2.Checked)
        {
            try
            {
                fbCheckExcel = true;
            }
            catch (Exception ex)
            {
                Response.Write(ex.ToString());
                Response.Write(ex.StackTrace);
            }
        }
        if (rdbtnPDF.Checked)
        {
            ViewState["DistributionWireInstFile"] = null;
            generatePDF();

            Bindddl();
            /*For Distribution WIre Instruction File throw*/
            try
            {
                if (ViewState["DistributionWireInstFile"] != null)
                {
                    string ls = Convert.ToString(ViewState["DistributionWireInstFile"]);

                    Random rand = new Random();
                    string strRndNumber = Convert.ToString(rand.Next(5));
                    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + strRndNumber;

                    string dbiFIleNamethr = "Distribution Wire Instruction_" + strGUID + "";

                    String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + dbiFIleNamethr + ".pdf";

                    string FileName = dbiFIleNamethr;
                    FileInfo loFile = new FileInfo(ls);
                    loFile.CopyTo(fsFinalLocation.Replace(".xls", ".pdf"), true);

                    Response.Write("<script>");
                    string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + FileName + ".pdf";
                    Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    Response.Write("</script>");
                }
                if (ViewState["BillingInvoiceInstFile"] != null)
                {
                    string ls = Convert.ToString(ViewState["BillingInvoiceInstFile"]);
                    string DateRange = Convert.ToString(ViewState["DateRange"]);

                    Random rand = new Random();
                    string strRndNumber = Convert.ToString(rand.Next(5));
                    string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + strRndNumber;

                    string dbiFIleNamethr = "Billing_" + DateRange + "";

                    String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + dbiFIleNamethr + ".pdf";

                    string FileName = dbiFIleNamethr;
                    FileInfo loFile = new FileInfo(ls);
                    loFile.CopyTo(fsFinalLocation.Replace(".xls", ".pdf"), true);

                    Response.Write("<script>");
                    string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + FileName + ".pdf";
                    Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    Response.Write("</script>");
                }
            }
            catch (Exception exc)
            {
                Response.Write(exc.Message);
                lblError.Text = exc.Message.ToString();
            }
            /* == END == */
        }


        if (sw != null)
        {
            sw.Flush();
            sw.Close();
        }
    }


    private int UpdateMailRecords(string MailId)
    {
        int intResult = 0;
        //string test = ddlMailType.SelectedValue;
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;

        //string test = ddlMailType.SelectedValue;

        //lblError.Text = "";
        DataSet loInvoiceData = null;
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }

        try
        {

            #region Update Mailing records to Unify
            clsDB = new DB();
            string strSql = "SP_S_MAIL_UNIFY @MailingId = '" + Convert.ToString(MailId) + "'";
            DataSet UpdateTempRecords = clsDB.getDataSet(strSql);

            //  ssi_mailrecordstemp objMailRecordsTemp = null;
            Entity objMailRecordsTemp = null;


            if (UpdateTempRecords.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < UpdateTempRecords.Tables[0].Rows.Count; i++)
                {
                    // objMailRecordsTemp = new ssi_mailrecordstemp();
                    objMailRecordsTemp = new Entity("ssi_mailrecordstemp");

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_mailrecordstempId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                        //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_mailrecordstempId"]));
                        objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_mailrecordstempId"]));
                    }


                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_contactfullname"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_contactfullname = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_contactfullname"]);
                        objMailRecordsTemp["ssi_contactfullname"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_contactfullname"]);
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
                    {
                        //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                        //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_contactfullnameid"]));

                        objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_contactfullnameid"])));
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_dear_mail"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_dear_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_dear_mail"]);
                        objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_dear_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_dear_mail = "";
                        objMailRecordsTemp["ssi_dear_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_salutation_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_salutation_mail"]);
                        objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_salutation_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_salutation_mail = "";
                        objMailRecordsTemp["ssi_salutation_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline1_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline1_mail"]);
                        objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline1_mail"]); ;
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_addressline1_mail = "";
                        objMailRecordsTemp["ssi_addressline1_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline2_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline2_mail"]);
                        objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline2_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_addressline2_mail = "";
                        objMailRecordsTemp["ssi_addressline2_mail"] = "";

                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline3_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline3_mail"]);
                        objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline3_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_addressline3_mail = "";
                        objMailRecordsTemp["ssi_addressline3_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_city_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_city_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_city_mail"]);
                        objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_city_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_city_mail = "";
                        objMailRecordsTemp["ssi_city_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_stateprovince_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_stateprovince_mail"]);
                        objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_stateprovince_mail"]); ;
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_stateprovince_mail = "";
                        objMailRecordsTemp["ssi_stateprovince_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_zipcode_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_zipcode_mail"]);
                        objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_zipcode_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_zipcode_mail = "";
                        objMailRecordsTemp["ssi_zipcode_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_countryregion_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_countryregion_mail"]);
                        objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_countryregion_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_countryregion_mail = "";
                        objMailRecordsTemp["ssi_countryregion_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_fullname_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_fullname_mail"]);
                        objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_fullname_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_fullname_mail = "";
                        objMailRecordsTemp["ssi_fullname_mail"] = "";
                    }

                    //File Name Added on 9 oct 2014 //commented becuase not required in unify logic
                    //if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_Filename"]) != "")
                    //{
                    //    objMailRecordsTemp.ssi_filename = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_Filename"]);
                    //}

                    //objMailRecordsTemp.ssi_updateunifyflg = new CrmBoolean();
                    //objMailRecordsTemp.ssi_updateunifyflg.Value = true;

                    objMailRecordsTemp["ssi_updateunifyflg"] = true;

                    service.Update(objMailRecordsTemp);
                    intResult++;
                }
            }



            #endregion

            #region Update Unify Flg

            if (UpdateTempRecords.Tables[1].Rows.Count > 0)
            {
                for (int i = 0; i < UpdateTempRecords.Tables[1].Rows.Count; i++)
                {
                    //   objMailRecordsTemp = new ssi_mailrecordstemp();
                    objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
                    // if (ddlMailId.SelectedValue != "")
                    //{
                    if (Convert.ToString(UpdateTempRecords.Tables[1].Rows[i]["Ssi_mailrecordstempId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                        //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[1].Rows[i]["Ssi_mailrecordstempId"]));
                        objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(UpdateTempRecords.Tables[1].Rows[i]["Ssi_mailrecordstempId"]));

                    }

                    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
                    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
                    objMailRecordsTemp["ssi_unifiedflg"] = true;

                    service.Update(objMailRecordsTemp);
                    intResult++;
                    // }
                }
            }

            #endregion

            #region Update Salutation to Unify

            if (UpdateTempRecords.Tables[2].Rows.Count > 0)
            {
                for (int i = 0; i < UpdateTempRecords.Tables[2].Rows.Count; i++)
                {
                    //objMailRecordsTemp = new ssi_mailrecordstemp();
                    objMailRecordsTemp = new Entity("ssi_mailrecordstemp");

                    if (Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Ssi_mailrecordstempId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                        //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Ssi_mailrecordstempId"]));
                        objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Ssi_mailrecordstempId"]));
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Sas_JointSalutation"]) != "")
                    {
                        //objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Sas_JointSalutation"]);
                        objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Sas_JointSalutation"]);
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["sas_dear2"]) != "")
                    {
                        //objMailRecordsTemp.ssi_dear_mail = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["sas_dear2"]);
                        objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["sas_dear2"]);
                    }

                    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
                    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
                    objMailRecordsTemp["ssi_unifiedflg"] = true;

                    service.Update(objMailRecordsTemp);

                    intResult++;
                }
            }



            #endregion


            if (intResult > 0)
            {
                //  if (ddlMailType.SelectedItem.Text != "")
                //  {
                //  lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully and records Unified successfully";
                //lblError.Text = ddlMailType.SelectedItem.Text + " records Unified successfully";
                //  }
                // else
                // {
                lblError.Text = "Records Unified successfully";
                // }
            }
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }

        if (intResult > 0)
            return 1;
        else
            return 0;

    }

    protected void btnBack_Click(object sender, EventArgs e)
    {
        //mvShowReport.ActiveViewIndex = 0;
        //txtAsofdate.Text = "";
        //txtpriorperiod.Text = "";
    }



    public void generatePDF()
    {

        //Page.Title = "Please Wait";

        //trbtnSubmit.Visible = false;

        #region Validate Group

        bool checkDuplicate = false;

        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        Type tp = this.GetType();



        #region Check for ddlGroupTemplate1

        if (ddlGroupTemplate1.SelectedValue != "" && ddlGroupTemplate1.SelectedValue != "0")
        {
            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate1.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }
        }

        #endregion

        #region Check for ddlGroupTemplate2

        if (ddlGroupTemplate2.SelectedValue != "" && ddlGroupTemplate2.SelectedValue != "0")
        {
            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate2.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate3

        if (ddlGroupTemplate3.SelectedValue != "" && ddlGroupTemplate3.SelectedValue != "0")
        {
            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate3.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate4

        if (ddlGroupTemplate4.SelectedValue != "" && ddlGroupTemplate4.SelectedValue != "0")
        {
            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate4.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate5

        if (ddlGroupTemplate5.SelectedValue != "" && ddlGroupTemplate5.SelectedValue != "0")
        {
            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate5.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate6

        if (ddlGroupTemplate6.SelectedValue != "" && ddlGroupTemplate6.SelectedValue != "0")
        {
            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate6.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate7

        if (ddlGroupTemplate7.SelectedValue != "" && ddlGroupTemplate7.SelectedValue != "0")
        {
            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate7.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate8

        if (ddlGroupTemplate8.SelectedValue != "" && ddlGroupTemplate8.SelectedValue != "0")
        {
            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate8.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate9

        if (ddlGroupTemplate9.SelectedValue != "" && ddlGroupTemplate9.SelectedValue != "0")
        {
            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate9.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate10

        if (ddlGroupTemplate10.SelectedValue != "" && ddlGroupTemplate10.SelectedValue != "0")
        {
            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate10.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate11

        if (ddlGroupTemplate11.SelectedValue != "" && ddlGroupTemplate11.SelectedValue != "0")
        {
            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate11.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate12

        if (ddlGroupTemplate12.SelectedValue != "" && ddlGroupTemplate12.SelectedValue != "0")
        {
            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate12.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }


        }

        #endregion

        #region Check for ddlGroupTemplate13

        if (ddlGroupTemplate13.SelectedValue != "" && ddlGroupTemplate13.SelectedValue != "0")
        {
            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate13.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate14

        if (ddlGroupTemplate14.SelectedValue != "" && ddlGroupTemplate14.SelectedValue != "0")
        {
            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate14.SelectedValue == ddlGroupTemplate15.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion

        #region Check for ddlGroupTemplate15

        if (ddlGroupTemplate15.SelectedValue != "" && ddlGroupTemplate15.SelectedValue != "0")
        {
            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate1.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate2.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate3.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate4.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate5.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate6.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate7.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate8.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate9.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate10.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate11.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate12.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate13.SelectedValue)
            {
                checkDuplicate = true;
            }

            if (ddlGroupTemplate15.SelectedValue == ddlGroupTemplate14.SelectedValue)
            {
                checkDuplicate = true;
            }

        }

        #endregion



        if (checkDuplicate == true)
        {
            sb.Append("\n<script type=text/javascript>\n");
            sb.Append("\n alert('This template is already selected.Please select some other template.');");
            sb.Append("</script>");
            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
            return;
        }



        #endregion

        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;
        //ssi_batch objBatch = null;
        Entity objBatch = null;

        Entity objBillingInvoice = null;

        //ssi_mailrecords objMailRecords = null;
        Entity objMailRecords = null;

        //  ssi_mailrecordstemp objMailRecordsTemp = null;
        Entity objMailRecordsTemp = null;

        string HouseHold = "";

        int selectedCount = 0;

        string ReportName = string.Empty;
        string CustomFlg = string.Empty;
        string CustomTemplateType = string.Empty;
        string TemplateId = string.Empty;
        string PdfFinalPath = string.Empty;
        string strMailIdTemp = string.Empty;
        string MailIdList = string.Empty;

        string UserId = string.Empty;

        int UniqueMailingId = 0;
        lblError.Visible = true;
        lblError.Text = "";

        DataSet loInvoiceData = null;
        bool bOpsApproveRequestFlg = false;

        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        string[] strTemplate = null;
        string ConsolidatePdfFileName = string.Empty;
        string DestinationPath = string.Empty;
        bool bDistributionWire = false;
        bool bBillingInvoice = false;
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;



        for (int j = 1; j < 16; j++)
        {
            Control GroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

            if (GroupTemplate != null)
            {
                DropDownList ddlGroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

                if (ddlGroupTemplate.SelectedValue != "" && ddlGroupTemplate.SelectedValue != "0")
                {
                    MailIdList = MailIdList + "," + ddlGroupTemplate.SelectedValue;
                }
            }
        }


        if (MailIdList.Length == 0)
        {
            sb.Append("\n<script type=text/javascript>\n");
            sb.Append("\n alert('Please select template.');");
            sb.Append("</script>");
            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
            return;
        }



        DateTime dt = DateTime.Now;

        string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
        string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
        string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();

        string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
        string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
        string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

        //ViewState["ParentFolder"] = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;
        string ReportOpFolder = string.Empty;
        ////string ReportOpFolder = "\\\\Fs01\\_ops_C_I_R_group\\Quarterly_Reports\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

        ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports); // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

        if (Request.Url.AbsoluteUri.Contains("localhost"))
        {
            ReportOpFolder = Request.MapPath("..\\ExcelTemplate\\BATCH REPORTS\\");  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
        }
        else
            ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();


        string ContactFolderName = string.Empty;

        string strMailId = MailIdList.Substring(1, MailIdList.Length - 1);
        MailIdList = MailIdList == "" || MailIdList == "0" ? "null" : "'" + MailIdList.Substring(1, MailIdList.Length - 1) + "'";

        //added 4_21_2020 Exception Report
        int strMailingList = 0;

        string sql = "SP_S_LegalEntityContact @MailIDList=" + MailIdList;
        DataSet DSLegalEntity = clsDB.getDataSet(sql);
        int Count = 0;
        for (int i = 0; i < DSLegalEntity.Tables[0].Rows.Count; i++)
        {
            //added 4_21_2020 Exception Report
            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]) != "")
            {
                strMailingList = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]);
            }

            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"]) != "")
            {
                LegalEntityId = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"]).Replace("'", "''");
            }


            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
            {
                ContactId = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]).Replace("'", "''");
            }
            //    ContactId = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]) == "" ? "null" : "'" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]) + "'";


            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_aumasofdate"]) != "")
                AUMAsOfDate = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_aumasofdate"]);
            else
                AUMAsOfDate = "";

            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHoldID"]) != "")
                HHId = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHoldID"]).Replace("'", "''");
            else
                HHId = "";

            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]) != "")
            {
                BillingName = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]).Replace("'", "''");
            }


            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) != "")
            {
                Random rnd = new Random();
                int rndNum = rnd.Next(1000, 9999);
                string RandomNumStr = "#" + rndNum.ToString();
                ConsolidatePdfFileName = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr + ".pdf";
                ConsolidatePdfFileName = GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);

                DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);
                string strFinalPath = ConsolidatePdfFileName;// Server.MapPath("") + @"\ExcelTemplate\DTS\" + DateTime.Now.ToString("ddMMMyyyymmss") + LegalEntityId + ".pdf"; //FileName
                string[] MailId = strMailId.Split(',');
                Template = string.Empty;

                for (int k = 0; k < MailId.Length; k++)
                {
                    strMailIdTemp = MailId[k];
                    string strsql = "";

                    if (ContactId != "" && ContactId != null)
                    {
                        if (!string.IsNullOrEmpty(LegalEntityId))
                            strsql = "exec SP_S_LegalEntityContact_Template @MailIDList='" + strMailIdTemp + "',@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID='" + ContactId + "'";
                        else
                        {
                            strsql = "exec SP_S_LegalEntityContact_Template @MailIDList='" + strMailIdTemp + "',@LegalEntityNameID=null,@ContactFullnameID='" + ContactId + "'";
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(LegalEntityId))
                            strsql = "exec SP_S_LegalEntityContact_Template @MailIDList='" + strMailIdTemp + "',@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID= null";
                        else
                            strsql = "exec SP_S_LegalEntityContact_Template @MailIDList='" + strMailIdTemp + "',@LegalEntityNameID=null,@ContactFullnameID= null";
                    }

                    DataSet DSNew = clsDB.getDataSet(strsql);

                    if (DSNew.Tables[0].Rows.Count > 0)
                    {
                        if (DSNew.Tables[0].Rows[0]["templatename"].ToString() != "")
                        {
                            ReportName = DSNew.Tables[0].Rows[0]["templatename"].ToString();
                            CustomFlg = DSNew.Tables[0].Rows[0]["ssi_customflg"].ToString();
                            CustomTemplateType = DSNew.Tables[0].Rows[0]["ssi_custometemplatetype"].ToString();
                            TemplateId = DSNew.Tables[0].Rows[0]["ssi_templateID"].ToString();

                            if (MailId.Length == 1)
                            {
                                if (ReportName.ToUpper() == "Distribution Wire Instructions".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
                                {
                                    bDistributionWire = true;
                                    ViewState["MailId"] = strMailIdTemp;
                                    ViewState["TemplateId"] = TemplateId;
                                }
                                if (ReportName.ToUpper() == "Billing Invoice".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
                                {
                                    bBillingInvoice = true;
                                    ViewState["MailId"] = strMailIdTemp;
                                    ViewState["Ssi_MailingID"] = strMailingList;
                                    ViewState["TemplateId"] = TemplateId;
                                }

                            }

                            PdfFinalPath = generateCombinedPDF(LegalEntityId, ContactId, strMailIdTemp, strFinalPath, ReportName, CustomFlg, CustomTemplateType, TemplateId, DSLegalEntity.Tables[0].Rows.Count, AUMAsOfDate, HHId, BillingName);

                            if (Template != "")
                            {
                                Template = Template + "|" + PdfFinalPath;
                            }
                            else
                            {
                                Template = "|" + PdfFinalPath;
                            }


                        }
                    }
                }


                #region Merge PDF

                // Template = Template.Substring(1, Template.Length - 1);
                strTemplate = Template.Split('|');

                for (int l = 0; l < strTemplate.Length; l++)
                {
                    if (FilePath != "")
                    {
                        if (strTemplate[l] != "")
                        {
                            FilePath = FilePath + "|" + strTemplate[l];
                        }
                    }
                    else
                    {
                        if (strTemplate[l] != "")
                        {
                            FilePath = "|" + strTemplate[l];
                        }
                    }
                }

                #endregion

                #region Delete Files

                if (FilePath != "" && FilePath != null)
                {

                    FilePath = FilePath.Substring(1, FilePath.Length - 1);
                    string[] strPath = FilePath.Split('|');
                    PDFMerge PDF = new PDFMerge();
                    PDF.MergeFiles(DestinationPath, strPath);

                    string filesToDelete = "*.pdf";
                    string Path = "C:\\AdventReport\\BatchReport\\ExcelTemplate\\pdfOutput";
                    string Path1 = "C:\\AdventReport\\ExcelTemplate\\pdfOutput";
                    //string Path = Request.MapPath("\\Advent Report\\BatchReport\\ExcelTemplate\\pdfOutput\\");
                    //string Path1 = Request.MapPath("\\Advent Report\\ExcelTemplate\\pdfOutput\\");
                    if (Path != "")
                    {
                        if (Directory.Exists(Path))
                        {
                            string[] fileList = System.IO.Directory.GetFiles(Path, filesToDelete);

                            foreach (string file in fileList)
                            {
                                try
                                {
                                    System.IO.File.Delete(file);
                                }
                                catch
                                { }

                                //sResult += "\n" + file + "\n";
                            }
                        }
                    }

                    if (Path1 != "")
                    {
                        if (Directory.Exists(Path1))
                        {
                            string[] fileList = System.IO.Directory.GetFiles(Path1, filesToDelete);

                            foreach (string file in fileList)
                            {
                                try
                                {
                                    System.IO.File.Delete(file);
                                }
                                catch
                                { }
                            }
                        }
                    }
                }

                #endregion

                //Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
                if (Request.QueryString["test"] != null)
                    temp = 1;
                else
                    temp = 0;

                if (temp != 1) //Temp
                {
                    if (FilePath != "" && FilePath != null)
                    {
                        #region Create Batch

                        //objBatch = new ssi_batch();
                        objBatch = new Entity("ssi_batch");


                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_Filename"]) != "")
                        {
                            //objBatch.ssi_batchdisplayfilename = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr + ".pdf";
                            objBatch["ssi_batchdisplayfilename"] = GeneralMethods.RemoveSpecialCharacters( Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_Filename"])) + RandomNumStr + ".pdf";
                        }

                        if (DestinationPath != "")
                        {
                            //objBatch.ssi_batchfilename = DestinationPath;
                            objBatch["ssi_batchfilename"] = DestinationPath;
                        }


                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) != "")
                        {
                            //objBatch.ssi_householdid = new Lookup();
                            //objBatch.ssi_householdid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"]));
                            objBatch["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"])));
                        }

                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"]) != "")
                        {
                            //objBatch.ssi_contactid = new Lookup();
                            //objBatch.ssi_contactid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]));
                            objBatch["ssi_contactid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"])));
                        }

                        //objBatch.ssi_reporttrackerstatus = new Picklist();
                        //objBatch.ssi_reporttrackerstatus.Value = 12;//Initial Review
                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(12);
                        #region OLD ssi_sharepointreportfolder and Clientportal Fields
                        //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHold"]) != "" && Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]) != "")
                        //{
                        //    HouseHold = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHold"]).Replace(" Family", "");
                        //  ////  objBatch.ssi_sharepointreportfolder = "http://sp02/ClientServ/Documents/Clients/Active/" + HouseHold + "/Correspondence/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                        //  //  objBatch.ssi_sharepointreportfolder = "https://greshampartners.sharepoint.com/clientserv/Documents/Clients/Active/" + HouseHold + "/Correspondence/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                        //    objBatch["ssi_sharepointreportfolder"] = "https://greshampartners.sharepoint.com/clientserv/Documents/Clients/Active/" + HouseHold + "/Correspondence/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                        //}
                        #endregion
                        #region NEW CS Sharepoin Changes - added 2_21_2019
                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) != "")
                        {
                            objBatch["ssi_cshouseholdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"])));
                        }
                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["YearFromFile"]) != "")
                        {
                            objBatch["ssi_year"] = Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["YearFromFile"]));
                        }
                        objBatch["ssi_spsitetype"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000000);
                        #endregion
                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["HouseHold"]) != "" && Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["YearFromFile"]) != "")
                        {
                            // HouseHold = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHold"]).Replace(" Family", "");
                            // objBatch.ssi_clientportalfolder = "http://sp02/Client Portal/Documents/" + HouseHold + "/Investment Activity/NonMarketable/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                            if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_clientreportfolder"]) != "")
                            {
                                //objBatch.ssi_clientportalfolder = "https://greshampartners.sharepoint.com/ClientPortal/Documents taxonomy";
                                // objBatch["ssi_clientportalfolder"] = "https://greshampartners.sharepoint.com/ClientPortal/Documents taxonomy";
                                objBatch["ssi_clientportalfolder"] =  AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL) +"/Documents taxonomy";
                            }
                        }


                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]) != "")
                        {
                            //objBatch.ssi_name = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]);
                            objBatch["ssi_name"] = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]);
                        }

                        //Advisor Approval Required

                        //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "0" || Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "" || Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "False".ToUpper())
                        //{
                        //    objBatch.ssi_advisorapprovalreqd.Value = false;
                        //}
                        //else 
                        if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_advisorapprovalreqd"]) == "1" || Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "True".ToUpper())
                        {
                            //objBatch.ssi_advisorapprovalreqd = new CrmBoolean();
                            //objBatch.ssi_advisorapprovalreqd.Value = true;
                            objBatch["ssi_advisorapprovalreqd"] = true;


                        }

                        //objBatch.ssi_reporttracker = new CrmBoolean();
                        //objBatch.ssi_reporttracker.Value = true;
                        objBatch["ssi_reporttracker"] = true;


                        //objBatch.ssi_type = new Picklist();
                        //objBatch.ssi_type.Value = 4;//Merge
                        objBatch["ssi_type"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                        //objBatch.ssi_approvalreqd = new CrmBoolean();
                        //objBatch.ssi_approvalreqd.Value = true;
                        objBatch["ssi_approvalreqd"] = true;


                        UserId = GetcurrentUser();

                        if (UserId != "" && UserId != null)
                        {
                            //objBatch.ssi_createdbycustomid = new Lookup();
                            //objBatch.ssi_createdbycustomid.Value = new Guid(UserId);
                            objBatch["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                        }

                        if (DSLegalEntity.Tables[1].Rows[i]["ssi_clientportalname"].ToString() != "")
                        {

                            //objBatch.ssi_clientportalname = DSLegalEntity.Tables[0].Rows[i]["ssi_clientportalname"].ToString();
                            objBatch["ssi_clientportalname"] = DSLegalEntity.Tables[1].Rows[i]["ssi_clientportalname"].ToString();

                        }


                        service.Create(objBatch);
                        intResult++;
                        ViewState["intResult"] = intResult;
                        if (intResult > 0)
                        {
                            clsDB = new DB();

                            string HouseHoldId = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) == "" ? "null" : "'" + Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) + "'";
                            string strContactId = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"]) == "" ? "null" : "'" + Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"]) + "'";
                            string BatchName = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]) == "" ? "null" : "'" + Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]).Replace("'", "''") + "'";

                            string strsql = "SP_S_BatchID @ssi_householdid=" + HouseHoldId + ",@Ssi_ContactId=" + strContactId + ",@Ssi_name=" + BatchName;
                            DataSet DSBatch = clsDB.getDataSet(strsql);

                            if (DSBatch.Tables[0].Rows.Count > 0)
                            {
                                if (Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]) != "")
                                {
                                    #region Update Batch Owner
                                    //objBatch.ssi_batchid = new Key();
                                    //objBatch.ssi_batchid.Value = new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]));
                                    objBatch["ssi_batchid"] = new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]));

                                    //SecurityPrincipal assignee = new SecurityPrincipal();
                                    //assignee.PrincipalId = new Guid(AppLogic.GetParam(AppLogic.ConfigParam.OpsReporting));////OPS Reporting Gresham

                                    //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();

                                    //targetAssign.EntityId = new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]));
                                    //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                                    //AssignRequest assign = new AssignRequest();
                                    //assign.Assignee = assignee;
                                    //assign.Target = targetAssign;
                                    //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);


                                    //service.Update(objBatch);

                                    //AssignRequest assignRequest = new AssignRequest()
                                    //{
                                    //    Assignee = new EntityReference
                                    //    {
                                    //        //LogicalName = "team",
                                    //        Id = new Guid(AppLogic.GetParam(AppLogic.ConfigParam.OpsReporting))
                                    //    },

                                    //    Target = new EntityReference("ssi_batch",new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"])))

                                    //};

                                    AssignRequest assignRequest = new AssignRequest
                                    {
                                        Assignee = new EntityReference("systemuser", new Guid(AppLogic.GetParam(AppLogic.ConfigParam.OpsReporting))),
                                        Target = new EntityReference("ssi_batch", new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"])))
                                    };

                                    service.Execute(assignRequest);
                                    service.Update(objBatch);

                                    #endregion

                                    #region Create New Mail Records

                                    //objMailRecords = new ssi_mailrecords();
                                    objMailRecords = new Entity("ssi_mailrecords");

                                    // BatchId Lookup 
                                    if (Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]) != "")
                                    {
                                        //objMailRecords.ssi_batchid = new Lookup();
                                        //objMailRecords.ssi_batchid.Value = new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]));
                                        objMailRecords["ssi_batchid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_batch", new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"])));

                                    }

                                    // Batch Id text
                                    if (Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]) != "")
                                    {
                                        //objMailRecords.ssi_batchidtxt = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);
                                        objMailRecords["ssi_batchidtxt"] = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);
                                    }

                                    //Batch Name 
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]) != "")
                                    {
                                        //objMailRecords.ssi_batchnametxt = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]);
                                        objMailRecords["ssi_batchnametxt"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]);
                                    }

                                    //Name
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Name"]) != "")
                                    {
                                        //objMailRecords.ssi_name = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Name"]);
                                        objMailRecords["ssi_name"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Name"]);
                                    }

                                    //Mail Status
                                    //objMailRecords.ssi_mailstatus = new Picklist();
                                    //objMailRecords.ssi_mailstatus.Value = 1;//Pending
                                    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(1);

                                    //Mail Type
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Type"]) != "")
                                    {
                                        //objMailRecords.ssi_mailtypeid = new Lookup();
                                        //objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Type"]));
                                        objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Type"])));
                                    }


                                    //Contact
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
                                    {
                                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]));
                                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"])));
                                    }

                                    //HouseHold
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"]) != "")
                                    {
                                        //objMailRecords.ssi_accountid = new Lookup();
                                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"]));
                                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"])));
                                    }


                                    //ssi_legalentitynameid *
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"]) != "")
                                    {
                                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"]));
                                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"])));

                                    }


                                    //Salutation
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Salutation"]) != "")
                                    {
                                        //objMailRecords.ssi_salutation_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Salutation"]);
                                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Salutation"]);
                                    }

                                    // FullName
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Full Name"]) != "")
                                    {
                                        //objMailRecords.ssi_fullname_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Full Name"]);
                                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Full Name"]);
                                    }

                                    //Address Line 1
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 1"]) != "")
                                    {
                                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 1"]);
                                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 1"]);
                                    }

                                    //Address Line 2
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 2"]) != "")
                                    {
                                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 2"]);
                                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 2"]);
                                    }

                                    //Address Line 3
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 3"]) != "")
                                    {
                                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 3"]);
                                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 3"]);
                                    }

                                    //City
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["City"]) != "")
                                    {
                                        //objMailRecords.ssi_city_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["City"]);
                                        objMailRecords["ssi_city_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["City"]);
                                    }

                                    //State/Province
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["State/Province"]) != "")
                                    {
                                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["State/Province"]);
                                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["State/Province"]);
                                    }

                                    //Zip Code
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Zip Code"]) != "")
                                    {
                                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Zip Code"]);
                                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Zip Code"]);
                                    }

                                    //Country/Region
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Country/Region"]) != "")
                                    {
                                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Country/Region"]);
                                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Country/Region"]);
                                    }


                                    //Ssi_MailingID
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]) != "")
                                    {
                                        //objMailRecords.ssi_mailingid = new CrmNumber();
                                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]);
                                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]);
                                    }


                                    //AsOfDate
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["AsOfDate"]) != "")
                                    {
                                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                                        //objMailRecords.ssi_asofdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["AsOfDate"]);
                                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["AsOfDate"]));
                                    }

                                    //Dear
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Dear"]) != "")
                                    {
                                        //objMailRecords.ssi_dear_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Dear"]);
                                        objMailRecords["ssi_dear_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Dear"]);
                                    }

                                    //Spouse/Partner
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Spouse/Partner"]) != "")
                                    {
                                        //objMailRecords.ssi_spousepart_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Spouse/Partner"]);
                                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Spouse/Partner"]);
                                    }

                                    //Mail
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail"]) != "")
                                    {
                                        //objMailRecords.ssi_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail"]);
                                        objMailRecords["ssi_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail"]);
                                    }

                                    //Mail Preference
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Preference"]) != "")
                                    {
                                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Preference"]);
                                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Preference"]);
                                    }

                                    //Status Reason
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Status Reason"]) != "")
                                    {
                                        //objMailRecords.ssi_status = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Status Reason"]);
                                        objMailRecords["ssi_status"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Status Reason"]);
                                    }

                                    //Contact Owner First Name
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner First Name"]) != "")
                                    {
                                        //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner First Name"]);
                                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner First Name"]);
                                    }


                                    //Contact Owner Last Name
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner Last Name"]) != "")
                                    {
                                        //objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner Last Name"]);
                                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner Last Name"]);
                                    }

                                    //Household Owner First Name
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner First Name"]) != "")
                                    {
                                        //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner First Name"]);
                                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner First Name"]);
                                    }

                                    //Household Owner Last Name
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner Last Name"]) != "")
                                    {
                                        //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner Last Name"]);
                                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner Last Name"]);
                                    }

                                    //Secondary Owner First Name
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
                                    {
                                        //objMailRecords.ssi_secownerfname_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner First Name"]);
                                        objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner First Name"]);
                                    }

                                    //Secondary Owner Last Name
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
                                    {
                                        //objMailRecords.ssi_secownerlname_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                                        objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_wireasofdate"]) != "")
                                    {
                                        //objMailRecords.ssi_wireasofdate = new CrmDateTime();
                                        //objMailRecords.ssi_wireasofdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_wireasofdate"]);
                                        objMailRecords["ssi_wireasofdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_wireasofdate"]));
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_letterdate"]) != "")
                                    {
                                        //objMailRecords.ssi_letterdate = new CrmDateTime();
                                        //objMailRecords.ssi_letterdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_letterdate"]);
                                        objMailRecords["ssi_letterdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_letterdate"]));
                                    }



                                    //CustomMailPreference
                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["CustomMailPreference"]) != "")
                                    {
                                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["CustomMailPreference"]);
                                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["CustomMailPreference"]);
                                    }


                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) != "")
                                    {
                                        //objMailRecords.ssi_filename = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr;
                                        objMailRecords["ssi_filename"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr;
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]) != "")
                                    {
                                        //objMailRecords.ssi_fundname = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]);
                                        objMailRecords["ssi_fundname"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]);
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]) != "")
                                    {
                                        //objMailRecords.ssi_anzianoid = new CrmNumber();
                                        //objMailRecords.ssi_anzianoid.Value = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["ssi_anzianoid"]);
                                        objMailRecords["ssi_anzianoid"] = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["ssi_anzianoid"]);
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_tnrid_nv"]) != "")
                                    {
                                        //objMailRecords.ssi_tnrid_nv = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_tnrid_nv"]);
                                        objMailRecords["ssi_tnrid_nv"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_tnrid_nv"]);
                                    }

                                    //*** Records for Billing Start ***//

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_AUM"]) != "")
                                    {
                                        //objMailRecords.ssi_aum = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_AUM"]);
                                        objMailRecords["ssi_aum"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_AUM"]);
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_AUMasofdate"]) != "")
                                    {
                                        //objMailRecords.ssi_aumasofdate = new CrmDateTime();
                                        //objMailRecords.ssi_aumasofdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_AUMasofdate"]);
                                        objMailRecords["ssi_aumasofdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_AUMasofdate"]));
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_annualfeebilling"]) != "")
                                    {
                                        //objMailRecords.ssi_annualfeebilling = new CrmMoney();
                                        //objMailRecords.ssi_annualfeebilling.Value = Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_annualfeebilling"]));
                                        objMailRecords["ssi_annualfeebilling"] = new Money(Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_annualfeebilling"])));
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeratebilling"]) != "")
                                    {
                                        //objMailRecords.ssi_feeratebilling = new CrmFloat();
                                        //objMailRecords.ssi_feeratebilling.Value = Convert.ToDouble(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeratebilling"]));
                                        objMailRecords["ssi_feeratebilling"] = Convert.ToDouble(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeratebilling"]));
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingidbilling"]) != "")
                                    {
                                        //objMailRecords.ssi_billingid = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingidbilling"]);
                                        objMailRecords["ssi_billingid"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingidbilling"]);
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_quarterlyfeebilling"]) != "")
                                    {
                                        //objMailRecords.ssi_quarterlyfeebilling = new CrmMoney();
                                        //objMailRecords.ssi_quarterlyfeebilling.Value = Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_quarterlyfeebilling"]));
                                        objMailRecords["ssi_quarterlyfeebilling"] = new Money(Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_quarterlyfeebilling"])));

                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_invoicedatebilling"]) != "")
                                    {
                                        //objMailRecords.ssi_invoicedatebilling = new CrmDateTime();
                                        //objMailRecords.ssi_invoicedatebilling.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_invoicedatebilling"]);
                                        objMailRecords["ssi_invoicedatebilling"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_invoicedatebilling"]));
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_autodebitdate"]) != "")
                                    {
                                        //objMailRecords.ssi_autodebitdate = new CrmDateTime();
                                        //objMailRecords.ssi_autodebitdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_autodebitdate"]);
                                        objMailRecords["ssi_autodebitdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_autodebitdate"]));
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_CustomBillingREF"]) != "")
                                    {
                                        //objMailRecords.ssi_custombillingref = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_CustomBillingREF"]);
                                        objMailRecords["ssi_custombillingref"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_CustomBillingREF"]);
                                    }

                                    if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]) != "")
                                    {
                                        //objMailRecords.ssi_billingname = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]);
                                        objMailRecords["ssi_billingname"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]);
                                    }

                                    #region not used
                                    //*** Records for Billing End ***//


                                    ////ssi_calledtodatep_ccsf
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_calledtodatep_ccsf"]) != "")
                                    //{
                                    //    objMailRecords.ssi_calledtodatep_ccsf = new CrmDecimal();
                                    //   objMailRecords.ssi_calledtodatep_ccsf.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_calledtodatep_ccsf"]);
                                    //}


                                    ////ssi_remainingcommitment_ccsf
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_remainingcommitment_ccsf"]) != "")
                                    //{
                                    //    objMailRecords.ssi_remainingcommitment_ccsf = new CrmMoney();
                                    //    objMailRecords.ssi_remainingcommitment_ccsf.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_remainingcommitment_ccsf"]);
                                    //}


                                    ////ssi_remainingcommitmentp_ccsf
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_remainingcommitmentp_ccsf"]) != "")
                                    //{
                                    //    objMailRecords.ssi_remainingcommitmentp_ccsf = new CrmDecimal();
                                    //    objMailRecords.ssi_remainingcommitmentp_ccsf.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_remainingcommitmentp_ccsf"]);
                                    //}


                                    ////ssi_totalcommitment_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_totalcommitment_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_totalcommitment_db = new CrmMoney();
                                    //    objMailRecords.ssi_totalcommitment_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_totalcommitment_db"]);
                                    //}


                                    ////ssi_capitaldistribution_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_capitaldistribution_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_capitaldistribution_db = new CrmMoney();
                                    //    objMailRecords.ssi_capitaldistribution_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_capitaldistribution_db"]);
                                    //}


                                    ////ssi_curdistp_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_curdistp_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_curdistp_db = new CrmDecimal();
                                    //    objMailRecords.ssi_curdistp_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_curdistp_db"]);
                                    //}


                                    ////ssi_curdistp_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_priordistributions_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_priordistributions_db = new CrmMoney();
                                    //    objMailRecords.ssi_priordistributions_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_priordistributions_db"]);
                                    //}

                                    ////ssi_priordistp_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_priordistp_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_priordistp_db = new CrmDecimal();
                                    //    objMailRecords.ssi_priordistp_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_priordistp_db"]);
                                    //}


                                    ////ssi_priordistp_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_disttodate_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_disttodate_db = new CrmMoney();
                                    //    objMailRecords.ssi_disttodate_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_disttodate_db"]);
                                    //}


                                    ////ssi_distdatep_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_distdatep_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_distdatep_db = new CrmDecimal();
                                    //    objMailRecords.ssi_distdatep_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_distdatep_db"]);
                                    //}

                                    ////ssi_calledtodate_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_calledtodate_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_calledtodate_db = new CrmMoney();
                                    //    objMailRecords.ssi_calledtodate_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_calledtodate_db"]);
                                    //}


                                    ////ssi_remainingcommitment_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_remainingcommitment_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_remainingcommitment_db = new CrmMoney();
                                    //    objMailRecords.ssi_remainingcommitment_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_remainingcommitment_db"]);
                                    //}


                                    ////ssi_percentcalled_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_percentcalled_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_percentcalled_db = new CrmDecimal();
                                    //    objMailRecords.ssi_percentcalled_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_percentcalled_db"]);
                                    //}

                                    ////ssi_feeadj_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeadj_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_feeadj_db = new CrmMoney();
                                    //    objMailRecords.ssi_feeadj_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_feeadj_db"]);
                                    //}

                                    ////ssi_actualcashdistributions_db
                                    //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_actualcashdistributions_db"]) != "")
                                    //{
                                    //    objMailRecords.ssi_actualcashdistributions_db = new CrmMoney();
                                    //    objMailRecords.ssi_actualcashdistributions_db.Value = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_actualcashdistributions_db"]);
                                    //}
                                    #endregion

                                    //objMailRecords.ssi_ir_status = new Picklist();
                                    //objMailRecords.ssi_ir_status.Value = 1;//Pending;
                                    objMailRecords["ssi_ir_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(1);

                                    UserId = GetcurrentUser();


                                    if (UserId != "")
                                    {
                                        //objMailRecords.ssi_createdbycustomid = new Lookup();
                                        //objMailRecords.ssi_createdbycustomid.Value = new Guid(UserId);
                                        objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                                    }

                                    service.Create(objMailRecords);

                                    #endregion

                                    #region Mail Records Temp Update

                                    string strMailSql = "SP_S_MailRecordsTempID_List @MailIDList=" + MailIdList + ",@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID='" + ContactId + "'";
                                    DataSet MailRecordsDataset = clsDB.getDataSet(strMailSql);

                                    for (int j = 0; j < MailRecordsDataset.Tables[0].Rows.Count; j++)
                                    {
                                        //objMailRecordsTemp = new ssi_mailrecordstemp();
                                        objMailRecordsTemp = new Entity("ssi_mailrecordstemp");

                                        if (Convert.ToString(MailRecordsDataset.Tables[0].Rows[j]["ssi_mailrecordstempid"]) != "")
                                        {
                                            //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                                            //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(MailRecordsDataset.Tables[0].Rows[j]["ssi_mailrecordstempid"]));
                                            objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(MailRecordsDataset.Tables[0].Rows[j]["ssi_mailrecordstempid"]));

                                            //objMailRecordsTemp.ssi_batchstatus = new Picklist();
                                            //objMailRecordsTemp.ssi_batchstatus.Value = 1;//Batched
                                            objMailRecordsTemp["ssi_batchstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(1);

                                            //objMailRecordsTemp.ssi_batchidtxt = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);
                                            objMailRecordsTemp["ssi_batchidtxt"] = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);

                                            service.Update(objMailRecordsTemp);
                                        }
                                    }

                                    #endregion

                                    intResult++;

                                }
                            }


                        }

                        if (intResult > 0)
                        {
                            if (Convert.ToString(ViewState["UnifyResult"]) != "0" && Convert.ToString(ViewState["UnifyResult"]) != "")
                                lblError.Text = "Group Templates and Batches Created Successfully and records Unified successfully";
                            else
                                lblError.Text = "Group Templates and Batches Created Successfully.";

                            lblError.Visible = true;
                            Button1.Visible = true;
                            // Label1.Visible = false;
                        }


                        //EXECUTE dbo.[SP_S_BatchID] @ssi_householdid = '8D0A713D-6A15-DE11-8391-001D09665E8F', @Ssi_ContactId = 'FC063302-DD15-DE11-8391-001D09665E8F'
                        //, @Ssi_name = 'James J. Glasser | Glasser, James J. | 04/30/2012'
                        //SP_S_BatchID

                        #endregion
                    }

                    #region Mail Records Temp Update

                    string strMailSql1 = "SP_S_MailRecordsTempID_List @MailIDList=" + MailIdList;// +",@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID='" + ContactId + "'";
                    DataSet MailRecordsDataset1 = clsDB.getDataSet(strMailSql1);

                    for (int j = 0; j < MailRecordsDataset1.Tables[0].Rows.Count; j++)
                    {
                        //objMailRecordsTemp = new ssi_mailrecordstemp();
                        objMailRecordsTemp = new Entity("ssi_mailrecordstemp");

                        try
                        {
                            if (Convert.ToString(MailRecordsDataset1.Tables[0].Rows[j]["ssi_mailrecordstempid"]) != "")
                            {
                                //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                                //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(MailRecordsDataset1.Tables[0].Rows[j]["ssi_mailrecordstempid"]));
                                objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(MailRecordsDataset1.Tables[0].Rows[j]["ssi_mailrecordstempid"]));

                                //objMailRecordsTemp.ssi_batchstatus = new Picklist();
                                //objMailRecordsTemp.ssi_batchstatus.Value = 1;//Batched
                                objMailRecordsTemp["ssi_batchstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(1);

                                //objMailRecordsTemp.ssi_batchidtxt = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);

                                service.Update(objMailRecordsTemp);
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }

                    #endregion
                }
            }
            FilePath = "";
        }

        if (bDistributionWire == true)
        {
            clsReportTemplate objReportsTemplates = new clsReportTemplate();

            clsDB = new DB();
            string[] SourceFileName = new string[1];
            string DestinationPath1 = "";
            string ReportOpFolder1 = "";
            string PdfFinalPath1 = "";
            objReportsTemplates.MailID = ViewState["MailId"].ToString();


            objReportsTemplates.TemplateID = ViewState["TemplateId"].ToString();


            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder1 = Request.MapPath("\\Advent Report\\ExcelTemplate\\BATCH REPORTS\\Distribution Wire Instruction.pdf");  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
                ReportOpFolder1 = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports) + "\\Distribution Wire Instruction.pdf";


            //DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters("DistributionWire.pdf");

            string sql1 = "SP_S_LegalEntityContact @MailIDList=" + ViewState["MailId"].ToString();
            DataSet DSLegalEntity1 = clsDB.getDataSet(sql1);
            string Template1 = "";
            for (int i = 0; i < DSLegalEntity1.Tables[0].Rows.Count; i++)
            {
                if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]) != "")
                {
                    objReportsTemplates.LegalEntityID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]);

                }

                if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
                {
                    objReportsTemplates.ContactNameID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]);

                }

                PdfFinalPath1 = objReportsTemplates.DistributionWireInstruction();

                if (Template1 != "")
                {
                    Template1 = Template1 + "|" + PdfFinalPath1;
                }
                else
                {
                    Template1 = "|" + PdfFinalPath1;
                }



            }

            DestinationPath = ReportOpFolder1;
            SourceFileName[0] = Template1;


            #region Merge PDF


            Template1 = Template1.Substring(1, Template1.Length - 1);
            strTemplate = Template1.Split('|');

            for (int l = 0; l < strTemplate.Length; l++)
            {
                if (FilePath != "")
                {
                    if (strTemplate[l] != "")
                    {
                        FilePath = FilePath + "|" + strTemplate[l];
                    }
                }
                else
                {
                    if (strTemplate[l] != "")
                    {
                        FilePath = "|" + strTemplate[l];
                    }
                }
            }

            #endregion

            #region Delete Files

            if (FilePath != "")
            {

                FilePath = FilePath.Substring(1, FilePath.Length - 1);
                string[] strPath = FilePath.Split('|');
                if (DestinationPath != "")
                {
                    PDFMerge PDF = new PDFMerge();
                    PDF.MergeFiles(DestinationPath, strPath);
                }
                else
                {
                    PDFMerge PDF = new PDFMerge();
                    PDF.MergeFiles(DestinationPath1, strPath);
                }


                string filesToDelete = "*.pdf";
                string Path = "C:\\AdventReport\\BatchReport\\ExcelTemplate\\pdfOutput";
                string Path1 = "C:\\AdventReport\\ExcelTemplate\\pdfOutput";
                //string Path = Request.MapPath("\\Advent Report\\BatchReport\\ExcelTemplate\\pdfOutput\\");
                //string Path1 = Request.MapPath("\\Advent Report\\ExcelTemplate\\pdfOutput\\");
                if (Path != "")
                {
                    if (Directory.Exists(Path))
                    {
                        string[] fileList = System.IO.Directory.GetFiles(Path, filesToDelete);

                        foreach (string file in fileList)
                        {
                            try
                            {
                                System.IO.File.Delete(file);
                            }
                            catch
                            { }

                            //sResult += "\n" + file + "\n";
                        }
                    }
                }

                if (Path1 != "")
                {
                    if (Directory.Exists(Path1))
                    {
                        string[] fileList = System.IO.Directory.GetFiles(Path1, filesToDelete);

                        foreach (string file in fileList)
                        {
                            try
                            {
                                System.IO.File.Delete(file);
                            }
                            catch
                            { }
                        }
                    }
                }
            }

            #endregion

            if (FilePath != "")
            {
                ViewState["DistributionWireInstFile"] = DestinationPath;

            }

            lblError.Text = "Distribution Wire Letter pdf generated Successfully.";
            lblError.Visible = true;
            Button1.Visible = true;
            // Label1.Visible = false;

        }

        if (bBillingInvoice == true)
        {

            BillingBatchProcess(service, DSLegalEntity, DestinationPath);


            if (FilePath != "")
            {
                ViewState["BillingInvoiceInstFile"] = DestinationPath;
                #region Delete Files

                if (FilePath != "" && FilePath != null)
                {

                    FilePath = FilePath.Substring(1, FilePath.Length - 1);
                    string[] strPath = FilePath.Split('|');
                    PDFMerge PDF = new PDFMerge();
                    PDF.MergeFiles(DestinationPath, strPath);

                    string filesToDelete = "*.pdf";
                    string Path = "C:\\AdventReport\\BatchReport\\ExcelTemplate\\pdfOutput";
                    string Path1 = "C:\\AdventReport\\ExcelTemplate\\pdfOutput";
                    //string Path = Request.MapPath("\\Advent Report\\BatchReport\\ExcelTemplate\\pdfOutput\\");
                    //string Path1 = Request.MapPath("\\Advent Report\\ExcelTemplate\\pdfOutput\\");
                    if (Path != "")
                    {
                        if (Directory.Exists(Path))
                        {
                            string[] fileList = System.IO.Directory.GetFiles(Path, filesToDelete);

                            foreach (string file in fileList)
                            {
                                try
                                {
                                    System.IO.File.Delete(file);
                                }
                                catch
                                { }

                                //sResult += "\n" + file + "\n";
                            }
                        }
                    }

                    if (Path1 != "")
                    {
                        if (Directory.Exists(Path1))
                        {
                            string[] fileList = System.IO.Directory.GetFiles(Path1, filesToDelete);

                            foreach (string file in fileList)
                            {
                                try
                                {
                                    System.IO.File.Delete(file);
                                }
                                catch
                                { }
                            }
                        }
                    }
                }

                #endregion
            }
            //added 4_21_2020 - excel to show up reports that failed to generate

            int Ssi_MailingID = (int)ViewState["Ssi_MailingID"];
            string sql1 = "SP_S_BillingInvoiceException @MailIDList=" + ViewState["MailId"].ToString() + ",@Ssi_MailingID=" + Ssi_MailingID;
            DataSet DSLegalEntitytemp1 = clsDB.getDataSet(sql1);
            DataTable dtException = DSLegalEntitytemp1.Tables[0];
            int rowCount = dtException.Rows.Count;
            sw.WriteLine("sql1" + sql1);
            sw.WriteLine("rowCount" + rowCount.ToString() + DateTime.Now.ToString());

            //Response.Write("SQL :" + sql1);
            //  Response.Write("rowCount :" + rowCount);

            if (rowCount > 0)
            {
                string ExcelFilePath = GenerateExcel(DSLegalEntitytemp1);
                if (ExcelFilePath != "")
                {
                    lbtnExceptionReport.Visible = true;
                    ViewState["ExcetionReportPath"] = ExcelFilePath;
                }
            }

            int finalcount = Convert.ToInt32(ViewState["intResult"]);
            sw.WriteLine("finalcount:" + finalcount);
          //  Response.Write("finalcount:" + finalcount);
            if (finalcount > 0)
                lblError.Text = "Billing Invoice pdf generated Successfully.";
            else
                lblError.Text = "No record processed.";
            lblError.Visible = true;
            Button1.Visible = true;
            //  Label1.Visible = false;
        }


    }
    private void Download_File(string FilePath, string FileName)
    {
        Response.ContentType = ContentType;
        Response.AppendHeader("Content-Disposition", "attachment; filename=" + FileName);
        Response.WriteFile(FilePath);
        Response.End();
    }
    public string GenerateExcel(DataSet ds)
    {
        string Server = AppLogic.GetParam(AppLogic.ConfigParam.Server);
        try
        {
            //string aod = txtAUMDate.Text;
            DateTime dAsofDate = DateTime.Now;



            if (!Directory.Exists(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder"))
            {
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder");
            }

            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            string strhour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMin = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSec = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
            string append_timestamp = strMonth + "_" + strDay + "_" + strYear + "_" + strhour + "_" + strMin + "_" + strSec;
            String lsFileNamforFinalXls = string.Empty;
            if (Server.ToLower() == "prod")
            {
                lsFileNamforFinalXls = "ExceptionReport" + "_" + append_timestamp;
            }
            else
            {
                lsFileNamforFinalXls = "ExceptionReport" + "_" + append_timestamp + "_test";
            }


            string ExcelFilePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls + ".xlsx";

            if (System.IO.File.Exists(ExcelFilePath))
            {
                System.IO.File.Delete(ExcelFilePath);
            }


            if (ds.Tables.Count > 0)
            {

                #region Spire License Code
                string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
                Spire.License.LicenseProvider.SetLicenseKey(License);
                Spire.License.LicenseProvider.LoadLicense();
                #endregion

                // string SheetNme = ds.Tables[0].Rows[0][0].ToString();

                Workbook book = new Workbook();
                book.Version = ExcelVersion.Version2016;
                Worksheet sheet = book.Worksheets[0];
                //  sheet.Name = SheetNme;
                sheet.Range[1, 1, 1, ds.Tables[0].Columns.Count].Style.Font.IsBold = true;

                sheet.InsertDataTable(ds.Tables[0], true, 1, 1);

                sheet.Range[2, 12, ds.Tables[0].Rows.Count + 1, 12].NumberFormat = "$ #,##0.00_);($ #,##0.00)";
                sheet.Range[2, 13, ds.Tables[0].Rows.Count + 1, 13].NumberFormat = "$ #,##0.00_);($ #,##0.00)";
                sheet.Range[2, 14, ds.Tables[0].Rows.Count + 1, 14].NumberFormat = "$ #,##0.00_);($ #,##0.00)";

                sheet.Range[1, 1, ds.Tables[0].Rows.Count + 1, ds.Tables[0].Columns.Count].AutoFitColumns();
                sheet.Range[1, 1, ds.Tables[0].Rows.Count + 1, ds.Tables[0].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;

                book.SaveToFile(ExcelFilePath);
                //string vContain = "Excel Report Generated Succesfully ";
            }
            return ExcelFilePath;

        }
        catch (Exception e)
        {

            //string vContain = "Excel Report Genration Fail,  Error " + e.ToString();

            return "";
        }
    }
    //added on 08/07/2019 by brijesh
    public void BillingBatchProcess(IOrganizationService service, DataSet DSLegalEntity, string DestinationPath)
    {
        clsReportTemplate objReportsTemplates = new clsReportTemplate();
        clsDB = new DB();
        string[] SourceFileName = new string[1];
        string DestinationPath1 = "";
        string ReportOpFolder1 = "";
        string PdfFinalPath1 = "";
        objReportsTemplates.MailID = ViewState["MailId"].ToString();
        objReportsTemplates.TemplateID = ViewState["TemplateId"].ToString();

        int intResult = 0;
        string UserId = string.Empty;

        //DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters("DistributionWire.pdf");

        string sql1 = "SP_S_LegalEntityContact @MailIDList=" + ViewState["MailId"].ToString();
        DataSet DSLegalEntitytemp = clsDB.getDataSet(sql1);
        //  DataView dv = DSLegalEntitytemp.Tables[0].DefaultView;
        //dv.Sort = "HouseHoldIDName,ssi_contactfullnameidname,ssi_legalentitynameidname ASC";
        // dv.Sort = "HouseHoldIDName,ssi_CustomBillingREF ASC";
        //  DataTable dtnew = dv.ToTable();
        DataSet DSLegalEntity1 = new DataSet();
        DSLegalEntity1 = DSLegalEntity.Copy();
        // DSLegalEntity1.Tables.Add(dtnew);
        string Template1 = "";
        for (int i = 0; i < DSLegalEntity1.Tables[0].Rows.Count; i++)
        {
            if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]) != "")
            {
                objReportsTemplates.LegalEntityID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]);
            }

            if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
            {
                objReportsTemplates.ContactNameID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]).Replace("'", "''");
            }
            if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_aumasofdate"]) != "")
                objReportsTemplates.AUMAsOfDate = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_aumasofdate"]);
            else
                objReportsTemplates.AUMAsOfDate = "";

            if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["HouseHoldID"]) != "")
                objReportsTemplates.HHId = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["HouseHoldID"]).Replace("'", "''");
            else
                objReportsTemplates.HHId = "";

            if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["DateRange"]) != "")
                ViewState["DateRange"] = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["DateRange"]);
            else
                ViewState["DateRange"] = "";

            if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_billingname"]) != "")
                objReportsTemplates.BillingName = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_billingname"]).Replace("'", "''");
            else
                objReportsTemplates.BillingName = "";


            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) != "")
            {
                Random rnd = new Random();
                int rndNum = rnd.Next(1000, 9999);
                string RandomNumStr = "#" + rndNum.ToString();
                string ConsolidatePdfFileName = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr + ".pdf";
                ConsolidatePdfFileName = GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);


                string ReportOpFolder = string.Empty;
                ////string ReportOpFolder = "\\\\Fs01\\_ops_C_I_R_group\\Quarterly_Reports\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports); // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

                if (Request.Url.AbsoluteUri.Contains("localhost"))
                {
                    ReportOpFolder = Request.MapPath("..\\ExcelTemplate\\BATCH REPORTS\\");  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
                }
                else
                    ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);
                DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);
                string strFinalPath = ConsolidatePdfFileName;// Server.MapPath("") + @"\ExcelTemplate\DTS\" + DateTime.Now.ToString("ddMMMyyyymmss") + LegalEntityId + ".pdf"; //FileName

            }


            bool billingPdfFlag = true;
            PdfFinalPath1 = objReportsTemplates.GetBillingInvoice();

            try
            {
                File.Copy(PdfFinalPath1, DestinationPath, true);
            }
            catch (Exception ex)
            {
                billingPdfFlag = false;
            }

            if (billingPdfFlag && PdfFinalPath1 != "")
            {
                #region Create Batch

                //objBatch = new ssi_batch();
                Entity objBatch = new Entity("ssi_batch");


                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_Filename"]) != "")
                {
                    //objBatch.ssi_batchdisplayfilename = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr + ".pdf";
                    objBatch["ssi_batchdisplayfilename"] = GeneralMethods.RemoveSpecialCharacters(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_Filename"])) + RandomNumStr + ".pdf";
                }

                if (DestinationPath != "")
                {
                    //objBatch.ssi_batchfilename = DestinationPath;
                    objBatch["ssi_batchfilename"] = DestinationPath;
                }


                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) != "")
                {
                    //objBatch.ssi_householdid = new Lookup();
                    //objBatch.ssi_householdid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"]));
                    objBatch["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"])));
                }

                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"]) != "")
                {
                    //objBatch.ssi_contactid = new Lookup();
                    //objBatch.ssi_contactid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]));
                    objBatch["ssi_contactid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"])));
                }

                //objBatch.ssi_reporttrackerstatus = new Picklist();
                //objBatch.ssi_reporttrackerstatus.Value = 12;//Initial Review
                objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(12);
                #region OLD ssi_sharepointreportfolder and Clientportal Fields
                //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHold"]) != "" && Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]) != "")
                //{
                //    HouseHold = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHold"]).Replace(" Family", "");
                //  ////  objBatch.ssi_sharepointreportfolder = "http://sp02/ClientServ/Documents/Clients/Active/" + HouseHold + "/Correspondence/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                //  //  objBatch.ssi_sharepointreportfolder = "https://greshampartners.sharepoint.com/clientserv/Documents/Clients/Active/" + HouseHold + "/Correspondence/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                //    objBatch["ssi_sharepointreportfolder"] = "https://greshampartners.sharepoint.com/clientserv/Documents/Clients/Active/" + HouseHold + "/Correspondence/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                //}
                #endregion
                #region NEW CS Sharepoin Changes - added 2_21_2019
                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) != "")
                {
                    objBatch["ssi_cshouseholdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"])));
                }
                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["YearFromFile"]) != "")
                {
                    objBatch["ssi_year"] = Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["YearFromFile"]));
                }
                objBatch["ssi_spsitetype"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000000);
                #endregion
                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["HouseHold"]) != "" && Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["YearFromFile"]) != "")
                {
                    // HouseHold = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["HouseHold"]).Replace(" Family", "");
                    // objBatch.ssi_clientportalfolder = "http://sp02/Client Portal/Documents/" + HouseHold + "/Investment Activity/NonMarketable/" + Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["YearFromFile"]);
                    if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_clientreportfolder"]) != "")
                    {
                        //objBatch.ssi_clientportalfolder = "https://greshampartners.sharepoint.com/ClientPortal/Documents taxonomy";
                        //objBatch["ssi_clientportalfolder"] = "https://greshampartners.sharepoint.com/ClientPortal/Documents taxonomy";
                        objBatch["ssi_clientportalfolder"] = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL)+"/Documents taxonomy";
                    }
                }


                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]) != "")
                {
                    //objBatch.ssi_name = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]);
                    objBatch["ssi_name"] = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]);
                }

                //Advisor Approval Required

                //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "0" || Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "" || Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "False".ToUpper())
                //{
                //    objBatch.ssi_advisorapprovalreqd.Value = false;
                //}
                //else 
                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_advisorapprovalreqd"]) == "1" || Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "True".ToUpper())
                {
                    //objBatch.ssi_advisorapprovalreqd = new CrmBoolean();
                    //objBatch.ssi_advisorapprovalreqd.Value = true;
                    //objBatch["ssi_advisorapprovalreqd"] = true;

                    objBatch["ssi_advisorapprovalreqd"] = false;
                }

                //objBatch.ssi_reporttracker = new CrmBoolean();
                //objBatch.ssi_reporttracker.Value = true;
                objBatch["ssi_reporttracker"] = true;


                //objBatch.ssi_type = new Picklist();
                //objBatch.ssi_type.Value = 4;//Merge
                objBatch["ssi_type"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                //objBatch.ssi_approvalreqd = new CrmBoolean();
                //objBatch.ssi_approvalreqd.Value = true;
                //objBatch["ssi_approvalreqd"] = true;

                // objBatch["ssi_approvalreqd"] = true;
                objBatch["ssi_advisorapprovalreqd"] = true;
                UserId = GetcurrentUser();

                if (UserId != "" && UserId != null)
                {
                    //objBatch.ssi_createdbycustomid = new Lookup();
                    //objBatch.ssi_createdbycustomid.Value = new Guid(UserId);
                    objBatch["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                }

                if (DSLegalEntity.Tables[1].Rows[i]["ssi_clientportalname"].ToString() != "")
                {

                    //objBatch.ssi_clientportalname = DSLegalEntity.Tables[0].Rows[i]["ssi_clientportalname"].ToString();
                    objBatch["ssi_clientportalname"] = DSLegalEntity.Tables[1].Rows[i]["ssi_clientportalname"].ToString();

                }



                // ssi_billinginvoiceid
                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_billinginvoiceId"]) != "")
                {
                    objBatch["ssi_billinginvoiceid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_billinginvoice", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_billinginvoiceId"])));

                }

                // ssi_billingprimaryid
                if (Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_billingprimaryid"]) != "")
                {
                    objBatch["ssi_billingid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_billing", new Guid(Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["Ssi_billingprimaryid"])));

                }

                service.Create(objBatch);
                intResult++;
                ViewState["intResult"] = intResult;

                if (intResult > 0)
                {
                    clsDB = new DB();

                    string HouseHoldId = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) == "" ? "null" : "'" + Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_householdid"]) + "'";
                    string strContactId = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"]) == "" ? "null" : "'" + Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["ssi_contactfullnameid"]) + "'";
                    string BatchName = Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]) == "" ? "null" : "'" + Convert.ToString(DSLegalEntity.Tables[1].Rows[i]["BatchName"]).Replace("'", "''") + "'";

                    string strsql = "SP_S_BatchID @ssi_householdid=" + HouseHoldId + ",@Ssi_ContactId=" + strContactId + ",@Ssi_name=" + BatchName;
                    DataSet DSBatch = clsDB.getDataSet(strsql);

                    if (DSBatch.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]) != "")
                        {
                            #region Update Batch Owner

                            objBatch["ssi_batchid"] = new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]));

                            AssignRequest assignRequest = new AssignRequest
                            {
                                Assignee = new EntityReference("systemuser", new Guid(AppLogic.GetParam(AppLogic.ConfigParam.OpsReporting))),
                                Target = new EntityReference("ssi_batch", new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"])))
                            };

                            service.Execute(assignRequest);
                            service.Update(objBatch);

                            #endregion

                            #region Create New Mail Records

                            //objMailRecords = new ssi_mailrecords();
                            Entity objMailRecords = new Entity("ssi_mailrecords");

                            // BatchId Lookup 
                            if (Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]) != "")
                            {
                                //objMailRecords.ssi_batchid = new Lookup();
                                //objMailRecords.ssi_batchid.Value = new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]));
                                objMailRecords["ssi_batchid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_batch", new Guid(Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"])));

                            }

                            // Batch Id text
                            if (Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]) != "")
                            {
                                //objMailRecords.ssi_batchidtxt = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);
                                objMailRecords["ssi_batchidtxt"] = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);
                            }

                            //Batch Name 
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]) != "")
                            {
                                //objMailRecords.ssi_batchnametxt = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]);
                                objMailRecords["ssi_batchnametxt"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["BatchName"]);
                            }

                            //Name
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Name"]) != "")
                            {
                                //objMailRecords.ssi_name = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Name"]);
                                objMailRecords["ssi_name"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Name"]);
                            }

                            //Mail Status
                            //objMailRecords.ssi_mailstatus = new Picklist();
                            //objMailRecords.ssi_mailstatus.Value = 1;//Pending
                            objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(1);

                            //Mail Type
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Type"]) != "")
                            {
                                //objMailRecords.ssi_mailtypeid = new Lookup();
                                //objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Type"]));
                                objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Type"])));
                            }


                            //Contact
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
                            {
                                //objMailRecords.ssi_contactfullnameid = new Lookup();
                                //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"]));
                                objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_contactfullnameid"])));
                            }

                            //HouseHold
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"]) != "")
                            {
                                //objMailRecords.ssi_accountid = new Lookup();
                                //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"]));
                                objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_householdid"])));
                            }


                            //ssi_legalentitynameid *
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"]) != "")
                            {
                                //objMailRecords.ssi_legalentitynameid = new Lookup();
                                //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"]));
                                objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_legalentitynameid"])));

                            }


                            //Salutation
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Salutation"]) != "")
                            {
                                //objMailRecords.ssi_salutation_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Salutation"]);
                                objMailRecords["ssi_salutation_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Salutation"]);
                            }

                            // FullName
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Full Name"]) != "")
                            {
                                //objMailRecords.ssi_fullname_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Full Name"]);
                                objMailRecords["ssi_fullname_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Full Name"]);
                            }

                            //Address Line 1
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 1"]) != "")
                            {
                                //objMailRecords.ssi_addressline1_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 1"]);
                                objMailRecords["ssi_addressline1_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 1"]);
                            }

                            //Address Line 2
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 2"]) != "")
                            {
                                //objMailRecords.ssi_addressline2_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 2"]);
                                objMailRecords["ssi_addressline2_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 2"]);
                            }

                            //Address Line 3
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 3"]) != "")
                            {
                                //objMailRecords.ssi_addressline3_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 3"]);
                                objMailRecords["ssi_addressline3_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Address Line 3"]);
                            }

                            //City
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["City"]) != "")
                            {
                                //objMailRecords.ssi_city_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["City"]);
                                objMailRecords["ssi_city_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["City"]);
                            }

                            //State/Province
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["State/Province"]) != "")
                            {
                                //objMailRecords.ssi_stateprovince_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["State/Province"]);
                                objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["State/Province"]);
                            }

                            //Zip Code
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Zip Code"]) != "")
                            {
                                //objMailRecords.ssi_zipcode_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Zip Code"]);
                                objMailRecords["ssi_zipcode_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Zip Code"]);
                            }

                            //Country/Region
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Country/Region"]) != "")
                            {
                                //objMailRecords.ssi_countryregion_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Country/Region"]);
                                objMailRecords["ssi_countryregion_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Country/Region"]);
                            }


                            //Ssi_MailingID
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]) != "")
                            {
                                //objMailRecords.ssi_mailingid = new CrmNumber();
                                //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]);
                                objMailRecords["ssi_mailingid"] = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["Ssi_MailingID"]);
                            }


                            //AsOfDate
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["AsOfDate"]) != "")
                            {
                                //objMailRecords.ssi_asofdate = new CrmDateTime();
                                //objMailRecords.ssi_asofdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["AsOfDate"]);
                                objMailRecords["ssi_asofdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["AsOfDate"]));
                            }

                            //Dear
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Dear"]) != "")
                            {
                                //objMailRecords.ssi_dear_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Dear"]);
                                objMailRecords["ssi_dear_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Dear"]);
                            }

                            //Spouse/Partner
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Spouse/Partner"]) != "")
                            {
                                //objMailRecords.ssi_spousepart_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Spouse/Partner"]);
                                objMailRecords["ssi_spousepart_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Spouse/Partner"]);
                            }

                            //Mail
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail"]) != "")
                            {
                                //objMailRecords.ssi_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail"]);
                                objMailRecords["ssi_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail"]);
                            }

                            //Mail Preference
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Preference"]) != "")
                            {
                                //objMailRecords.ssi_mailpreference_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Preference"]);
                                objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Mail Preference"]);
                            }

                            //Status Reason
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Status Reason"]) != "")
                            {
                                //objMailRecords.ssi_status = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Status Reason"]);
                                objMailRecords["ssi_status"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Status Reason"]);
                            }

                            //Contact Owner First Name
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner First Name"]) != "")
                            {
                                //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner First Name"]);
                                objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner First Name"]);
                            }


                            //Contact Owner Last Name
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner Last Name"]) != "")
                            {
                                //objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner Last Name"]);
                                objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Contact Owner Last Name"]);
                            }

                            //Household Owner First Name
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner First Name"]) != "")
                            {
                                //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner First Name"]);
                                objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner First Name"]);
                            }

                            //Household Owner Last Name
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner Last Name"]) != "")
                            {
                                //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner Last Name"]);
                                objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Household Owner Last Name"]);
                            }

                            //Secondary Owner First Name
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
                            {
                                //objMailRecords.ssi_secownerfname_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner First Name"]);
                                objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner First Name"]);
                            }

                            //Secondary Owner Last Name
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
                            {
                                //objMailRecords.ssi_secownerlname_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                                objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_wireasofdate"]) != "")
                            {
                                //objMailRecords.ssi_wireasofdate = new CrmDateTime();
                                //objMailRecords.ssi_wireasofdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_wireasofdate"]);
                                objMailRecords["ssi_wireasofdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_wireasofdate"]));
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_letterdate"]) != "")
                            {
                                //objMailRecords.ssi_letterdate = new CrmDateTime();
                                //objMailRecords.ssi_letterdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_letterdate"]);
                                objMailRecords["ssi_letterdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_letterdate"]));
                            }



                            //CustomMailPreference
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["CustomMailPreference"]) != "")
                            {
                                //objMailRecords.ssi_mailpreference_mail = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["CustomMailPreference"]);
                                objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["CustomMailPreference"]);
                            }


                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) != "")
                            {
                                //objMailRecords.ssi_filename = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr;
                                objMailRecords["ssi_filename"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_Filename"]) + RandomNumStr;
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]) != "")
                            {
                                //objMailRecords.ssi_fundname = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]);
                                objMailRecords["ssi_fundname"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]);
                            }

                            //if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_fundname"]) != "")

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_anzianoid"]) != "")
                            {
                                //objMailRecords.ssi_anzianoid = new CrmNumber();
                                //objMailRecords.ssi_anzianoid.Value = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["ssi_anzianoid"]);
                                objMailRecords["ssi_anzianoid"] = Convert.ToInt32(DSLegalEntity.Tables[0].Rows[i]["ssi_anzianoid"]);
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_tnrid_nv"]) != "")
                            {
                                //objMailRecords.ssi_tnrid_nv = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_tnrid_nv"]);
                                objMailRecords["ssi_tnrid_nv"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_tnrid_nv"]);
                            }

                            //*** Records for Billing Start ***//

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_AUM"]) != "")
                            {
                                //objMailRecords.ssi_aum = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_AUM"]);
                                objMailRecords["ssi_aum"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_AUM"]);
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_AUMasofdate"]) != "")
                            {
                                //objMailRecords.ssi_aumasofdate = new CrmDateTime();
                                //objMailRecords.ssi_aumasofdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_AUMasofdate"]);
                                objMailRecords["ssi_aumasofdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_AUMasofdate"]));
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_annualfeebilling"]) != "")
                            {
                                //objMailRecords.ssi_annualfeebilling = new CrmMoney();
                                //objMailRecords.ssi_annualfeebilling.Value = Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_annualfeebilling"]));
                                objMailRecords["ssi_annualfeebilling"] = new Money(Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_annualfeebilling"])));
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeratebilling"]) != "")
                            {
                                //objMailRecords.ssi_feeratebilling = new CrmFloat();
                                //objMailRecords.ssi_feeratebilling.Value = Convert.ToDouble(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeratebilling"]));
                                objMailRecords["ssi_feeratebilling"] = Convert.ToDouble(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeratebilling"]));
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingidbilling"]) != "")
                            {
                                //objMailRecords.ssi_billingid = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingidbilling"]);
                                objMailRecords["ssi_billingid"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingidbilling"]);
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_quarterlyfeebilling"]) != "")
                            {
                                //objMailRecords.ssi_quarterlyfeebilling = new CrmMoney();
                                //objMailRecords.ssi_quarterlyfeebilling.Value = Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_quarterlyfeebilling"]));
                                objMailRecords["ssi_quarterlyfeebilling"] = new Money(Convert.ToDecimal(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_quarterlyfeebilling"])));

                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_invoicedatebilling"]) != "")
                            {
                                //objMailRecords.ssi_invoicedatebilling = new CrmDateTime();
                                //objMailRecords.ssi_invoicedatebilling.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_invoicedatebilling"]);
                                objMailRecords["ssi_invoicedatebilling"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_invoicedatebilling"]));
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_autodebitdate"]) != "")
                            {
                                //objMailRecords.ssi_autodebitdate = new CrmDateTime();
                                //objMailRecords.ssi_autodebitdate.Value = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_autodebitdate"]);
                                objMailRecords["ssi_autodebitdate"] = Convert.ToDateTime(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_autodebitdate"]));
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_CustomBillingREF"]) != "")
                            {
                                //objMailRecords.ssi_custombillingref = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_CustomBillingREF"]);
                                objMailRecords["ssi_custombillingref"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_CustomBillingREF"]);
                            }

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]) != "")
                            {
                                //objMailRecords.ssi_billingname = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]);
                                objMailRecords["ssi_billingname"] = Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_billingname"]);
                            }

                            #region New Field added on 08/08/2019 for billing


                            // ssi_billinginvoiceid
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_billinginvoiceId"]) != "")
                            {
                                objMailRecords["ssi_billinginvoiceid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_billinginvoice", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_billinginvoiceId"])));

                            }

                            // ssi_billingprimaryid
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_billingprimaryid"]) != "")
                            {
                                objMailRecords["ssi_billingprimaryid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_billing", new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_billingprimaryid"])));

                            }

                            //ssi_feeonfirst25mminbps

                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_feeonfirst25mminbps"]) != "")
                            {

                                objMailRecords["ssi_feeonfirst25mminbps"] = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_feeonfirst25mminbps"]);

                            }

                            // ssi_maximumfeeasa
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_maximumfeeasa"]) != "")
                            {
                                objMailRecords["ssi_maximumfeeasa"] = Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_maximumfeeasa"]);
                            }

                            // ssi_minimumfeein
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_minimumfeein"]) != "")
                            {
                                objMailRecords["ssi_minimumfeein"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_minimumfeein"]));

                            }

                            //  ssi_totalbillableassets
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_totalbillableassets"]) != "")
                            {
                                objMailRecords["ssi_totalbillableassets"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_totalbillableassets"]));

                            }

                            //ssi_securityfeeaum
                            if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["ssi_securityfeeaum"]) != "")
                            {
                                objMailRecords["ssi_securityfeeaum"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(DSLegalEntity.Tables[0].Rows[i]["ssi_securityfeeaum"]));

                            }


                            #endregion

                            //objMailRecords.ssi_ir_status = new Picklist();
                            //objMailRecords.ssi_ir_status.Value = 1;//Pending;
                            objMailRecords["ssi_ir_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(1);

                            UserId = GetcurrentUser();


                            if (UserId != "")
                            {
                                //objMailRecords.ssi_createdbycustomid = new Lookup();
                                //objMailRecords.ssi_createdbycustomid.Value = new Guid(UserId);
                                objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                            }

                            service.Create(objMailRecords);

                            #endregion

                            intResult++;

                        }
                    }


                }

                #endregion


                #region Update completed flag on billing Invoicce
                Entity objBillingInvoice = new Entity("ssi_billinginvoice");

                if (Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_billinginvoiceId"]) != "")
                {
                    objBillingInvoice["ssi_billinginvoiceid"] = new Guid(Convert.ToString(DSLegalEntity.Tables[0].Rows[i]["Ssi_billinginvoiceId"]));

                    objBillingInvoice["ssi_completed"] = true;
                    service.Update(objBillingInvoice);
                    sw.WriteLine("updated" + i);
                }
                #endregion
            }

        }


        //#region Update completed flag on billing invoice

        ////  string strbillingComplteFlag = "SP_S_MailRecordsTempID_List @MailIDList=" + ViewState["MailId"].ToString();// +",@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID='" + ContactId + "'";
        //DataSet updatebillingCompltFlag = DSLegalEntity1.Copy();
        //updatebillingCompltFlag.AcceptChanges();

        //sw.WriteLine("updatebillingCompltFlag.Tables.Count" + updatebillingCompltFlag.Tables.Count);

        //sw.WriteLine("updatebillingCompltFlag.Tables.Count" + updatebillingCompltFlag.Tables[2].Rows.Count);


        //if (updatebillingCompltFlag.Tables.Count > 1 && updatebillingCompltFlag.Tables[2].Rows.Count > 0)
        //{

        //    for (int j = 0; j < updatebillingCompltFlag.Tables[2].Rows.Count; j++)
        //    {
        //        Entity objBillingInvoice = new Entity("ssi_billinginvoice");

        //        if (Convert.ToString(updatebillingCompltFlag.Tables[2].Rows[j]["Ssi_billinginvoiceId"]) != "")
        //        {
        //            objBillingInvoice["ssi_billinginvoiceid"] = new Guid(Convert.ToString(updatebillingCompltFlag.Tables[2].Rows[j]["Ssi_billinginvoiceId"]));

        //            objBillingInvoice["ssi_completed"] = true;
        //            service.Update(objBillingInvoice);
        //            sw.WriteLine("updated" + j);
        //        }
        //    }
        //}
        ////}


        //#endregion

        if (FilePath != "")
        {
            ViewState["BillingInvoiceInstFile"] = DestinationPath;

        }


    }



    public string generateCombinedPDF(String LegalEntityId, String ContactId, String MailId, String DestinationFileName, String Reportname, String CustomFlg, String CustomTemplateType, String TemplateId, int count, String AUMAsOfDate, String HHId, String BillingName)
    {
        string filepdfname = MergeReports(DestinationFileName, Reportname, CustomFlg, LegalEntityId, ContactId, MailId, CustomTemplateType, TemplateId, count, AUMAsOfDate, HHId, BillingName);

        if (filepdfname == "")
        {
            return "";
        }
        else
            return filepdfname;
    }


    public string MergeReports(string DestinationFileName, string ReportName, string CustomFlg, string LegalEntityId, string ContactId, string MailId, string CustomTemplateType, string TemplateId, int count, string AUMAsOfDate, string HHId, string BillingName)
    {
        string strCapitalCall = string.Empty;
        string strCapitalCall1 = string.Empty;
        string strCapitalCall2 = string.Empty;
        string strDistribution1 = string.Empty;
        string strDistribution = string.Empty;

        clsReportTemplate objReportsTemplates = new clsReportTemplate();
        objReportsTemplates.LegalEntityID = LegalEntityId;
        objReportsTemplates.ContactNameID = ContactId;
        objReportsTemplates.MailID = MailId;
        objReportsTemplates.TemplateID = TemplateId;
        objReportsTemplates.AUMAsOfDate = AUMAsOfDate;

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        // = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\Gresham_" + strGUID + ".pdf";
        string[] SourceFileName = new string[1];

        if (ReportName.ToUpper() == "Capital Call Wire Instructions".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
            SourceFileName[0] = objReportsTemplates.CapitalCallWireInstruction();
        if (ReportName.ToUpper() == "SLOA".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
            SourceFileName[0] = objReportsTemplates.GetSLOA();
        else if (ReportName.ToUpper() == "Distribution Wire Instructions".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
        {
            string MailIdList1 = "";

            for (int j = 1; j < 16; j++)
            {
                Control GroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

                if (GroupTemplate != null)
                {
                    DropDownList ddlGroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

                    if (ddlGroupTemplate.SelectedValue != "" && ddlGroupTemplate.SelectedValue != "0")
                    {
                        MailIdList1 = MailIdList1 + "," + ddlGroupTemplate.SelectedValue;
                    }
                }
            }


            string strMailId = MailIdList1.Substring(1, MailIdList1.Length - 1);
            string[] MailId1 = strMailId.Split(',');

            if (MailId1.Length > 1)
            {
                clsDB = new DB();

                string ReportOpFolder = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\";
                string PdfFinalPath = "";

                string sql1 = "SP_S_LegalEntityContact @MailIDList=" + MailId;
                DataSet DSLegalEntity1 = clsDB.getDataSet(sql1);
                string Template1 = "";
                for (int i = 0; i < DSLegalEntity1.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]) != "")
                    {
                        objReportsTemplates.LegalEntityID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]);
                    }

                    if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
                    {
                        objReportsTemplates.ContactNameID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]);
                    }



                    PdfFinalPath = objReportsTemplates.DistributionWireInstruction();

                    if (Template1 != "")
                    {
                        Template1 = Template1 + "|" + PdfFinalPath;
                    }
                    else
                    {
                        Template1 = "|" + PdfFinalPath;
                    }
                }

                SourceFileName[0] = Template1;
            }
        }

        else if (ReportName.ToUpper() == "Distribution Statement".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
            SourceFileName[0] = objReportsTemplates.GetDistributionStatement();
        else if (ReportName.ToUpper() == "Capital Call Statement".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
            SourceFileName[0] = objReportsTemplates.GetCapitalCallStatement();
        else if (CustomTemplateType == "1" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Capital Call Letter
        {
            strCapitalCall = objReportsTemplates.GetCapitalCallStatement();
            strCapitalCall1 = objReportsTemplates.GetCapitalCallLetterStatementCustom();
            strCapitalCall2 = objReportsTemplates.GetCapitalCallLetterStatementCustomRetirement();
            SourceFileName[0] = strCapitalCall1 + "|" + strCapitalCall2 + "|" + strCapitalCall;
        }

        else if (CustomTemplateType == "2" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Distribution Letter
        {
            strDistribution = objReportsTemplates.GetDistributionLetterStatementCustom();
            strDistribution1 = objReportsTemplates.GetDistributionStatement();
            SourceFileName[0] = strDistribution + "|" + strDistribution1;
        }

        else if (CustomTemplateType == "3" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Gresham Advisors General Letter Non-Specific Recipient
            SourceFileName[0] = objReportsTemplates.GetGreshamAdvisorsGLNSRecipient();
        else if (CustomTemplateType == "4" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Gresham General Letter Non-Specific Recipient
            SourceFileName[0] = objReportsTemplates.GetGreshamAdvisorsGLNSRecipient();
        else if (CustomTemplateType == "5" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Gresham Advisors General Letter Regarding Legal Entity
            SourceFileName[0] = objReportsTemplates.GetGreshamAdvisorsGLRLegalEntity();
        else if (CustomTemplateType == "6" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Gresham General Letter Regarding Legal Entity
            SourceFileName[0] = objReportsTemplates.GetGreshamAdvisorsGLRLegalEntity();
        else if (CustomTemplateType == "7" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Gresham Advisors General Letter Regarding Fund
            SourceFileName[0] = objReportsTemplates.GetGreshamAdvisorsGLRFund();
        else if (CustomTemplateType == "8" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Gresham General Letter Regarding Fund
            SourceFileName[0] = objReportsTemplates.GetGreshamAdvisorsGLRFund();
        else if (CustomTemplateType == "9" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Memorandum Regarding Fund
            SourceFileName[0] = objReportsTemplates.GetFundMemoradum();
        else if (CustomTemplateType == "10" && CustomFlg.ToUpper() == "True".ToUpper()) // Custom Template : Memorandum Regarding Fund
            SourceFileName[0] = objReportsTemplates.GetUploadPdf();
        else if (ReportName.ToUpper() == "Billing Invoice".ToUpper() && CustomFlg.ToUpper() == "False".ToUpper())
        {
            string MailIdList1 = "";

            for (int j = 1; j < 16; j++)
            {
                Control GroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

                if (GroupTemplate != null)
                {
                    DropDownList ddlGroupTemplate = ((DropDownList)FindControl("ddlGroupTemplate" + j.ToString()));

                    if (ddlGroupTemplate.SelectedValue != "" && ddlGroupTemplate.SelectedValue != "0")
                    {
                        MailIdList1 = MailIdList1 + "," + ddlGroupTemplate.SelectedValue;
                    }
                }
            }


            string strMailId = MailIdList1.Substring(1, MailIdList1.Length - 1);
            string[] MailId1 = strMailId.Split(',');

            if (MailId1.Length > 1)
            {
                clsDB = new DB();

                string ReportOpFolder = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\";
                string PdfFinalPath = "";

                string sql1 = "SP_S_LegalEntityContact @MailIDList=" + MailId;
                DataSet DSLegalEntitytemp = clsDB.getDataSet(sql1);
                DataView dv = DSLegalEntitytemp.Tables[0].DefaultView;
                //dv.Sort = "HouseHoldIDName,ssi_contactfullnameidname,ssi_legalentitynameidname ASC";
                dv.Sort = "HouseHoldIDName,ssi_CustomBillingREF ASC";
                DataTable dtnew = dv.ToTable();
                DataSet DSLegalEntity1 = new DataSet();
                DSLegalEntity1.Tables.Add(dtnew);
                string Template1 = "";
                for (int i = 0; i < DSLegalEntity1.Tables[0].Rows.Count; i++)
                {
                    if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]) != "")
                    {
                        objReportsTemplates.LegalEntityID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_legalentitynameid"]).Replace("'", "''");
                    }

                    if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
                    {
                        objReportsTemplates.ContactNameID = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_contactfullnameid"]).Replace("'", "''");
                    }
                    if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_aumasofdate"]) != "")
                        objReportsTemplates.AUMAsOfDate = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_aumasofdate"]);
                    else
                        objReportsTemplates.AUMAsOfDate = "";

                    if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["HouseHoldID"]) != "")
                        objReportsTemplates.HHId = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["HouseHoldID"]).Replace("'", "''");
                    else
                        objReportsTemplates.HHId = "";

                    if (Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_billingname"]) != "")
                    {
                        objReportsTemplates.BillingName = Convert.ToString(DSLegalEntity1.Tables[0].Rows[i]["ssi_billingname"]).Replace("'", "''");
                    }


                    //   PdfFinalPath = objReportsTemplates.GetBillingInvoice();

                    if (Template1 != "")
                    {
                        Template1 = Template1 + "|" + PdfFinalPath;
                    }
                    else
                    {
                        Template1 = "|" + PdfFinalPath;
                    }
                }

                SourceFileName[0] = Template1;
            }
            //SourceFileName[0] = objReportsTemplates.GetBillingInvoice();

        }

        if (SourceFileName[0] != "No Record Found")
        {
            DestinationFileName = SourceFileName[0];
        }
        else
        {
            DestinationFileName = "";
        }
        return DestinationFileName;
    }


    //public static CrmService GetCrmService(string crmServerUrl, string organizationName)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    // string username = WindowsIdentity.GetCurrent().Name;

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
        string UserID = string.Empty;
        string strName = string.Empty;
        //  System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        //Response.Write("p.Identity.Name:" + strName + "<br/><br/>");
        //strName = HttpContext.Current.User.Identity.Name.ToString();
        //Response.Write("HttpContext.Current.User.Identity.Name:" + strName + "<br/><br/>");
        //strName = Request.ServerVariables["AUTH_USER"]; //Finding with name


        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            strName = "corp\\gbhagia";
        }
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            strName = claimsIdentity.Name;

        }


        //Response.Write("AUTH_USER:" + strName + "<br/><br/>");
        //////////
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        string sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        //Response.Write(strName + "<br/><br/>");
        //Response.Write(Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]));
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
        }
        else
        {
            return UserID = "";
        }
    }

    protected void lbtnExceptionReport_Click(object sender, EventArgs e)
    {
        string SampleFilePath = ViewState["ExcetionReportPath"].ToString();
        Download_File(SampleFilePath, "ExceptionReport.xlsx");
    }
}
