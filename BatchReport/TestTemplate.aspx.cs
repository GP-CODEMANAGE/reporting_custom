using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Winthusiasm.HtmlEditor;
using System.IO;
//using CrmSdk;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using iTextSharp.text.html;
using System.Data.SqlClient;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class TestTemplate : System.Web.UI.Page
{          
    public int liPageSize = 29;
    bool FundSpecific = false;
    bool bProceed = true;
    public int EditRows = 0;
    string strDescription;
    int num = 0;
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = null;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName;
    public string TemplateName = string.Empty;
    public string TemplateId = string.Empty;
    public string TemplateFundId = string.Empty;
    public string strAsOfDate = string.Empty;
    protected void Page_Load(object sender, EventArgs e)
    {
        Response.Cache.SetCacheability(HttpCacheability.NoCache); //check cahability
        if (!IsPostBack)
        {
            BindDropDownList();
            if (rdoYes.Checked == true && hdtblId.Value == "1")
            {
                trFund1.Visible = true;
                HideTR();
            }
        }

       
    }

    private void HideTR()
    {

        if (rdoDynamic.Checked == true)
        {
            trOrientation.Visible = false;
            trSignatureText.Visible = false;
        }
        else if (rdoStatic.Checked == true)
        {
            trOrientation.Visible = true;
            trSignatureText.Visible = true;
        }

        trFileupload.Visible = false;
        lblFileName.Visible = false;
        trFund2.Visible =false;//.Style.Add("display", "none");
        trFund3.Visible = false;//.Style.Add("display", "none");
        trFund4.Visible = false;//.Style.Add("display", "none");
        trFund5.Visible = false;//.Style.Add("display", "none");
        trFund6.Visible = false;//.Style.Add("display", "none");
        trFund7.Visible = false;//.Style.Add("display", "none");
        trFund8.Visible = false;//.Style.Add("display", "none");
        trFund9.Visible = false;//.Style.Add("display", "none");
        trFund10.Visible = false;//.Style.Add("display", "none");
        trFund11.Visible = false;// added 10_16_2018(sasmit)
        trFund12.Visible = false;// added 11_21_2018(sasmit)
        trFund13.Visible = false;// added 11_21_2018(sasmit)
        trFund14.Visible = false;// added 11_21_2018(sasmit)
        trFund15.Visible = false;// added 11_21_2018(sasmit)
        trFund16.Visible = false;// added 11_23_2018(sasmit)
        trFund17.Visible = false;// added 11_23_2018(sasmit)
        trFund18.Visible = false;// added 11_23_2018(sasmit)
        trFund19.Visible = false;// added 11_23_2018(sasmit)
        trFund20.Visible = false;// added 11_23_2018(sasmit)
    }

    public void BindDropDownList()
    {
        Orientation();
        HouseholdOwner();
        TemplateType();
        BindTemplate();
        BindFundType1();
        BindFundType2();
        BindFundType3();
        BindFundType4();
        BindFundType5();
        BindFundType6();
        BindFundType7();
        BindFundType8();
        BindFundType9();
        BindFundType10();
        BindFundType11(); // added 10_16_2018(sasmit)
        BindFundType12(); // added 11_21_2018(sasmit)
        BindFundType13(); // added 11_21_2018(sasmit)
        BindFundType14(); // added 11_21_2018(sasmit)
        BindFundType15(); // added 11_21_2018(sasmit)
        BindFundType16(); // added 11_23_2018(sasmit)
        BindFundType17(); // added 11_23_2018(sasmit)
        BindFundType18(); // added 11_23_2018(sasmit)
        BindFundType19(); // added 11_23_2018(sasmit)
        BindFundType20(); // added 11_23_2018(sasmit)


        //BindHeaderType();
        BindIncSigLine();
        BindFooterType();
        BindFund1();
        BindFund2();
        BindFund3();
        BindFund4();
        BindFund5();
        BindFund6();
        BindFund7();
        BindFund8();
        BindFund9();
        BindFund10();
        BindFund11(); // added 10_16_2018(sasmit)
        BindFund12(); // added 11_21_2018(sasmit)
        BindFund13(); // added 11_21_2018(sasmit)
        BindFund14(); // added 11_21_2018(sasmit)
        BindFund15(); // added 11_21_2018(sasmit)
        BindFund16(); // added 10_23_2018(sasmit)
        BindFund17(); // added 11_23_2018(sasmit)
        BindFund18(); // added 11_23_2018(sasmit)
        BindFund19(); // added 11_23_2018(sasmit)
        BindFund20(); // added 11_23_2018(sasmit)
    }


    public void Orientation()
    {
        string sql = "SP_S_Template_Orientation";
        clsGM.getBindDDL(ddlOrientation, sql, "Orientation", "ID");
    }

    public void HouseholdOwner()
    {
        string sql = "SP_S_Household_Owner";
        clsGM.getListForBindListBox(LstSignText, sql, "HouseholdOwner", "HouseholdOwner");

        LstSignText.Items.Insert(0, "All");
        LstSignText.Items[0].Value = "0";
        LstSignText.SelectedIndex = 0;
    }


    public void TemplateType()
    {
        string sql = "SP_S_TemplateType_List";
        clsGM.getListForBindDDL(ddlTemplateType, sql, "TemplateType", "ID");
    }

    public void BindTemplate()
    {
        string sql = "SP_S_Template @Ssi_Customflg=1";
        clsGM.getBindDDL(ddlTemplate, sql, "ssi_name", "ssi_templateid");

        if (ddlTemplate.SelectedValue == "0")
        {
            ddlTemplate.SelectedItem.Text = "New";
        }
    }


    
    #region Bind Funds

    private void BindFund1()
    {
        object FundTypeId1 = ddlFundType1.SelectedValue == "" || ddlFundType1.SelectedValue == "0" ? "null" : ddlFundType1.SelectedValue;
        string sql1 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId1;
        clsGM.getListForBindDDL(ddlFund1, sql1, "ssi_name", "ssi_FundId");
        ddlFund1.Items.Insert(0, "All");
        ddlFund1.Items[0].Value = "0";
        ddlFund1.SelectedIndex = 0;
    }

    private void BindFund2()
    {
        object FundTypeId2 = ddlFundType1.SelectedValue == "" || ddlFundType2.SelectedValue == "0" ? "null" : ddlFundType2.SelectedValue;
        string sql2 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId2;
        clsGM.getListForBindDDL(ddlFund2, sql2, "ssi_name", "ssi_FundId");
        ddlFund2.Items.Insert(0, "All");
        ddlFund2.Items[0].Value = "0";
        ddlFund2.SelectedIndex = 0;
    }

    private void BindFund3()
    {
        object FundTypeId3 = ddlFundType3.SelectedValue == "" || ddlFundType3.SelectedValue == "0" ? "null" : ddlFundType3.SelectedValue;
        string sql3 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId3;
        clsGM.getListForBindDDL(ddlFund3, sql3, "ssi_name", "ssi_FundId");
        ddlFund3.Items.Insert(0, "All");
        ddlFund3.Items[0].Value = "0";
        ddlFund3.SelectedIndex = 0;
    }

    private void BindFund4()
    {
        object FundTypeId4 = ddlFundType4.SelectedValue == "" || ddlFundType4.SelectedValue == "0" ? "null" : ddlFundType4.SelectedValue;
        string sql4 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId4;
        clsGM.getListForBindDDL(ddlFund4, sql4, "ssi_name", "ssi_FundId");
        ddlFund4.Items.Insert(0, "All");
        ddlFund4.Items[0].Value = "0";
        ddlFund4.SelectedIndex = 0;
    }

    private void BindFund5()
    {
        object FundTypeId5 = ddlFundType5.SelectedValue == "" || ddlFundType5.SelectedValue == "0" ? "null" : ddlFundType5.SelectedValue;
        string sql5 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId5;
        clsGM.getListForBindDDL(ddlFund5, sql5, "ssi_name", "ssi_FundId");
        ddlFund5.Items.Insert(0, "All");
        ddlFund5.Items[0].Value = "0";
        ddlFund5.SelectedIndex = 0;
    }

    private void BindFund6()
    {
        object FundTypeId6 = ddlFundType6.SelectedValue == "" || ddlFundType6.SelectedValue == "0" ? "null" : ddlFundType6.SelectedValue;
        string sql6 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId6;
        clsGM.getListForBindDDL(ddlFund6, sql6, "ssi_name", "ssi_FundId");
        ddlFund6.Items.Insert(0, "All");
        ddlFund6.Items[0].Value = "0";
        ddlFund6.SelectedIndex = 0;
    }

    private void BindFund7()
    {
        object FundTypeId7 = ddlFundType7.SelectedValue == "" || ddlFundType7.SelectedValue == "0" ? "null" : ddlFundType7.SelectedValue;
        string sql7 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId7;
        clsGM.getListForBindDDL(ddlFund7, sql7, "ssi_name", "ssi_FundId");
        ddlFund7.Items.Insert(0, "All");
        ddlFund7.Items[0].Value = "0";
        ddlFund7.SelectedIndex = 0;
    }

    private void BindFund8()
    {
        object FundTypeId8 = ddlFundType8.SelectedValue == "" || ddlFundType8.SelectedValue == "0" ? "null" : ddlFundType8.SelectedValue;
        string sql8 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId8;
        clsGM.getListForBindDDL(ddlFund8, sql8, "ssi_name", "ssi_FundId");
        ddlFund8.Items.Insert(0, "All");
        ddlFund8.Items[0].Value = "0";
        ddlFund8.SelectedIndex = 0;
    }

    private void BindFund9()
    {
        object FundTypeId9 = ddlFundType9.SelectedValue == "" || ddlFundType9.SelectedValue == "0" ? "null" : ddlFundType9.SelectedValue;
        string sql9 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId9;
        clsGM.getListForBindDDL(ddlFund9, sql9, "ssi_name", "ssi_FundId");
        ddlFund9.Items.Insert(0, "All");
        ddlFund9.Items[0].Value = "0";
        ddlFund9.SelectedIndex = 0;
    }

    public void BindFund10()
    {
        object FundTypeId10 = ddlFundType10.SelectedValue == "" || ddlFundType10.SelectedValue == "0" ? "null" : ddlFundType10.SelectedValue;
        string sql10 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId10; ;
        clsGM.getListForBindDDL(ddlFund10, sql10, "ssi_name", "ssi_FundId");
        ddlFund10.Items.Insert(0, "All");
        ddlFund10.Items[0].Value = "0";
        ddlFund10.SelectedIndex = 0;
    }
    public void BindFund11()
    {
        object FundTypeId11= ddlFundType11.SelectedValue == "" || ddlFundType11.SelectedValue == "0" ? "null" : ddlFundType11.SelectedValue;
        string sql11 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId11; ;
        clsGM.getListForBindDDL(ddlFund11, sql11, "ssi_name", "ssi_FundId");
        ddlFund11.Items.Insert(0, "All");
        ddlFund11.Items[0].Value = "0";
        ddlFund11.SelectedIndex = 0;
    }
    public void BindFund12()
    {
        object FundTypeId12 = ddlFundType12.SelectedValue == "" || ddlFundType12.SelectedValue == "0" ? "null" : ddlFundType12.SelectedValue;
        string sql12 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId12; ;
        clsGM.getListForBindDDL(ddlFund12, sql12, "ssi_name", "ssi_FundId");
        ddlFund12.Items.Insert(0, "All");
        ddlFund12.Items[0].Value = "0";
        ddlFund12.SelectedIndex = 0;
    }
    public void BindFund13()
    {
        object FundTypeId13 = ddlFundType13.SelectedValue == "" || ddlFundType13.SelectedValue == "0" ? "null" : ddlFundType13.SelectedValue;
        string sql13 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId13; ;
        clsGM.getListForBindDDL(ddlFund13, sql13, "ssi_name", "ssi_FundId");
        ddlFund13.Items.Insert(0, "All");
        ddlFund13.Items[0].Value = "0";
        ddlFund13.SelectedIndex = 0;
    }
    public void BindFund14()
    {
        object FundTypeId14 = ddlFundType14.SelectedValue == "" || ddlFundType14.SelectedValue == "0" ? "null" : ddlFundType14.SelectedValue;
        string sql14 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId14; ;
        clsGM.getListForBindDDL(ddlFund14, sql14, "ssi_name", "ssi_FundId");
        ddlFund14.Items.Insert(0, "All");
        ddlFund14.Items[0].Value = "0";
        ddlFund14.SelectedIndex = 0;
    }
    public void BindFund15()
    {
        object FundTypeId15 = ddlFundType15.SelectedValue == "" || ddlFundType15.SelectedValue == "0" ? "null" : ddlFundType15.SelectedValue;
        string sql15 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId15; ;
        clsGM.getListForBindDDL(ddlFund15, sql15, "ssi_name", "ssi_FundId");
        ddlFund15.Items.Insert(0, "All");
        ddlFund15.Items[0].Value = "0";
        ddlFund15.SelectedIndex = 0;
    }
    public void BindFund16()
    {
        object FundTypeId16 = ddlFundType16.SelectedValue == "" || ddlFundType16.SelectedValue == "0" ? "null" : ddlFundType16.SelectedValue;
        string sql16 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId16; ;
        clsGM.getListForBindDDL(ddlFund16, sql16, "ssi_name", "ssi_FundId");
        ddlFund16.Items.Insert(0, "All");
        ddlFund16.Items[0].Value = "0";
        ddlFund16.SelectedIndex = 0;
    }
    public void BindFund17()
    {
        object FundTypeId17 = ddlFundType17.SelectedValue == "" || ddlFundType17.SelectedValue == "0" ? "null" : ddlFundType17.SelectedValue;
        string sql17 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId17; ;
        clsGM.getListForBindDDL(ddlFund17, sql17, "ssi_name", "ssi_FundId");
        ddlFund17.Items.Insert(0, "All");
        ddlFund17.Items[0].Value = "0";
        ddlFund17.SelectedIndex = 0;
    }
    public void BindFund18()
    {
        object FundTypeId18 = ddlFundType18.SelectedValue == "" || ddlFundType18.SelectedValue == "0" ? "null" : ddlFundType18.SelectedValue;
        string sql18 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId18; ;
        clsGM.getListForBindDDL(ddlFund18, sql18, "ssi_name", "ssi_FundId");
        ddlFund18.Items.Insert(0, "All");
        ddlFund18.Items[0].Value = "0";
        ddlFund18.SelectedIndex = 0;
    }
    public void BindFund19()
    {
        object FundTypeId19 = ddlFundType19.SelectedValue == "" || ddlFundType19.SelectedValue == "0" ? "null" : ddlFundType19.SelectedValue;
        string sql19 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId19; ;
        clsGM.getListForBindDDL(ddlFund19, sql19, "ssi_name", "ssi_FundId");
        ddlFund19.Items.Insert(0, "All");
        ddlFund19.Items[0].Value = "0";
        ddlFund19.SelectedIndex = 0;
    }
    public void BindFund20()
    {
        object FundTypeId20 = ddlFundType20.SelectedValue == "" || ddlFundType20.SelectedValue == "0" ? "null" : ddlFundType20.SelectedValue;
        string sql20 = "SP_S_FUND_LKUP @FundTypeIdNmb=" + FundTypeId20; ;
        clsGM.getListForBindDDL(ddlFund20, sql20, "ssi_name", "ssi_FundId");
        ddlFund20.Items.Insert(0, "All");
        ddlFund20.Items[0].Value = "0";
        ddlFund20.SelectedIndex = 0;
    }

    #endregion



    public void BindFundType1()
    {
        string sql1 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType1, sql1, "Status", "ID");
    }

    public void BindFundType2()
    {
        string sql2 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType2, sql2, "Status", "ID");
    }

    public void BindFundType3()
    {
        string sql3 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType3, sql3, "Status", "ID");
    }

    public void BindFundType4()
    {
        string sql4 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType4, sql4, "Status", "ID");
    }

    public void BindFundType5()
    {
        string sql5 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType5, sql5, "Status", "ID");
    }
    public void BindFundType6()
    {
        string sql6 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType6, sql6, "Status", "ID");
    }

    public void BindFundType7()
    {
        string sql7 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType7, sql7, "Status", "ID");
    }
    public void BindFundType8()
    {
        string sql8 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType8, sql8, "Status", "ID");
    }
    public void BindFundType9()
    {
        string sql9 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType9, sql9, "Status", "ID");
    }

    public void BindFundType10()
    {
        string sql10 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType10, sql10, "Status", "ID");
    }
    public void BindFundType11()
    {
        string sql11 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType11, sql11, "Status", "ID");
    }
    public void BindFundType12()
    {
        string sql12 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType12, sql12, "Status", "ID");
    }
    public void BindFundType13()
    {
        string sql13 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType13, sql13, "Status", "ID");
    }
    public void BindFundType14()
    {
        string sql14 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType14, sql14, "Status", "ID");
    }
    public void BindFundType15()
    {
        string sql15 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType15, sql15, "Status", "ID");
    }
    public void BindFundType16()
    {
        string sql16 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType16, sql16, "Status", "ID");
    }
    public void BindFundType17()
    {
        string sql17 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType17, sql17, "Status", "ID");
    }
    public void BindFundType18()
    {
        string sql18 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType18, sql18, "Status", "ID");
    }
    public void BindFundType19()
    {
        string sql19 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType19, sql19, "Status", "ID");
    }
    public void BindFundType20()
    {
        string sql20 = "SP_S_FUNDTYPE";
        clsGM.getBindDDL(ddlFundType20, sql20, "Status", "ID");
    }
    public void BindIncSigLine()
    {
        //string sql = "SP_S_SIGNATURELINE_TEMPLATE";
        //clsGM.getBindDDL(ddlSigLine, sql, "Status", "ID");
    }

    public void BindFooterType()
    {
        //string sql = "SP_S_FOOTER_TEMPLATE";
        //clsGM.getBindDDL(ddlFooterType, sql, "Status", "ID");
    }

    private void BindValues(string strTemplate)
    {
        try
        {
            clsDB = new DB();
            string sql = "SP_S_TEMPLATE_FUND_LIST @TemplateId='" + strTemplate + "'";
            DataSet DS = clsDB.getDataSet(sql);
            DataSet NewDataSet = clsDB.getDataSet(sql);

            for (int j = 0; j < NewDataSet.Tables[0].Rows.Count; j++)
            {
                hdtblId.Value = Convert.ToString(DS.Tables[0].Rows.Count);
                hdFunds.Value = Convert.ToString(DS.Tables[0].Rows.Count);
                hdEditRows.Value = Convert.ToString(DS.Tables[0].Rows.Count);

                if (NewDataSet.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToString(NewDataSet.Tables[0].Rows[j]["FundId"]) == "" || Convert.ToString(NewDataSet.Tables[0].Rows[j]["Ssi_FundSpecificFlg"]) == "" || Convert.ToString(NewDataSet.Tables[0].Rows[j]["Ssi_FundSpecificFlg"]).ToUpper() == "False".ToUpper())
                    {
                        NewDataSet.Tables[0].Rows.RemoveAt(j);
                        NewDataSet.AcceptChanges();
                        
                        DS = NewDataSet.Copy();
                        DS.AcceptChanges();

                        hdtblId.Value = Convert.ToString(DS.Tables[0].Rows.Count);
                        hdFunds.Value = Convert.ToString(DS.Tables[0].Rows.Count);
                        hdEditRows.Value = Convert.ToString(DS.Tables[0].Rows.Count);

                        for (int k = 0; k < NewDataSet.Tables[0].Rows.Count; k++)
                        {
                            if (Convert.ToString(NewDataSet.Tables[0].Rows[k]["FundId"]) == "" && Convert.ToString(NewDataSet.Tables[0].Rows[k]["Ssi_FundSpecificFlg"]) == "" || Convert.ToString(NewDataSet.Tables[0].Rows[k]["Ssi_FundSpecificFlg"]).ToUpper() == "False".ToUpper())
                            {
                                NewDataSet.Tables[0].Rows.RemoveAt(k);
                                NewDataSet.AcceptChanges();
                               
                                DS = NewDataSet.Copy();
                                DS.AcceptChanges();

                                hdtblId.Value = Convert.ToString(DS.Tables[0].Rows.Count);
                                hdFunds.Value = Convert.ToString(DS.Tables[0].Rows.Count);
                                hdEditRows.Value = Convert.ToString(DS.Tables[0].Rows.Count);

                            }
                        }
                    }
                }
            }



            #region Bind Template values

            if (DS.Tables[1].Rows.Count > 0)
            {

                if (Convert.ToString(DS.Tables[1].Rows[0]["ssi_templateid"]) != "")
                {
                    ViewState["TemplateId"] = Convert.ToString(DS.Tables[1].Rows[0]["ssi_templateid"]);
                }

                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_AsofDate"]) != "")
                {
                    txtAsOfDate.Text = Convert.ToDateTime(DS.Tables[1].Rows[0]["Ssi_AsofDate"]).ToString("MM/dd/yyyy");
                }
                else
                {
                    txtAsOfDate.Text = "";
                }

                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_LetterText"]) != "")
                {
                    txtLetterText.Value = Convert.ToString(DS.Tables[1].Rows[0]["Ssi_LetterText"]);
                }
                else
                {
                    txtLetterText.Value = "";
                }

                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_includesignatureline"]) != "")
                {
                    string[] strSignText = Convert.ToString(DS.Tables[1].Rows[0]["Ssi_includesignatureline"]).Split(',');

                    for (int k = 0; k < strSignText.Length; k++)
                    {
                        for (int m = 0; m < LstSignText.Items.Count; m++)
                        {
                            if (LstSignText.Items[m].Value == strSignText[k])
                            {
                                LstSignText.Items[m].Selected = true;
                            }
                        }
                    }

                    //LstSignText.Text = Convert.ToString(DS.Tables[1].Rows[0]["Ssi_includesignatureline"]);
                }
                else
                {
                    LstSignText.SelectedValue = "0";
                }

                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_DateOfLetter_txt"]) != "")
                {
                    txtDateOfLetter.Text = Convert.ToDateTime(DS.Tables[1].Rows[0]["Ssi_DateOfLetter_txt"]).ToString("MM/dd/yyyy");
                }
                else
                {
                    txtDateOfLetter.Text = "";
                }

                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_TemplateName"]) != "")
                {
                    txtTemplate.Text = Convert.ToString(DS.Tables[1].Rows[0]["Ssi_TemplateName"]);
                }
                else
                {
                    txtTemplate.Text = "";
                }

                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_CustomeTemplateType"]) != "")
                {
                    ddlTemplateType.SelectedValue = Convert.ToString(DS.Tables[1].Rows[0]["Ssi_CustomeTemplateType"]);
                    ddlTemplateType.Enabled = false;

                    if (ddlTemplateType.SelectedValue == "10")
                    {
                        trFileupload.Visible = true;
                        lblFileName.Visible = true;
                        trFundSpecific.Visible = false;
                     
                    }
                    else
                    {
                        trFileupload.Visible = false;
                        lblFileName.Visible = false;
                        trFundSpecific.Visible = true;
                    }
                }
                else
                {
                    txtTemplate.Text = "";
                    ddlTemplateType.Enabled = false;
                }


                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_dynamicflg"]).ToUpper() == "True".ToUpper())
                {
                    rdoDynamic.Checked = true;
                    trOrientation.Visible = false;
                    trSignatureText.Visible = false;
                }
                else if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_dynamicflg"]).ToUpper() == "False".ToUpper())
                {
                    rdoStatic.Checked = true;
                    trOrientation.Visible = true;
                    trSignatureText.Visible = true;
                }


                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_orientation"]) != "")
                {
                    ddlOrientation.SelectedValue = Convert.ToString(DS.Tables[1].Rows[0]["Ssi_orientation"]);
                }
                else 
                {
                    ddlOrientation.SelectedValue = "0";
                }

                if (Convert.ToString(DS.Tables[1].Rows[0]["Ssi_FileName"]) != "")
                {
                    lblFileName.Text = Convert.ToString(DS.Tables[1].Rows[0]["Ssi_FileName"]);
                    lblFileName.Visible = true;
                }
                else
                {
                    lblFileName.Text = "";
                }

            }
            #endregion

            #region Bind Template Fund Values

            for (int i = 0; i < DS.Tables[0].Rows.Count; i++)
            {
                if (DS.Tables[0].Rows.Count > 0)
                {


                    //rdoYes.Checked = true;
                    if (Convert.ToString(DS.Tables[0].Rows[i]["Ssi_templatefundId"]) != "")
                    {
                        ViewState["TemplateFundId"] = Convert.ToString(DS.Tables[0].Rows[i]["Ssi_templatefundId"]);
                    }

                    if (Convert.ToString(DS.Tables[0].Rows[i]["Ssi_FundSpecificFlg"]) == "" || Convert.ToString(DS.Tables[0].Rows[i]["Ssi_FundSpecificFlg"]).ToUpper() == "False".ToUpper())
                    {
                        rdoNo.Checked = true;
                        FundSpecific = false;
                    }
                    else if (Convert.ToString(DS.Tables[0].Rows[i]["Ssi_FundSpecificFlg"]).ToUpper() == "True".ToUpper())
                    {
                        rdoYes.Checked = true;
                        FundSpecific = rdoYes.Checked;
                    }


                    if (Convert.ToString(DS.Tables[0].Rows[i]["FundId"]) != "")
                    {

                        if (i == 0)
                        {
                            num = 1;
                        }
                        else
                        {
                            num++;
                        }
                        Control trFund = FindControl("trFund" + num.ToString());
                        Control ddlFund = FindControl("ddlFund" + num.ToString());
                        Control ddlFundType = FindControl("ddlFundType" + num.ToString());
                        Control lnkRemove = FindControl("lnkRemove" + num.ToString());
                        Control txtFundDesc = FindControl("txtFundDesc" + num.ToString());


                        if (trFund != null)
                        {
                            trFund.Visible = true;
                        }

                        if (txtFundDesc != null)
                        {
                            FredCK.FCKeditorV2.FCKeditor FundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + num.ToString()));
                            //HtmlEditor FundDesc = ((HtmlEditor)FindControl("txtFundDesc" + num.ToString()));
                            if (Convert.ToString(DS.Tables[0].Rows[i]["ssi_fundtxt"]) != "")
                            {
                                FundDesc.Value = Convert.ToString(DS.Tables[0].Rows[i]["ssi_fundtxt"]);
                            }
                            else
                            {
                                FundDesc.Value = "";
                            }
                        }

                        if (ddlFundType != null)
                        {
                            DropDownList drpFundType = ((DropDownList)FindControl("ddlFundType" + num.ToString()));
                            if (Convert.ToString(DS.Tables[0].Rows[i]["FundType"]) != "")
                            {
                                drpFundType.SelectedValue = Convert.ToString(DS.Tables[0].Rows[i]["FundType"]);
                            }
                            else
                            {
                                drpFundType.SelectedValue = "0";
                            }
                        }

                        if (ddlFund != null)
                        {
                            DropDownList drpFund = ((DropDownList)FindControl("ddlFund" + num.ToString()));
                            if (Convert.ToString(DS.Tables[0].Rows[i]["FundId"]) != "")
                            {
                                drpFund.SelectedValue = Convert.ToString(DS.Tables[0].Rows[i]["FundId"]);
                            }
                            else
                            {
                                drpFund.SelectedValue = "0";
                            }
                        }



                        if (lnkRemove != null)
                        {
                            LinkButton aRemove = ((LinkButton)FindControl("lnkRemove" + num.ToString()));
                            if (Convert.ToString(DS.Tables[0].Rows[i]["Ssi_templatefundId"]) != "")
                            {
                                aRemove.CommandArgument = Convert.ToString(DS.Tables[0].Rows[i]["Ssi_templatefundId"]);
                            }
                            else
                            {
                                aRemove.CommandArgument = "";
                            }
                        }
                    }
                    else
                    {
                        num = 1;
                    }
                    ShowHideBindValues(DS, num);
                }

            }


            if (DS.Tables[0].Rows.Count == 0)
            {
                ShowHideBindValues(DS, 1);
                if (ViewState["TemplateFundId"] != "")
                {
                    ViewState["TemplateFundId"] = "";
                }
            }



            #endregion
        }
        catch (Exception ex)
        {
 
        }
        
    }
    

    protected void btnSave_Click(object sender, EventArgs e)
    {
        int intResult = 0;

        #region validate save

        System.Text.StringBuilder sb1 = new System.Text.StringBuilder();
        Type tp1 = this.GetType();

        if (ddlTemplateType.SelectedValue != "" || ddlTemplateType.SelectedValue != "0")
        {

            if (ddlTemplateType.SelectedValue == "10")
            {
                if (FileUpload1.HasFile == false)
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please select file to upload.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }

                if (System.IO.Path.GetExtension(FileUpload1.FileName) != ".pdf")
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please upload .pdf files only.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }
            }

            if (ddlTemplateType.SelectedValue == "1" || ddlTemplateType.SelectedValue == "2" || ddlTemplateType.SelectedValue =="10")
            {
                if (txtAsOfDate.Text == "")
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please enter asofdate.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }

                if (txtDateOfLetter.Text == "")
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please enter letter date.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }



                //  for (int j = 1; j < 11; j++)
                //for (int j = 1; j < 12; j++)
                // for (int j = 1; j < 16; j++) // added 11_21_2018
                for (int j = 1; j < 21; j++) // added 11_23_2018
                {

                    Control trFund = FindControl("trFund" + j.ToString());
                    if (trFund.Visible == true)
                    {
                        FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + Convert.ToString(j)));
                        //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                        //txtFundDesc.Text = "";
                        if (txtFundDesc.Value == "")
                        {
                            sb1.Append("\n<script type=text/javascript>\n");
                            sb1.Append("\n alert('Please enter fund details.');");
                            sb1.Append("</script>");
                            ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                            return;
                        }
                    }

                    
                }
            }

            if (ddlTemplateType.SelectedValue == "3" || ddlTemplateType.SelectedValue == "4" || ddlTemplateType.SelectedValue == "5" || ddlTemplateType.SelectedValue == "6" || ddlTemplateType.SelectedValue == "7" || ddlTemplateType.SelectedValue == "8" || ddlTemplateType.SelectedValue == "9")
            {
                if (txtAsOfDate.Text == "")
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please enter asofdate.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }


                if (txtLetterText.Value == "")
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please enter letter text.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }
            }


            if (rdoStatic.Checked == true)
            {
                if (ddlOrientation.SelectedValue == "" || ddlOrientation.SelectedValue == "0")
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please select orientation.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }

                if (LstSignText.SelectedValue == "" || LstSignText.SelectedValue == "0")
                {
                    sb1.Append("\n<script type=text/javascript>\n");
                    sb1.Append("\n alert('Please select signature text.');");
                    sb1.Append("</script>");
                    ClientScript.RegisterStartupScript(tp1, "Script", sb1.ToString());
                    return;
                }
            }

        }

       

        if (rdoYes.Checked == true)
        {
            //for (int i = 1; i < 11; i++)
            // for (int i = 1; i < 12; i++)
          //  for (int i = 1; i < 16; i++) // added 11_21_2018
                for (int i = 1; i < 21; i++) // added 11_23_2018
                {

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();

                Control trFund = FindControl("trFund" + i.ToString());
                Control ddlFundType = FindControl("ddlFundType" + i.ToString());
                Control ddlFund = FindControl("ddlFund" + i.ToString());
                Control ddlFund1 = FindControl("ddlFund" + i.ToString());
                Control htmlEditor = FindControl("txtFundDesc" + i.ToString());
                DropDownList Fund = null;
                DropDownList Fund1 = null;
                if (trFund != null)
                {
                    if (trFund.Visible == true)
                    {
                        if (ddlFund != null)
                        {
                            Fund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                            //Fund.SelectedValue = "0";
                            if (Fund.SelectedValue == "" || Fund.SelectedValue == "0")
                            {
                                sb.Append("\n<script type=text/javascript>\n");
                                sb.Append("\n alert('Please select fund.');");
                                sb.Append("</script>");
                                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                                return;
                            }
                        }

                        //if (ddlFund != null)
                        //{
                        //    if (Fund.SelectedValue == Fund.SelectedValue)
                        //    {
                        //        //Fund.Focus();
                        //        sb.Append("\n<script type=text/javascript>\n");
                        //        sb.Append("\n alert('Please select some other fund from fund.');");
                        //        sb.Append("</script>");
                        //        ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                        //        return;
                        //    }
                        //}

                        if (ddlFundType != null)
                        {
                            DropDownList FundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
                            // FundType.SelectedValue = "0";
                            if (FundType.SelectedValue != "" || FundType.SelectedValue != "0")
                            {
                                if (Fund.SelectedValue == "" || Fund.SelectedValue == "0")
                                {
                                    sb.Append("\n<script type=text/javascript>\n");
                                    sb.Append("\n alert('Please select fund.');");
                                    sb.Append("</script>");
                                    ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                                    return;
                                }
                            }
                        }

                        FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + Convert.ToString(i)));
                        //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                        //txtFundDesc.Text = "";
                        if (txtFundDesc.Value == "")
                        {
                            sb.Append("\n<script type=text/javascript>\n");
                            sb.Append("\n alert('Please enter fund details.');");
                            sb.Append("</script>");
                            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                            return;
                        }
                    }
                }
            }
        }
        #endregion
        if (ddlTemplate.SelectedValue == "0")
        {
            Save();
        }
        else
        {
            if (ViewState["TemplateId"] != null || ViewState["TemplateFundId"] != null)
            {
                if (ViewState["TemplateFundId"] == null)
                {
                    ViewState["TemplateFundId"] = "";
                }
                EditSave(ViewState["TemplateId"].ToString(), ViewState["TemplateFundId"].ToString());
            }
        }

      

    }



    private void Save()
    {
        
        int intResult = 0;
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;

        lblError.Text = "";
        DataSet loInvoiceData = null;
        string UserId = GetcurrentUser(false);
        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service =clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
       // catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
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

        //ssi_template objTemplate = null;
        //ssi_templatefund objTemplateFund = null;
        Entity objTemplate = new Entity("ssi_template");
        Entity objTemplateFund = new Entity("ssi_templatefund");
        try
        {
            if (ddlTemplate.SelectedValue == "0")
            {

                #region Template
               // objTemplate = new ssi_template();
                
                if (txtAsOfDate.Text != "")
                {
                    DateTime AsofDate = Convert.ToDateTime(txtAsOfDate.Text);
                    strAsOfDate = AsofDate.ToString("yyyy-MMdd"); ;
                }

                string UserName = GetcurrentUser(true);

                if (txtTemplate.Text != "")
                {
                    txtTemplate.Text = txtTemplate.Text;
                }
                else
                {
                    if (ddlTemplateType.SelectedValue == "1")
                    {
                        txtTemplate.Text = ddlTemplateType.SelectedItem.Text;
                    }
                    else if (txtTemplate.Text == "")
                    {
                        if (ddlTemplateType.SelectedValue != "1")
                        {
                            txtTemplate.Text = ddlTemplateType.SelectedItem.Text;
                        }
                    }
                }

                TemplateName = strAsOfDate + "-" + txtTemplate.Text + "-" + UserName;
                //objTemplate.ssi_name = TemplateName;//user name
                //objTemplate.ssi_templatename = txtTemplate.Text;
                objTemplate["ssi_name"] = TemplateName;//user name
                objTemplate["ssi_templatename"] = txtTemplate.Text;

                if (txtAsOfDate.Text != "")
                {
                    //objTemplate.ssi_asofdate = new CrmDateTime();
                    //objTemplate.ssi_asofdate.Value = txtAsOfDate.Text;
                    objTemplate["ssi_asofdate"] = Convert.ToDateTime(txtAsOfDate.Text);

                }

                if (txtDateOfLetter.Text != "")
                {
                    //objTemplate.ssi_dateofletter_txt = new CrmDateTime();
                    //objTemplate.ssi_dateofletter_txt.Value = txtDateOfLetter.Text;
                    objTemplate["ssi_dateofletter_txt"] = Convert.ToDateTime(txtDateOfLetter.Text);
                }

                if (LstSignText.SelectedValue != "" && LstSignText.Text !="0")
                {
                    string txtSignature = clsGM.GetMultipleSelectedItemsFromListBox(LstSignText);
                   // objTemplate.ssi_includesignatureline = txtSignature;
                    objTemplate["ssi_includesignatureline"] = txtSignature;
                }

                if (txtLetterText.Value != "")
                {
                   // objTemplate.ssi_lettertext = txtLetterText.Value;
                    objTemplate["ssi_lettertext"] = txtLetterText.Value;
                }

                //objTemplate.ssi_custometemplatetype = new Picklist();
                //objTemplate.ssi_custometemplatetype.Value = Convert.ToInt32(ddlTemplateType.SelectedValue);
                objTemplate["ssi_custometemplatetype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ddlTemplateType.SelectedValue));



                //objTemplate.ssi_customflg = new CrmBoolean();
                //objTemplate.ssi_customflg.Value = true;
                objTemplate["ssi_customflg"] = true;


                if (rdoDynamic.Checked == true)
                {
                    //objTemplate.ssi_dynamicflg = new CrmBoolean();
                    //objTemplate.ssi_dynamicflg.Value = true;
                    objTemplate["ssi_dynamicflg"] = true;

                    ddlOrientation.SelectedValue = "0";
                    LstSignText.SelectedValue = "0";
                }
                else if (rdoStatic.Checked == false)
                {
                    //objTemplate.ssi_dynamicflg = new CrmBoolean();
                    //objTemplate.ssi_dynamicflg.Value = false;
                    objTemplate["ssi_dynamicflg"] = false;
                }

                if (ddlOrientation.SelectedValue != "" && ddlOrientation.SelectedValue != "0")
                {
                    //objTemplate.ssi_orientation = new Picklist();
                    //objTemplate.ssi_orientation.Value = Convert.ToInt32(ddlOrientation.SelectedValue);
                    objTemplate["ssi_orientation"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ddlOrientation.SelectedValue));

                }

                if (ddlTemplateType.SelectedValue == "10")//File Upload =10 
                {
                    if (System.IO.Path.GetExtension(FileUpload1.FileName) == ".pdf")
                    {
                        FileUpload1.PostedFile.SaveAs(AppLogic.GetParam(AppLogic.ConfigParam.FileUploadUrl) + FileUpload1.FileName);
                        //objTemplate.ssi_filename = FileUpload1.FileName;
                        objTemplate["ssi_filename"] = FileUpload1.FileName;
                        lblFileName.Text = FileUpload1.FileName;
                        lblFileName.Visible = true;
                    }
                }


                service.Create(objTemplate);
                intResult++;
                if (intResult > 0)
                {
                    clsDB = new DB();
                    string sql = "select top 1 ssi_templateid from ssi_template Where ssi_name= '" + TemplateName + "' order by ssi_templateid desc";
                    DataSet loDataset = clsDB.getDataSet(sql);
                    if (Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_templateid"]) != "")
                    {
                        BindTemplate();
                        ddlTemplate.SelectedValue = Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_templateid"]);
                    }
                }
                #endregion

                #region Template Fund

               // objTemplateFund = new ssi_templatefund();
                
                //objTemplateFund.ssi_asofdate = new CrmDateTime();
                //objTemplateFund.ssi_asofdate.Value = txtAsOfDate.Text;
                objTemplateFund["ssi_asofdate"] = Convert.ToDateTime(txtAsOfDate.Text);

                if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                {
                    //objTemplateFund.ssi_templateid = new Lookup();
                    //objTemplateFund.ssi_templateid.Value = new Guid(ddlTemplate.SelectedValue);
                    objTemplateFund["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(ddlTemplate.SelectedValue));
                }

                if (rdoYes.Checked == true)
                {
                    //objTemplateFund.ssi_fundspecificflg = new CrmBoolean();
                    //objTemplateFund.ssi_fundspecificflg.Value =true;
                    objTemplateFund["ssi_fundspecificflg"] = true;

                }
                else if (rdoNo.Checked == true)
                {
                    //objTemplateFund.ssi_fundspecificflg = new CrmBoolean();
                    //objTemplateFund.ssi_fundspecificflg.Value = false;
                    objTemplateFund["ssi_fundspecificflg"] = false;
                }


                if (rdoYes.Checked == true)
                {
                    // for (int i = 1; i < 11; i++)
                    //for (int i = 1; i < 12; i++)
                    //  for (int i = 1; i < 16; i++)// added 11_21_2018
                    for (int i = 1; i < 21; i++)// added 11_23_2018
                    {
                        Control trFund = FindControl("trFund" + i.ToString());
                        Control ddlFundType = FindControl("ddlFundType" + i.ToString());
                        Control ddlFund = FindControl("ddlFund" + i.ToString());
                        Control htmlEditor = FindControl("txtFundDesc" + i.ToString());
                        DropDownList Fund = null;

                        if (trFund != null)
                        {
                            if (trFund.Visible == true)
                            {
                                if (ddlFund != null)
                                {
                                    Fund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                                    if (Fund.SelectedValue != "" && Fund.SelectedValue != "0")
                                    {
                                        //objTemplateFund.ssi_fundidid = new Lookup();
                                        //objTemplateFund.ssi_fundidid.Value = new Guid(Fund.SelectedValue);
                                        objTemplateFund["ssi_fundidid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Fund.SelectedValue));
                                    }
                                }

                                if (htmlEditor != null)
                                {
                                    FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
                                    //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                                    if (txtFundDesc.Value != "")
                                    {
                                       // objTemplateFund.ssi_fundtxt = txtFundDesc.Value;
                                         objTemplateFund["ssi_fundtxt"] = txtFundDesc.Value;
                                    }
                                }

                                service.Create(objTemplateFund);
                            }
                        }
                    }
                    
                }
                else if (rdoNo.Checked == true)
                {
                    service.Create(objTemplateFund);
                    intResult++;
                }

               #endregion

                if (intResult > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(TemplateId);
                    }
                    lblError.Visible = true;
                    lblError.Text = "Saved Successfully.";
                   
                    
                }
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
    }


    private void EditSave(string TemplateId, string TemplateFundId)
    {
        int intResult = 0;
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;
        lblError.Text = "";
        DataSet loInvoiceData = null;
        string UserId = GetcurrentUser(false);
        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
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

        //ssi_template objTemplate = null;
        //ssi_templatefund objTemplateFund = null;
        //ssi_templatefund objTempFundCreate = null;
        Entity objTemplate = new Entity("ssi_template");
        Entity objTemplateFund = new Entity("ssi_templatefund");
        Entity objTempFundCreate = new Entity("ssi_templatefund");
        try
        {
            //if ((ddlTemplate.SelectedValue != "0") && txtTemplate.Text != "")
            //{

                #region Template
                //objTemplate = new ssi_template();
                //objTemplate.ssi_templateid = new Key();
                //objTemplate.ssi_templateid.Value = new Guid(TemplateId);
                objTemplate["ssi_templateid"] = new Guid(Convert.ToString(TemplateId));

                //objTemplate.ssi_name = txtTemplate.Text;
                if (txtAsOfDate.Text != "")
                {
                    DateTime AsofDate = Convert.ToDateTime(txtAsOfDate.Text);
                    strAsOfDate = AsofDate.ToString("yyyy-MMdd");
                }
            
                string UserName = GetcurrentUser(true);
                if (txtTemplate.Text != "")
                {
                    TemplateName = strAsOfDate + "-" + txtTemplate.Text + "-" + UserName;
                    //objTemplate.ssi_name = TemplateName;//user name
                    objTemplate["ssi_name"] = TemplateName;//user name
                }
                             
                if (txtAsOfDate.Text != "")
                {
                    //objTemplate.ssi_asofdate = new CrmDateTime();
                    //objTemplate.ssi_asofdate.Value = txtAsOfDate.Text;
                    objTemplate["ssi_asofdate"] = Convert.ToDateTime(txtAsOfDate.Text);

                }

                if (txtDateOfLetter.Text != "")
                {
                    //objTemplate.ssi_dateofletter_txt = new CrmDateTime();
                    //objTemplate.ssi_dateofletter_txt.Value = txtDateOfLetter.Text;
                    objTemplate["ssi_dateofletter_txt"] = Convert.ToDateTime(txtDateOfLetter.Text);
                }

                if (LstSignText.SelectedValue != "" && LstSignText.SelectedValue !="0")
                {
                    string txtSignature = clsGM.GetMultipleSelectedItemsFromListBox(LstSignText);
                   // objTemplate.ssi_includesignatureline = txtSignature;
                    objTemplate["ssi_includesignatureline"] = txtSignature;
                }

                if (txtLetterText.Value != "")
                {
                    //objTemplate.ssi_lettertext = txtLetterText.Value;
                    objTemplate["ssi_lettertext"] = txtLetterText.Value;
                }

                if (rdoDynamic.Checked == true)
                {
                    //objTemplate.ssi_dynamicflg = new CrmBoolean();
                    //objTemplate.ssi_dynamicflg.Value = true;
                    objTemplate["ssi_dynamicflg"] = true;

                    //objTemplate.ssi_orientation = new Picklist();
                    //objTemplate.ssi_orientation.IsNull = true;
                    //objTemplate.ssi_orientation.IsNullSpecified = true;
                    objTemplate["ssi_orientation"] = null;
                    

                   // objTemplate.ssi_includesignatureline = "";
                    objTemplate["ssi_includesignatureline"] = "";

                    //LstSignText.SelectedValue = "0";
                }
                else if (rdoStatic.Checked == true)
                {
                    //objTemplate.ssi_dynamicflg = new CrmBoolean();
                    //objTemplate.ssi_dynamicflg.Value = false;
                    objTemplate["ssi_dynamicflg"]= false;
                }

                if (ddlOrientation.SelectedValue != "" && ddlOrientation.SelectedValue != "0")
                {
                    //objTemplate.ssi_orientation = new Picklist();
                    //objTemplate.ssi_orientation.Value = Convert.ToInt32(ddlOrientation.SelectedValue);
                    objTemplate["ssi_orientation"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ddlOrientation.SelectedValue));

                }



                if (ddlTemplateType.SelectedValue == "10")//File Upload =10 
                {
                    if (System.IO.Path.GetExtension(FileUpload1.FileName) == ".pdf")
                    {
                        FileUpload1.PostedFile.SaveAs(AppLogic.GetParam(AppLogic.ConfigParam.FileUploadUrl) + FileUpload1.FileName);
                        //objTemplate.ssi_filename = FileUpload1.FileName;
                        objTemplate["ssi_filename"] = FileUpload1.FileName;
                    }
                }
                service.Update(objTemplate);
                intResult++;
                #endregion

                #region Edit Template Fund
                           
                //Update Existing Template Fund
                if (TemplateFundId != "")
                {
                    
                    if (rdoYes.Checked == true)
                    {
                        for (int i = 0; i < Convert.ToInt32(hdEditRows.Value); i++)
                        {
                            //objTemplateFund = new ssi_templatefund();

                            //objTemplateFund.ssi_templatefundid = new Key();
                            //objTemplateFund.ssi_templatefundid.Value = new Guid(TemplateFundId);
                            objTemplateFund["ssi_templatefundid"] = new Guid(TemplateFundId);

                            //objTemplateFund.ssi_asofdate = new CrmDateTime();
                            //objTemplateFund.ssi_asofdate.Value = txtAsOfDate.Text;
                            objTemplateFund["ssi_asofdate"] = Convert.ToDateTime(txtAsOfDate.Text);

                            if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                            {
                                //objTemplateFund.ssi_templateid = new Lookup();
                                //objTemplateFund.ssi_templateid.Value = new Guid(ddlTemplate.SelectedValue);
                                objTemplateFund["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(ddlTemplate.SelectedValue));
                            }


                            if (rdoYes.Checked == true)
                            {
                                //objTemplateFund.ssi_fundspecificflg = new CrmBoolean();
                                //objTemplateFund.ssi_fundspecificflg.Value = true;
                                objTemplateFund["ssi_fundspecificflg"] = true;

                            }
                            else if (rdoNo.Checked == true)
                            {
                                //objTemplateFund.ssi_fundspecificflg = new CrmBoolean();
                                //objTemplateFund.ssi_fundspecificflg.Value = false;
                                objTemplateFund["ssi_fundspecificflg"] = false;
                            }
                            if (i == 0)
                            {
                                num = 1;
                            }
                            else
                            {
                                num++;
                            }

                            Control trFund = FindControl("trFund" + num.ToString());
                            Control ddlFundType = FindControl("ddlFundType" + num.ToString());
                            Control ddlFund = FindControl("ddlFund" + num.ToString());
                            Control htmlEditor = FindControl("txtFundDesc" + num.ToString());
                            DropDownList Fund = null;

                            if (trFund != null)
                            {
                                if (trFund.Visible == true)
                                {
                                    if (ddlFund != null)
                                    {
                                        Fund = ((DropDownList)FindControl("ddlFund" + num.ToString()));
                                        if (Fund.SelectedValue != "" && Fund.SelectedValue != "0")
                                        {
                                            //objTemplateFund.ssi_fundidid = new Lookup();
                                            //objTemplateFund.ssi_fundidid.Value = new Guid(Fund.SelectedValue);
                                            objTemplateFund["ssi_fundidid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Fund.SelectedValue));


                                        }
                                    }

                                    if (htmlEditor != null)
                                    {
                                        FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + num.ToString()));
                                        //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + num.ToString()));
                                        if (txtFundDesc.Value != "")
                                        {
                                            //objTemplateFund.ssi_fundtxt = txtFundDesc.Value;
                                            objTemplateFund["ssi_fundtxt"] = txtFundDesc.Value;
                                        }
                                    }

                                    service.Update(objTemplateFund);
                                    intResult++;
                                }
                            }
                        }

                    }
                    else if (rdoNo.Checked == true)
                    {
                        if (TemplateFundId != "")
                        {
                          //  objTemplateFund = new ssi_templatefund();

                            //objTemplateFund.ssi_templatefundid = new Key();
                            //objTemplateFund.ssi_templatefundid.Value = new Guid(TemplateFundId);
                            objTemplateFund["ssi_templatefundid"] = new Guid(Convert.ToString(TemplateFundId));


                            //objTemplateFund.ssi_asofdate = new CrmDateTime();
                            //objTemplateFund.ssi_asofdate.Value = txtAsOfDate.Text;
                            objTemplateFund["ssi_asofdate"] = Convert.ToDateTime(txtAsOfDate.Text);

                            if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                            {
                                //objTemplateFund.ssi_templateid = new Lookup();
                                //objTemplateFund.ssi_templateid.Value = new Guid(ddlTemplate.SelectedValue);
                                objTemplateFund["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(ddlTemplate.SelectedValue));
                            }


                            if (rdoYes.Checked == true)
                            {
                               // objTemplateFund.ssi_fundspecificflg = new CrmBoolean();
                              //  objTemplateFund.ssi_fundspecificflg.Value = true;
                                  objTemplateFund["ssi_fundspecificflg"] = true;
                            }
                            else if (rdoNo.Checked == true)
                            {
                               // objTemplateFund.ssi_fundspecificflg = new CrmBoolean();
                                //objTemplateFund.ssi_fundspecificflg.Value = false;
                                 objTemplateFund["ssi_fundspecificflg"] = false;
                            }
                            service.Update(objTemplateFund);
                            intResult++;
                        }
                        
                    }
                }


                if (TemplateId != "")
                {
                    
                    if (rdoYes.Checked == true)
                    {
                        int val = Convert.ToInt32(hdFunds.Value) + 1;
                    // for (int i = val; i < 11; i++)
                    //for (int i = val; i < 12; i++)
                    // for (int i = 1; i < 16; i++)// added 11_21_2018
                    for (int i = 1; i < 21; i++)// added 11_23_2018
                    {
                          //  objTempFundCreate = new ssi_templatefund();

                            //objTempFundCreate.ssi_asofdate = new CrmDateTime();
                            //objTempFundCreate.ssi_asofdate.Value = txtAsOfDate.Text;
                            objTempFundCreate["ssi_asofdate"] = Convert.ToDateTime(txtAsOfDate.Text);

                            if (TemplateId != "" && TemplateId != "0")
                            {
                                //objTempFundCreate.ssi_templateid = new Lookup();
                                //objTempFundCreate.ssi_templateid.Value = new Guid(TemplateId);
                                objTempFundCreate["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(TemplateId));
                            }

                            if (rdoYes.Checked == true)
                            {
                                //objTempFundCreate.ssi_fundspecificflg = new CrmBoolean();
                                //objTempFundCreate.ssi_fundspecificflg.Value = true;
                                objTempFundCreate["ssi_fundspecificflg"] = true;
                            }
                            else if (rdoNo.Checked == true)
                            {
                                //objTempFundCreate.ssi_fundspecificflg = new CrmBoolean();
                                //objTempFundCreate.ssi_fundspecificflg.Value = false;
                                 objTempFundCreate["ssi_fundspecificflg"] = false;
                            }
                            Control trFund = FindControl("trFund" + i.ToString());
                            Control ddlFundType = FindControl("ddlFundType" + i.ToString());
                            Control ddlFund = FindControl("ddlFund" + i.ToString());
                            Control htmlEditor = FindControl("txtFundDesc" + i.ToString());
                            DropDownList Fund = null;

                            if (trFund != null)
                            {
                                if (trFund.Visible == true)
                                {
                                    if (ddlFund != null)
                                    {
                                        Fund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                                        if (Fund.SelectedValue != "" && Fund.SelectedValue != "0")
                                        {
                                            //objTempFundCreate.ssi_fundidid = new Lookup();
                                            //objTempFundCreate.ssi_fundidid.Value = new Guid(Fund.SelectedValue);
                                            objTempFundCreate["ssi_fundidid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Fund.SelectedValue));
                                        }
                                    }

                                    FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
                                    //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                                    if (txtFundDesc.Value != "")
                                    {
                                       // objTempFundCreate.ssi_fundtxt = txtFundDesc.Value;
                                        objTempFundCreate["ssi_fundtxt"] = txtFundDesc.Value;
                                    }

                                    service.Create(objTempFundCreate);
                                    intResult++;
                                }
                            }
                        }

                    }
                    else if (rdoNo.Checked == true)
                    {
                        //objTempFundCreate = new ssi_templatefund();

                        //objTempFundCreate.ssi_asofdate = new CrmDateTime();
                        //objTempFundCreate.ssi_asofdate.Value = txtAsOfDate.Text;
                        objTempFundCreate["ssi_asofdate"] = Convert.ToDateTime(txtAsOfDate.Text);

                        if (TemplateId != "" && TemplateId != "0")
                        {
                            //objTempFundCreate.ssi_templateid = new Lookup();
                            //objTempFundCreate.ssi_templateid.Value = new Guid(TemplateId);
                            objTempFundCreate["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(TemplateId));
                        }

                        if (rdoYes.Checked == true)
                        {
                            //objTempFundCreate.ssi_fundspecificflg = new CrmBoolean();
                            //objTempFundCreate.ssi_fundspecificflg.Value = true;
                            objTempFundCreate["ssi_fundspecificflg"] = true;
                        }
                        else if (rdoNo.Checked == true)
                        {
                            //objTempFundCreate.ssi_fundspecificflg = new CrmBoolean();
                            //objTempFundCreate.ssi_fundspecificflg.Value = false;
                            objTempFundCreate["ssi_fundspecificflg"] = false;
                        }

                        service.Create(objTempFundCreate);
                        intResult++;
                    }
                }
                          
    

                // Remaining  To add fund 


                

                #endregion

            //}

                if (intResult > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(TemplateId);
                    }
                    lblError.Visible = true;
                    lblError.Text = "Updated Successfully.";
                   

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

    }
    #region OLDCODE CRM2016UPGRADE
    ///// <summary>
    ///// Set up the CRM Service.
    ///// </summary>
    ///// <param name="organizationName">My Organization</param>
    ///// <returns>CrmService configured with AD Authentication</returns>
    //public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    //string username = WindowsIdentity.GetCurrent().Name;

    //    //if (username == "CORP\\gbhagia")
    //    //{
    //    //    // Use the global user ID of the system user that is to be impersonated.
    //    //    token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
    //    //    //token.CallerId = new Guid("C42C7E05-8303-DE11-A38C-001D09665E8F");//gary                
    //    //}
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

    //    return service;
    //}
    #endregion

    private DataSet AddTotals(DataSet lodataset)
    {
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            DataRow dr = lodataset.Tables[0].NewRow();

            for (int j = 0; j < lodataset.Tables[0].Rows.Count; j++)
            {
                try
                {
                    for (int k = 0; k < lodataset.Tables[0].Columns.Count; k++)
                    {

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Current Portfolio %"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Current Portfolio Value"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }


                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Suggested Allocation"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0.0M;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }

                        //if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Tactical Tilt"))
                        //{
                        //    if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                        //        dr[lodataset.Tables[0].Columns[k].ColumnName] = 0;

                        //    dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        //}

                        if (lodataset.Tables[0].Columns[k].ColumnName.Contains("Tactical Target"))
                        {
                            if (Convert.ToString(dr[lodataset.Tables[0].Columns[k].ColumnName]) == "")
                                dr[lodataset.Tables[0].Columns[k].ColumnName] = 0;

                            dr[lodataset.Tables[0].Columns[k].ColumnName] = Convert.ToDecimal(dr[lodataset.Tables[0].Columns[k].ColumnName]) + Convert.ToDecimal(lodataset.Tables[0].Rows[j][k]);
                        }


                    }
                }
                catch
                {

                }

            }



            dr["_LineFlg"] = 2;
            dr["Asset Class"] = "TOTAL";
            lodataset.Tables[0].Rows.Add(dr);
            lodataset.AcceptChanges();


        }
        return lodataset;
    }
    public void SetBorder(Cell foCell, bool IsTop, bool IsBottom, bool IsLeft, bool IsRight)
    {
        if (IsTop == true)
        {
            foCell.BorderWidthTop = 1F;
            foCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsBottom == true)
        {
            foCell.BorderWidthBottom = 1F;
            foCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsLeft == true)
        {
            foCell.BorderWidthLeft = 1F;
            foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        if (IsRight == true)
        {
            foCell.BorderWidthRight = 1F;
            foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        }
        //if (TopBottom == true)
        //{
        //    foCell.BorderWidthBottom = 1F;
        //    foCell.BorderColorBottom = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthLeft = 0.2F;
        //    foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthRight = 0.2F;
        //    foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        //}
        //else
        //{
        //    foCell.BorderWidthTop= 1F;
        //    foCell.BorderColorTop = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthLeft = 0.2F;
        //    foCell.BorderColorLeft = new iTextSharp.text.Color(System.Drawing.Color.Black);

        //    foCell.BorderWidthRight = 0.2F;
        //    foCell.BorderColorRight = new iTextSharp.text.Color(System.Drawing.Color.Black);
        //}

    }
    


    protected void ddlTemplate_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
        {
            for (int m = 0; m < LstSignText.Items.Count; m++)
            {
                LstSignText.Items[m].Selected = false;
            }
            BindValues(ddlTemplate.SelectedValue);
                       
        
            trFundSpecific.Visible = false;
            trDynamic.Visible = false;
           

            //txtTemplate.Text = "";
            lblError.Text = "";
        }
        else if (ddlTemplate.SelectedValue == "0")
        {
            lblError.Text = "";
            txtTemplate.Text = "";
            txtAsOfDate.Text = "";
            txtDateOfLetter.Text="";
            txtLetterText.Value = "";
            for (int m = 0; m < LstSignText.Items.Count; m++)
            {
                LstSignText.Items[m].Selected = false;
            }
            rdoNo.Checked = false;
            trFundSpecific.Visible = true;
            ddlTemplateType.Enabled = true;
            ddlTemplateType.SelectedValue = "1";

            if (rdoNo.Checked == false)
            {
                rdoYes.Checked = true;
            }


            //for (int i = 1; i < 11; i++)
            // for (int i = 1; i < 12; i++)
            // for (int i = 1; i < 16; i++)// added 11_21_2018
            for (int i = 1; i < 21; i++)// added 11_23_2018
            {
                Control trFund = FindControl("trFund" + i.ToString());
                Control ddlFund = FindControl("ddlFund" + i.ToString());
                Control ddlFundType = FindControl("ddlFundType" + i.ToString());

                if (i > 1)
                {
                    if (ddlFund != null)
                    {
                        DropDownList drpFund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                        drpFund.SelectedValue = "0";
                    }

                    if (ddlFundType != null)
                    {
                        DropDownList drpFundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
                        drpFundType.SelectedValue = "0";
                    }

                    //if (htmlEditor != null)
                    //{
                    FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
                        //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                        txtFundDesc.Value = "";
                    //}

                    if (trFund != null)
                    {
                        trFund.Visible = false;
                    }
                }
                else
                {
                    if (ddlFund != null)
                    {
                        DropDownList drpFund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                        drpFund.SelectedValue = "0";
                    }

                    if (ddlFundType != null)
                    {
                        DropDownList drpFundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
                        drpFundType.SelectedValue = "0";
                    }

                    FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
                    //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                    txtFundDesc.Value = "";
               
                    if (trFund != null)
                    {
                        trFund.Visible = true;
                    }
                }
                
            }
            
            trAddAnother.Visible = true;
            hdtblId.Value = "1";

            trFileupload.Visible = false;
            lblFileName.Text = "";

        }
    }


    private string GetcurrentUser(bool GetName)
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = p.Identity.Name;//Request.LogonUserIdentity.Name;// 
        //Changed Windows to - ADFS Claims Login 8_9_2019
        IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        string strName = claimsIdentity.Name;

        //Response.Write("p.Identity.Name:" + strName + "<br/><br/>");
        //strName = HttpContext.Current.User.Identity.Name.ToString();
        //Response.Write("HttpContext.Current.User.Identity.Name:" + strName + "<br/><br/>");
        //strName = Request.ServerVariables["AUTH_USER"]; //Finding with name
        //Response.Write("AUTH_USER:" + strName + "<br/><br/>");
        //////////
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        sqlstr = "select top 1 internalemailaddress,systemuserid,left(ltrim(firstname),1) + '' + LastName as customname from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        //Response.Write(strName + "<br/><br/>");
        //Response.Write(Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]));
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            if (GetName == true)
                return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["customname"]);
            else
                return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
        }
        else
        {
            return UserID = "";
        }
    }
    protected void ddlFundType1_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund1();
    }
    protected void ddlFundType2_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund2();
        trFund2.Visible=true;//.Style.Add("display", "inline");
    }
    protected void ddlFundType3_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund3();
        trFund3.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType4_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund4();
        trFund4.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType5_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund5();
        trFund5.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType6_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund6();
        trFund6.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType7_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund7();
        trFund7.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType8_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund8();
        trFund8.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType9_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund9();
        trFund9.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType10_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund10();
        trFund10.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType11_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund11();
        trFund11.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType12_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund12();
        trFund12.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType13_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund13();
        trFund13.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType14_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund14();
        trFund14.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType15_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund15();
        trFund15.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType16_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund16();
        trFund16.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType17_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund17();
        trFund17.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType18_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund18();
        trFund18.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType19_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund19();
        trFund19.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }
    protected void ddlFundType20_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindFund20();
        trFund20.Visible = true;//.Style.Add("display", "inline");.Style.Add("display", "inline");
    }

    protected void lnkRemove1_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
                //i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
                lblError.Text = "Error in remove : " + ex.Message.ToString();
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund1.Visible = false;
            ddlFundType1.ClearSelection();
            ddlFund1.ClearSelection();
            txtFundDesc1.Value = "";
            hdtblId.Value = "0";
        }
    }
    protected void lnkRemove2_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
               // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund2.Visible = false;
            ddlFundType2.ClearSelection();
            ddlFund2.ClearSelection();
            txtFundDesc2.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove3_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
               // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund3.Visible = false;
            ddlFundType3.ClearSelection();
            ddlFund3.ClearSelection();
            txtFundDesc3.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove4_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
                //i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund4.Visible = false;
            ddlFundType4.ClearSelection();
            ddlFund4.ClearSelection();
            txtFundDesc4.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove5_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
              //  i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund5.Visible = false;
            ddlFundType5.ClearSelection();
            ddlFund5.ClearSelection();
            txtFundDesc5.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove6_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
                //i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund6.Visible = false;
            ddlFundType6.ClearSelection();
            ddlFund6.ClearSelection();
            txtFundDesc6.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove7_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
               // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund7.Visible = false;
            ddlFundType7.ClearSelection();
            ddlFund7.ClearSelection();
            txtFundDesc7.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove8_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
               // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");

                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund8.Visible = false;
            ddlFundType8.ClearSelection();
            ddlFund8.ClearSelection();
            txtFundDesc8.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove9_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();
               // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund9.Visible = false;
            ddlFundType9.ClearSelection();
            ddlFund9.ClearSelection();
            txtFundDesc9.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove10_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

               // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund10.Visible = false;
            ddlFundType10.ClearSelection();
            ddlFund10.ClearSelection();
            txtFundDesc10.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }

    protected void lnkRemove11_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund11.Visible = false;
            ddlFundType11.ClearSelection();
            ddlFund11.ClearSelection();
            txtFundDesc11.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove12_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund12.Visible = false;
            ddlFundType12.ClearSelection();
            ddlFund12.ClearSelection();
            txtFundDesc12.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove13_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund13.Visible = false;
            ddlFundType13.ClearSelection();
            ddlFund13.ClearSelection();
            txtFundDesc13.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove14_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund14.Visible = false;
            ddlFundType14.ClearSelection();
            ddlFund14.ClearSelection();
            txtFundDesc14.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove15_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund15.Visible = false;
            ddlFundType15.ClearSelection();
            ddlFund15.ClearSelection();
            txtFundDesc15.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove16_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund16.Visible = false;
            ddlFundType16.ClearSelection();
            ddlFund16.ClearSelection();
            txtFundDesc16.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove17_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund17.Visible = false;
            ddlFundType17.ClearSelection();
            ddlFund17.ClearSelection();
            txtFundDesc17.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove18_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund18.Visible = false;
            ddlFundType18.ClearSelection();
            ddlFund18.ClearSelection();
            txtFundDesc18.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove19_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund19.Visible = false;
            ddlFundType19.ClearSelection();
            ddlFund19.ClearSelection();
            txtFundDesc19.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void lnkRemove20_Click(object sender, EventArgs e)
    {
        string TemplateFundId = (((LinkButton)sender).CommandArgument);
        if (TemplateFundId != "")
        {
            //string strsql = " UPDATE ssi_templatefund SET deletionstatecode = 2 WHERE deletionstatecode = 0 AND ssi_templatefundid='" + TemplateFundId + "'";

            int i = 0;
            //SqlCommand command = new SqlCommand();
            //SqlConnection connection = new SqlConnection("Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01");

            try
            {

                //command.CommandText = strsql;
                //command.Connection = connection;

                //Mark As Stored Procedure
                //command.CommandType = CommandType.Text;

                //connection.Open();
                //i = command.ExecuteNonQuery();

                // i = DeleteData(TemplateFundId, EntityName.ssi_templatefund);
                i = DeleteData(TemplateFundId, "ssi_templatefund");
                if (i > 0)
                {
                    if (ddlTemplate.SelectedValue != "" && ddlTemplate.SelectedValue != "0")
                    {
                        BindValues(ddlTemplate.SelectedValue);
                    }
                }
                //connection.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                //if (connection.State == ConnectionState.Open) connection.Close();
                //command.Dispose();
                //connection.Dispose();
            }
            //return i;
        }
        else
        {
            trFund20.Visible = false;
            ddlFundType20.ClearSelection();
            ddlFund20.ClearSelection();
            txtFundDesc20.Value = "";
            hdtblId.Value = Convert.ToString(Convert.ToInt32(hdtblId.Value) - 1);
        }
    }
    protected void btnAddFund_Click(object sender, EventArgs e)
    {
        #region validate save
        if (rdoYes.Checked == true)
        {
            // for (int i = 1; i < 11; i++)
            for (int i = 1; i < 12; i++)
            {

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();

                Control trFund = FindControl("trFund" + i.ToString());
                Control ddlFundType = FindControl("ddlFundType" + i.ToString());
                Control ddlFund = FindControl("ddlFund" + i.ToString());
                Control ddlFund1 = FindControl("ddlFund" + i.ToString());
                Control htmlEditor = FindControl("txtFundDesc" + i.ToString());
                DropDownList Fund = null;
                DropDownList Fund1 = null;
                if (trFund != null)
                {
                    if (trFund.Visible == true)
                    {
                        if (ddlFund != null)
                        {
                            Fund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                            //Fund.SelectedValue = "0";
                            if (Fund.SelectedValue == "" || Fund.SelectedValue == "0")
                            {
                                sb.Append("\n<script type=text/javascript>\n");
                                sb.Append("\n alert('Please select fund.');");
                                sb.Append("</script>");
                                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                                return;
                            }
                        }

                        //if (ddlFund != null)
                        //{
                        //    if (Fund.SelectedValue == Fund.SelectedValue)
                        //    {
                        //        //Fund.Focus();
                        //        sb.Append("\n<script type=text/javascript>\n");
                        //        sb.Append("\n alert('Please select some other fund from fund.');");
                        //        sb.Append("</script>");
                        //        ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                        //        return;
                        //    }
                        //}

                        if (ddlFundType != null)
                        {
                            DropDownList FundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
                            // FundType.SelectedValue = "0";
                            if (FundType.SelectedValue != "" || FundType.SelectedValue != "0")
                            {
                                if (Fund.SelectedValue == "" || Fund.SelectedValue == "0")
                                {
                                    sb.Append("\n<script type=text/javascript>\n");
                                    sb.Append("\n alert('Please select fund.');");
                                    sb.Append("</script>");
                                    ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                                    return;
                                }

                            }
                        }

                        FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + Convert.ToString(i)));
                        //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                        //txtFundDesc.Text = "";
                        if (txtFundDesc.Value == "")
                        {
                            sb.Append("\n<script type=text/javascript>\n");
                            sb.Append("\n alert('Please enter fund details.');");
                            sb.Append("</script>");
                            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                            return;
                        }
                    }
                }
            }
        }
        #endregion

        if (rdoYes.Checked == true)
        {
            trAddAnother.Visible = true;

            int val = Convert.ToInt32(hdtblId.Value) + 1;
            // if (val < 11)
            // if (val < 12)
           // if (val < 16) // added (11_21_2018)
                if (val < 21) // added (11_23_2018)
                {
                hdtblId.Value = Convert.ToString(val);
                Control trFund = FindControl("trFund" + val);
                Control ddlFundType = FindControl("ddlFundType" + val);
                Control ddlFund = FindControl("ddlFund" + val);

                if (trFund != null)
                {
                    trFund.Visible = true;
                }

                if (ddlFund != null)
                {
                    DropDownList Fund = ((DropDownList)FindControl("ddlFund" + val));
                   Fund.SelectedValue = "0";
                }

                if (ddlFund != null)
                {
                    DropDownList FundType = ((DropDownList)FindControl("ddlFundType" + val));
                    FundType.SelectedValue = "0";
                }

                FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + val));
                //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + val));
                txtFundDesc.Value = "";
            }
        }
        else if (rdoNo.Checked == true)
        {
            hdtblId.Value = "1";
            trAddAnother.Visible = false;
            // for (int i = 1; i < 11; i++)
            // for (int i = 1; i < 12; i++)
            // for (int i = 1; i < 16; i++)// added (11_21_2018)
            for (int i = 1; i < 21; i++)// added (11_23_2018)
            {
                Control trFund = FindControl("trFund" + i.ToString());
                Control ddlFund = FindControl("ddlFund" + i.ToString());
                Control ddlFundType = FindControl("ddlFundType" + i.ToString());

                if (ddlFund != null)
                {
                    DropDownList drpFund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                    drpFund.SelectedValue = "0";
                }

                if (ddlFundType != null)
                {
                    DropDownList drpFundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
                    drpFundType.SelectedValue = "0";
                }

                FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
                //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                txtFundDesc.Value = "";

                if (trFund != null)
                {
                    trFund.Visible = false;
                }
            }
        }
    }
    protected void rdoYes_CheckedChanged(object sender, EventArgs e)
    {
        trAddAnother.Visible = true;

        int val = Convert.ToInt32(hdtblId.Value) + 1;
        // if (val < 11)
        //if (val < 12)
        // if (val < 16)// added (11_21_2018)
        if (val < 21)// added (11_23_2018)
        {
            hdtblId.Value = Convert.ToString(val);
            Control trFund = FindControl("trFund" + val);

            if (trFund != null)
            {
                trFund.Visible = true;
            }
        }
    }
    protected void rdoNo_CheckedChanged(object sender, EventArgs e)
    {
        hdtblId.Value = "1";
        trAddAnother.Visible = false;
        // for (int i = 1; i < 11; i++)
        //for (int i = 1; i < 12; i++)
        //for (int i = 1; i < 16; i++)// added (11_21_2018)
        for (int i = 1; i < 21; i++)// added (11_23_2018)
        {
            Control trFund = FindControl("trFund" + i.ToString());
            Control ddlFund = FindControl("ddlFund" + i.ToString());
            Control ddlFundType = FindControl("ddlFundType" + i.ToString());
            
            if (ddlFund != null)
            {
                DropDownList drpFund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                drpFund.SelectedValue = "0";
            }

            if (ddlFundType != null)
            {
                DropDownList drpFundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
                drpFundType.SelectedValue = "0";
            }

            FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
            //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
            txtFundDesc.Value = "";
           
            if (trFund != null)
            {
                trFund.Visible = false;
            }
        }
    }

    private void ShowHideBindValues(DataSet DS,int num)
    {
        if (FundSpecific == false)
        {
            Unit objUnit = new Unit(80, UnitType.Percentage);
            if (ddlTemplateType.SelectedValue == "10")
            {
                trFundSpecific.Visible = false;
                lblFileName.Visible = true;
            }
            else
            {
                lblFileName.Visible = false;
                trFundSpecific.Visible = true;
            }
            rdoNo.Checked = true;
            trAddAnother.Visible = false;
           
            txtLetterText.Width = objUnit;
            hdtblId.Value = Convert.ToString(DS.Tables[0].Rows.Count);
            int val = num + 1;// Convert.ToInt32(hdtblId.Value) + 1;
            Control trFund = FindControl("trFund" + num.ToString());

            if (trFund != null)
            {
                trFund.Visible = false;
            }

            // for (int l = val; l < 11; l++)
            // for (int l = val; l < 12; l++)
            //for (int l = val; l < 16; l++)// added (11_21_2018)
            for (int l = val; l < 21; l++)// added (11_23_2018)
            {
                //hdtblId.Value = Convert.ToString(val);
                Control trFund1 = FindControl("trFund" + l.ToString());

                if (trFund1 != null)
                {
                    trFund1.Visible = false;
                }
            }
        }
        else if (FundSpecific == true)
        {
            rdoYes.Checked = true;
            rdoNo.Checked = false;
            trFundSpecific.Visible = false;
            trAddAnother.Visible = true;
            hdtblId.Value = Convert.ToString(DS.Tables[0].Rows.Count);
            int val = Convert.ToInt32(hdtblId.Value) + 1;
            Control trFund = FindControl("trFund" + num.ToString());

            if (trFund != null)
            {
                trFund.Visible = true;
            }

            // for (int l = val; l < 11; l++)
            // for (int l = val; l < 12; l++)
            //for (int l = val; l < 16; l++)// added (11_21_2018)
            for (int l = val; l < 21; l++)// added (11_23_2018)
            {
                //hdtblId.Value = Convert.ToString(val);
                Control trFund1 = FindControl("trFund" + l.ToString());

                if (trFund1 != null)
                {
                    trFund1.Visible = false;
                }
            }
        }
    }

    protected void rdoDynamic_CheckedChanged(object sender, EventArgs e)
    {
        trOrientation.Visible = false;
        trSignatureText.Visible = false;
    }
    protected void raoStatic_CheckedChanged(object sender, EventArgs e)
    {
        trOrientation.Visible = true;
        trSignatureText.Visible = true;
    }
    protected void ddlTemplateType_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlTemplateType.SelectedValue == "10")
        {
            lblFileName.Text = FileUpload1.FileName;
            trFileupload.Visible = true;
            rdoNo.Checked = true;
            trFundSpecific.Visible = false;
            trDynamic.Visible = false;
            hdtblId.Value = "1";
            trAddAnother.Visible = false;
            // for (int i = 1; i < 11; i++)
            // for (int i = 1; i < 12; i++)
            // for (int i = 1; i < 16; i++)// added (11_21_2018)
            for (int i = 1; i < 21; i++)// added (11_23_2018)
            {
                Control trFund = FindControl("trFund" + i.ToString());
                Control ddlFund = FindControl("ddlFund" + i.ToString());
                Control ddlFundType = FindControl("ddlFundType" + i.ToString());

                if (ddlFund != null)
                {
                    DropDownList drpFund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
                    drpFund.SelectedValue = "0";
                }

                if (ddlFundType != null)
                {
                    DropDownList drpFundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
                    drpFundType.SelectedValue = "0";
                }

                FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
                //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
                txtFundDesc.Value = "";

                if (trFund != null)
                {
                    trFund.Visible = false;
                }
            }
        }
        else
        {
            trFileupload.Visible = false;
            hdtblId.Value = "1";
            trAddAnother.Visible = true;
            trFund1.Visible = true;
            //for (int i = 1; i < 11; i++)
            //{
            //    Control trFund = FindControl("trFund" + i.ToString());
            //    Control ddlFund = FindControl("ddlFund" + i.ToString());
            //    Control ddlFundType = FindControl("ddlFundType" + i.ToString());

            //    if (ddlFund != null)
            //    {
            //        DropDownList drpFund = ((DropDownList)FindControl("ddlFund" + i.ToString()));
            //        drpFund.SelectedValue = "0";
            //    }

            //    if (ddlFundType != null)
            //    {
            //        DropDownList drpFundType = ((DropDownList)FindControl("ddlFundType" + i.ToString()));
            //        drpFundType.SelectedValue = "0";
            //    }

            //    FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc" + i.ToString()));
            //    //HtmlEditor txtFundDesc = ((HtmlEditor)FindControl("txtFundDesc" + i.ToString()));
            //    txtFundDesc.Value = "";


            //    if (trFund != null)
            //    {
            //        trFund.Visible = false;
            //    }
            //}

            //if (rdoYes.Checked == true && hdtblId.Value == "1")
            //{
            //    trFund1.Visible = true;
            //    HideTR();
            //}
        }

    }

    //private int DeleteData(string DeleteID, EntityName entityName)
    private int DeleteData(string DeleteID, string entityName)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        //string orgName = "GreshamPartners";
        //CrmService service = null;
        IOrganizationService service = null;

        int successcount = 0;
        try
        {
            string UserId = GetcurrentUser(false);
            //service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
       // catch (System.Web.Services.Protocols.SoapException exc)
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            //bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
            //sw.WriteLine(strDescription);
        }
        catch (Exception exc)
        {
            //bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
            //sw.WriteLine(strDescription);
            lblError.Text = "Error in remove : " + exc.Message.ToString();
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            //for (int j = 0; j < dt.Rows.Count; j++)
            //{
            Guid UUID = new Guid(Convert.ToString(DeleteID));
            service.Delete(entityName.ToString(), UUID);
            successcount = successcount + 1;
            //}
        }
        catch (Exception ex)
        {
            //Response.Write("<br/>" + ex.Message);
            //sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "TNR Load");
            lblError.Text = "Error in remove : " + ex.Message.ToString();
        }
        return successcount;
    }
}
