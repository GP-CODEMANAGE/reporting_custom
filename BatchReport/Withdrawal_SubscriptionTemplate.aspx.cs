using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class BatchReport_Withdrawal_SubscriptionTemplate : System.Web.UI.Page
{

    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    int recommendationCount = 0;
    int SuccessCount = 0;
    protected void Page_Load(object sender, EventArgs e)
    {
        //string NAme = "Gresham Partnerships_GGES ---- Domestic ETFs_GP China Venture Capital II, L_Additional,, Contribution_20180604.pdf";
        //NAme = RemoveSpecialCharacters(NAme);

        // string amount= string.Format("{0:$#,##0.00}", 58445);
        // GPRedemeptionRequestForm();
        if (!IsPostBack)
        {
            BindAssociate(ddlAssociate);
            BindHousehold(ddlHouseHold);
            BindLegalEntity(ddlLegalEntity);
            BindStatus(ddlStatus);
            BindGridView();
            #region CHECK
            //try
            //{
            //    BindAssociate(ddlAssociate);
            //    BindHousehold(ddlHouseHold);
            //    BindLegalEntity(ddlLegalEntity);
            //    BindGridView();

            //    foreach (GridViewRow row in GridView1.Rows)
            //    {
            //        string RecommendationId = row.Cells[9].Text;                   
            //        string FinalFileName = row.Cells[10].Text;
            //        string HouseholdId = row.Cells[1].Text;
            //        string LegalEntity = row.Cells[2].Text;
            //        string CloseDate = row.Cells[3].Text;
            //        string Fund = row.Cells[4].Text;
            //        string Amount = row.Cells[11].Text;
            //        string Signer1Name = row.Cells[12].Text;
            //        string Signer1Title = row.Cells[13].Text;
            //        string Signer2Name = row.Cells[14].Text;
            //        string Signer2Title = row.Cells[15].Text;
            //        string Signer3Name = row.Cells[16].Text;
            //        string Signer3Title = row.Cells[17].Text;

            //        string[] SignerNames = new string[3];
            //        string[] SignerTitle = new string[3];

            //        SignerNames[0] = Signer1Name.Replace("&nbsp;", "");
            //        SignerNames[1] = Signer2Name.Replace("&nbsp;", "");
            //        SignerNames[2] = Signer3Name.Replace("&nbsp;", "");

            //        SignerTitle[0] = Signer1Title.Replace("&nbsp;", "");
            //        SignerTitle[1] = Signer2Title.Replace("&nbsp;", "");
            //        SignerTitle[2] = Signer3Title.Replace("&nbsp;", "");

            //        DataTable dtReportType = GetDataTable(RecommendationId);
            //        string ReportType = Convert.ToString(dtReportType.Rows[0]["Report"]);



            //        if (ReportType.ToLower() == "gpes")
            //        {                       
            //            string SignerFile = SignerPage(LegalEntity, SignerNames, SignerTitle);                     
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    Response.Write("ERRR : " +ex.Message.ToString());
            //}
            #endregion


        }
    }
    public void BindAssociate(DropDownList ddl)
    {

        string sqlstr = string.Empty;
        object AssociateId = ddlAssociate.SelectedValue == "0" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        ddl.Items.Clear();

        sqlstr = "EXEC SP_S_ASSOCIATE @ActRecmmFlg = 1";
        clsGM.getListForBindDDL(ddl, sqlstr, "Ssi_SecondaryOwnerIdName", "SSi_SecondaryOwnerId");

        if (ddl.Items.Count == 1)
        {
            if (ddl.Items[0].Value == "0")
                ddl.Items.Remove(ddl.Items[0]);
        }
        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }
    public void BindHousehold(DropDownList ddl)
    {
        string sqlstr = string.Empty;
        object AssociateId = ddlAssociate.SelectedValue == "0" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        ddl.Items.Clear();

        sqlstr = "EXEC sp_s_HouseHoldName @ActRecmmFlg = 1 , @AssociateId = " + AssociateId;
        clsGM.getListForBindDDL(ddl, sqlstr, "Name", "AccountId");

        if (ddl.Items.Count == 1)
        {
            if (ddl.Items[0].Value == "0")
                ddl.Items.Remove(ddl.Items[0]);
        }
        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }
    public void BindLegalEntity(DropDownList ddl)
    {
        string sqlstr = string.Empty;
        object HouseholdId = ddlHouseHold.SelectedValue == "0" ? "null" : "'" + ddlHouseHold.SelectedValue + "'";
        object AssociateId = ddlAssociate.SelectedValue == "0" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        ddl.Items.Clear();

        sqlstr = "EXEC SP_S_LEGAL_ENTITY_LIST @ActRecmmFlg = 1 ,@HouseHoldID =" + HouseholdId + ", @AssociateId = " + AssociateId;
        clsGM.getListForBindDDL(ddl, sqlstr, "LegalEntityName", "LegalEntityNameID");

        if (ddl.Items.Count == 1)
        {
            if (ddl.Items[0].Value == "0")
                ddl.Items.Remove(ddl.Items[0]);
        }
        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }
    public void BindStatus(DropDownList ddl)
    {
        string sqlstr = string.Empty;
        //object HouseholdId = ddlHouseHold.SelectedValue == "0" ? "null" : "'" + ddlHouseHold.SelectedValue + "'";
        //object AssociateId = ddlAssociate.SelectedValue == "0" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        ddl.Items.Clear();

        sqlstr = "Exec SP_S_SUBSCRIPTION_Status";
        clsGM.getListForBindDDL(ddl, sqlstr, "StatusIDName", "StatusID");

        if (ddl.Items.Count == 1)
        {
            if (ddl.Items[0].Value == "0")
                ddl.Items.Remove(ddl.Items[0]);
        }
        //ddl.Items.Insert(0, "All");
        //ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 2;
    }
    public void ClearControls()
    {
        lblError.Text = "";
        lblSuccess.Text = "";
    }
    protected void ddlAssociate_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        BindHousehold(ddlHouseHold);
        BindLegalEntity(ddlLegalEntity);
        BindGridView();
    }
    protected void ddlHouseHold_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindAssociate(ddlAssociate);
        BindLegalEntity(ddlLegalEntity);
        BindGridView();
    }
    protected void ddlLegalEntity_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindAssociate(ddlAssociate);
        //BindHousehold(ddlHouseHold);
        BindGridView();
    }
    protected void ddlStatus_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        BindGridView();
    }
    public void BindGridView()
    {
        string sql = string.Empty;
        object HouseholdId = ddlHouseHold.SelectedValue == "0" || ddlHouseHold.SelectedValue == "" ? "null" : "'" + ddlHouseHold.SelectedValue + "'";
        object AssociateId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object LegalEntityId = ddlLegalEntity.SelectedValue == "0" || ddlLegalEntity.SelectedValue == "" ? "null" : "'" + ddlLegalEntity.SelectedValue + "'";
        object StatusId = ddlStatus.SelectedValue == "0" || ddlStatus.SelectedValue == "" ? "null" : "'" + ddlStatus.SelectedValue + "'";

        sql = "EXEC SP_S_SUBSCRIPTION_RECOMM_LIST @StatusID=" + StatusId + ", @HouseHoldID =" + HouseholdId + ", @AssociateId = " + AssociateId + ", @LegalEntityID = " + LegalEntityId;
        DataSet loDataset = clsDB.getDataSet(sql);

        GridView1.Columns[9].Visible = true;
        GridView1.Columns[10].Visible = true;
        GridView1.Columns[11].Visible = true;// Year
        GridView1.Columns[12].Visible = true;//GRAS Year Series
        GridView1.Columns[13].Visible = true;//GPES Year Series
        #region Commented - 12_20_2018
        /* Account Signer*/
        //GridView1.Columns[11].Visible = true;
        //GridView1.Columns[12].Visible = true;
        //GridView1.Columns[13].Visible = true;
        //GridView1.Columns[14].Visible = true;
        //GridView1.Columns[15].Visible = true;
        //GridView1.Columns[16].Visible = true;
        /* Legal Entity Signer*/
        //GridView1.Columns[17].Visible = true;
        //GridView1.Columns[18].Visible = true;
        //GridView1.Columns[19].Visible = true;
        //GridView1.Columns[20].Visible = true;
        //GridView1.Columns[21].Visible = true;
        //GridView1.Columns[22].Visible = true;
        //GridView1.Columns[23].Visible = true;//show flg
        ////GridView1.Columns[17].Visible = true;
        #endregion

        GridView1.DataSource = loDataset;
        GridView1.DataBind();


        GridView1.Columns[9].Visible = false;
        GridView1.Columns[10].Visible = false;
       GridView1.Columns[11].Visible = false;// Year
        GridView1.Columns[12].Visible = false;//GRAS Year Series
        GridView1.Columns[13].Visible = false;//GPES Year Series

        #region Commented - 12_20_2018
        //GridView1.Columns[11].Visible = false;
        //GridView1.Columns[12].Visible = false;
        //GridView1.Columns[13].Visible = false;
        //GridView1.Columns[14].Visible = false;
        //GridView1.Columns[15].Visible = false;
        //GridView1.Columns[16].Visible = false;
        //GridView1.Columns[17].Visible = false;
        //GridView1.Columns[18].Visible = false;
        //GridView1.Columns[19].Visible = false;
        //GridView1.Columns[20].Visible = false;
        //GridView1.Columns[21].Visible = false;
        //GridView1.Columns[22].Visible = false;
        //GridView1.Columns[23].Visible = false;
        ////GridView1.Columns[17].Visible = false;
        #endregion


        if (GridView1.Rows.Count < 1)
        {
            lblCount.Visible = false;
            lblError.Text = "Record not found";
            lblError.Visible = true;
            return;
        }
        else
        {
            lblCount.Visible = true;
            lblCount.Text = "Count: " + GridView1.Rows.Count.ToString();
            lblError.Visible = false;
        }
    }
    public string PrivateEquityStrategiesCommitment(string Amount, string LegalEntity, string CloseDate, string Year, string Series)
    {
        string file = string.Empty;
        iTextSharp.text.Document document = null;
        try
        {
            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 5, 2, 30, 10);//10,10
            string FolderPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\";
            string fileName = System.DateTime.Now.ToString("MMddyyhhmmss") + "PrivateEquityStrategiesCommitment.pdf";
            file = Path.Combine(FolderPath, fileName);
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(file, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            // AddFooter(document, "Gresham Partners LLC    333 W Wacker Drive, Suite 700 Chicago, IL 60606    (312) 960-0200  Fax (312) 960-0204" );
            // AddFooter(document, "Page 1 of 2");
            document.Open();

            PdfPTable tblHeader = new PdfPTable(2);
            tblHeader.TotalWidth = 100f;
            //  int[] headerwidths4 = { 25, 75 };
            //int[] headerwidths4 = { 1, 55, 44 };
            //int[] headerwidths4 = { 1, 24, 75 };
            int[] headerwidths4 = { 60, 40 };
            tblHeader.SetWidths(headerwidths4);

            #region Logo
          //  iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");

            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");


            png.SetAbsolutePosition(65, 815);//540
            png.ScalePercent(8.75f);
            document.Add(png);
            #endregion

            #region Header
            Paragraph lochunk = new Paragraph();
            PdfPCell loCell = new PdfPCell();

            float FontSize = 9;
            PdfPCell loCell1 = new PdfPCell();

            string ReportHeader1 = "Investor/Entity: " + LegalEntity;
            lochunk = new Paragraph(ReportHeader1, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            loCell1.AddElement(lochunk);
            loCell1.Border = 0;
            tblHeader.AddCell(loCell1);

            loCell1 = new PdfPCell();
            ReportHeader1 = "Close Date: " + CloseDate;
            lochunk = new Paragraph(ReportHeader1, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            loCell1.AddElement(lochunk);
            loCell1.Border = 0;
            tblHeader.AddCell(loCell1);

            loCell1 = new PdfPCell();
            ReportHeader1 = "";
            lochunk = new Paragraph(ReportHeader1, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            loCell1.AddElement(lochunk);
            loCell1.Border = 0;
            loCell1.Colspan = 2;
            tblHeader.AddCell(loCell1);

            FontSize = 12;
            // string ReportHeader = "GP 2018 Private Equity Strategies Commitment Agreement";
            // string ReportHeader = "GP 2019 Private Equity Strategies Commitment Agreement"; //changed sasmit(10_2_2018)
            string ReportHeader = "";
            if (Year != "")
            {
                 ReportHeader = "GP " + Year + " Private Equity Strategies Commitment Agreement"; //changed sasmit(12_26_2018)
            }
            else
            {
                 ReportHeader = "GP Private Equity Strategies Commitment Agreement"; //changed sasmit(12_26_2018)
            }

           
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            loCell.AddElement(lochunk);
            loCell.Border = 0;

            //loHeader.AddCell(loCell);


            ReportHeader = "Gresham Private Equity Strategies, L.P.";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            lochunk.SetLeading(13, 0);

            loCell.AddElement(lochunk);
            //loHeader.AddCell(loCell);

            FontSize = 9f;
            ReportHeader = "\nGresham Private Equity Strategies, L.P. \nc/o Gresham Advisors, L.L.C.\n333 West Wacker Drive\nSuite 700 \nChicago, Illinois 60606 \n\nDear Sir or Madam:";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.SetLeading(13, 0);
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");

            loCell.AddElement(lochunk);

            //ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Private Equity Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP 2018 Private Equity Strategies (the “Series”) in the amount of: " + Amount + ".";
            // ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Private Equity Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP 2019 Private Equity Strategies (the “Series”) in the amount of: " + Amount + "."; //changed sasmit(10_2_2018)
            if (Year != "")
            {
                ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Private Equity Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP " + Year + " Private Equity Strategies (the “Series”) in the amount of: " + Amount + "."; //changed sasmit(12_26_2018)
            }
            else
            {
                ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Private Equity Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP Private Equity Strategies (the “Series”) in the amount of: " + Amount + "."; //changed sasmit(12_26_2018)
            }
                
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);

            loCell.AddElement(lochunk);

            // ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the 2011/2012 Series, 2013 Series, 2014 Series, 2015 Series, 2016 Series and/or 2017 Series and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein. ";
            //ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the 2011/2012 Series, 2013 Series, 2014 Series, 2015 Series, 2016 Series, 2017 Series and/or 2018 Series and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein. "; //changed sasmit(10_2_2018)
            if (Series != "")
            {
                ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the " + Series + " and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein. "; //changed sasmit(12_26_2018)

            }
            else
            {
                ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein. "; //changed sasmit(12_26_2018)

            }
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);

            loCell.AddElement(lochunk);

            //ReportHeader = "\nThe undersigned hereby (in addition and not by way of limitation of the power of attorney as set forth in the Partnership Agreement) irrevocably constitutes and appoints the General Partner, its successors and assignee, and the officers of the foregoing, as the undersigned’s true and lawful Attorney-in-Fact, with full power of substitution, in the undersigned’s name, place and stead, to: (a) file, prosecute, defend, settle or compromise litigation, claims or arbitrations on behalf of the Series and/or the Partnership; (b) make, execute, sign, acknowledge, swear to, deliver, record and file any documents or instruments, including, without limitation, Certificates of Limited Partnership and amendments thereto, the Partnership Agreement and amendments thereto, that may be considered necessary or desirable by the General Partner to carry out fully the provisions of the Partnership Agreement, including, without limitation, those (if any) necessary or desirable to effect the undersigned’s admission as a Limited Partner; and (c) to perform all other acts contemplated by the Partnership Agreement.  This Power of Attorney shall be deemed to be coupled with an interest and shall be irrevocable and survive and not be affected by the undersigned’s subsequent death, incapacity, disability, insolvency or dissolution or any delivery by the undersigned of an assignment of the whole or any portion of the undersigned’s Interest.\nThis Agreement shall be governed in accordance with the laws of the State of Delaware (without regard to conflict of law principles).";
            ReportHeader = "\nThe undersigned hereby (in addition and not by way of limitation of the power of attorney as set forth in the Partnership Agreement) irrevocably constitutes and appoints the General Partner, its successors and assignees, and the officers of the foregoing, as the undersigned’s true and lawful Attorney-in-Fact, with full power of substitution, in the undersigned’s name, place and stead, to: (a) file, prosecute, defend, settle or compromise litigation, claims or arbitrations on behalf of the Series and/or the Partnership; (b) make, execute, sign, acknowledge, swear to, deliver, record and file any documents or instruments, including, without limitation, Certificates of Limited Partnership and amendments thereto, the Partnership Agreement and amendments thereto, that may be considered necessary or desirable by the General Partner to carry out fully the provisions of the Partnership Agreement, including, without limitation, those (if any) necessary or desirable to effect the undersigned’s admission as a Limited Partner; and (c) to perform all other acts contemplated by the Partnership Agreement.  This Power of Attorney shall be deemed to be coupled with an interest and shall be irrevocable and survive and not be affected by the undersigned’s subsequent death, incapacity, disability, insolvency or dissolution or any delivery by the undersigned of an assignment of the whole or any portion of the undersigned’s Interest.";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);

            loCell.AddElement(lochunk);

            ReportHeader = "\n(Signature page to follow)";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(9, 1, 1));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            lochunk.SetLeading(13, 0);

            loCell.AddElement(lochunk);


            // ReportHeader = "\n\nPage 1 of 2";
            // lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            // lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            //// lochunk.Alignment = Element.ALIGN_BOTTOM;
            // loCell.AddElement(lochunk);


            //lochunk = new Paragraph(" ", setFontsAllVerdana(FontSize, 1, 0));
            //lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            //loCell.AddElement(lochunk);
            // loHeader.AddCell(loCell);
            //  loCell.Border = 1;
            string Address = "Gresham Partners LLC    333 W Wacker Drive, Suite 700 Chicago, IL 60606    (312) 960-0200  Fax (312) 960-0204";
            PdfPTable TabFooter = addFooterAddress(Address, 1, 2);
            TabFooter.HorizontalAlignment = Element.ALIGN_CENTER;
            TabFooter.WidthPercentage = 100f;
            //  TabFooter.TotalWidth = 100f;
            TabFooter.TotalWidth = 600;


            TabFooter.WriteSelectedRows(0, 4, 0, 40, writer.DirectContent);
            loCell.Colspan = 2;
            tblHeader.AddCell(loCell);

            document.Add(tblHeader);

            // document.Add(new Phrase("\n"));
            #endregion

            document.Close();
        }
        catch (Exception ex)
        {
            file = "";
            lblError.Text = "Error :" + ex.Message.ToString();
            lblError.Visible = true;
            Response.Write("ERROR :" + ex.Message.ToString());
        }
        finally
        {
            document.Close();

        }

        return file;

    }//gpes
    public string RealAssetsStrategiesCommitment(string Amount, string LegalEntity, string CloseDate, string Year, string Series)
    {
        iTextSharp.text.Document document = null;
        string file = string.Empty;
        try
        {
            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 5, 2, 30, 10);//10,10
            string FolderPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\";
            string fileName = System.DateTime.Now.ToString("MMddyyhhmmss") + "RealAssetsStrategiesCommitment.pdf";
            file = Path.Combine(FolderPath, fileName);
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(file, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();

            PdfPTable loHeader = new PdfPTable(2);
            loHeader.TotalWidth = 100f;
            int[] headerwidths4 = { 60, 40 };
            loHeader.SetWidths(headerwidths4);

            #region Logo
            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
            png.SetAbsolutePosition(65, 815);//540
            png.ScalePercent(8.75f);
            document.Add(png);
            #endregion

            #region Header
            Paragraph lochunk = new Paragraph();
            PdfPCell loCell = new PdfPCell();


            float FontSize = 9;
            PdfPCell loCell1 = new PdfPCell();

            string ReportHeader1 = "Investor/Entity: " + LegalEntity;
            lochunk = new Paragraph(ReportHeader1, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            loCell1.AddElement(lochunk);
            loCell1.Border = 0;
            loHeader.AddCell(loCell1);

            loCell1 = new PdfPCell();
            ReportHeader1 = "Close Date: " + CloseDate;
            lochunk = new Paragraph(ReportHeader1, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            loCell1.AddElement(lochunk);
            loCell1.Border = 0;
            loHeader.AddCell(loCell1);

            loCell1 = new PdfPCell();
            ReportHeader1 = "";
            lochunk = new Paragraph(ReportHeader1, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            loCell1.AddElement(lochunk);
            loCell1.Border = 0;
            loCell1.Colspan = 2;
            loHeader.AddCell(loCell1);



            FontSize = 12;
            // string ReportHeader = "GP 2018 Real Assets Strategies Commitment Agreement";
            // string ReportHeader = "GP 2019 Real Assets Strategies Commitment Agreement"; // changed sasmit(10_2_2018)
            string ReportHeader = "";
            if (Year != "")
            {
                ReportHeader = "GP " + Year + " Real Assets Strategies Commitment Agreement"; // changed sasmit(12_26_2018)
            }
            else
            {
                ReportHeader = "GP Real Assets Strategies Commitment Agreement"; // changed sasmit(12_26_2018)
            }
             
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            loCell.AddElement(lochunk);
            //lochunk.SetLeading(13, 0);
            loCell.Border = 0;
            //loHeader.AddCell(loCell);


            ReportHeader = "Gresham Real Assets Strategies, L.P.";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);
            //loHeader.AddCell(loCell);

            FontSize = 9;
            ReportHeader = "\nGresham Real Asset Strategies, L.P. \nc/o Gresham Advisors, L.L.C.\n333 West Wacker Drive\nSuite 700 \nChicago, Illinois 60606 \n \nDear Sir or Madam:";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);

            // ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Real Assets Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP 2018 Real Assets Strategies (the “Series”) in the amount of: " + Amount + ".";
            //  ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Real Assets Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP 2019 Real Assets Strategies (the “Series”) in the amount of: " + Amount + ".";//changed sasmit(10_2_2018
            if (Year != "")
            {
                ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Real Assets Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP " + Year + " Real Assets Strategies (the “Series”) in the amount of: " + Amount + ".";//changed sasmit(12_26_2018
            }
            else
            {
                ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Real Assets Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP Real Assets Strategies (the “Series”) in the amount of: " + Amount + ".";//changed sasmit(12_26_2018
            }
                
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);

            // ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the 2011/2012 Series, 2013 Series, 2013 Natural Resources Series, 2014 Series, 2015 Series, 2016 Series and/or 2017 Series and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein.";
            //  ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the 2011/2012 Series, 2013 Series, 2013 Natural Resources Series, 2014 Series, 2015 Series, 2016 Series, 2017 Series and/or 2018 Series and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein.";
            if (Series != "")
            {
                ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the " + Series + " and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein.";//Changed (12_26_018)

            }
            else
            {
                ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein.";//Changed (12_26_018)

            }
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);

            // ReportHeader = "\nThe undersigned hereby (in addition and not by way of limitation of the power of attorney as set forth in the Partnership Agreement) irrevocably constitutes and appoints the General Partner, its successors and assignee, and the officers of the foregoing, as the undersigned’s true and lawful Attorney-in-Fact, with full power of substitution, in the undersigned’s name, place and stead, to: (a) file, prosecute, defend, settle or compromise litigation, claims or arbitrations on behalf of the Series and/or the Partnership; (b) make, execute, sign, acknowledge, swear to, deliver, record and file any documents or instruments, including, without limitation, Certificates of Limited Partnership and amendments thereto, the Partnership Agreement and amendments thereto, that may be considered necessary or desirable by the General Partner to carry out fully the provisions of the Partnership Agreement, including, without limitation, those (if any) necessary or desirable to effect the undersigned’s admission as a Limited Partner; and (c) to perform all other acts contemplated by the Partnership Agreement.  This Power of Attorney shall be deemed to be coupled with an interest and shall be irrevocable and survive and not be affected by the undersigned’s subsequent death, incapacity, disability, insolvency or dissolution or any delivery by the undersigned of an assignment of the whole or any portion of the undersigned’s Interest.\nThis Agreement shall be governed in accordance with the laws of the State of Delaware (without regard to conflict of law principles).";
            ReportHeader = "\nThe undersigned hereby (in addition and not by way of limitation of the power of attorney as set forth in the Partnership Agreement) irrevocably constitutes and appoints the General Partner, its successors and assignees, and the officers of the foregoing, as the undersigned’s true and lawful Attorney-in-Fact, with full power of substitution, in the undersigned’s name, place and stead, to: (a) file, prosecute, defend, settle or compromise litigation, claims or arbitrations on behalf of the Series and/or the Partnership; (b) make, execute, sign, acknowledge, swear to, deliver, record and file any documents or instruments, including, without limitation, Certificates of Limited Partnership and amendments thereto, the Partnership Agreement and amendments thereto, that may be considered necessary or desirable by the General Partner to carry out fully the provisions of the Partnership Agreement, including, without limitation, those (if any) necessary or desirable to effect the undersigned’s admission as a Limited Partner; and (c) to perform all other acts contemplated by the Partnership Agreement.  This Power of Attorney shall be deemed to be coupled with an interest and shall be irrevocable and survive and not be affected by the undersigned’s subsequent death, incapacity, disability, insolvency or dissolution or any delivery by the undersigned of an assignment of the whole or any portion of the undersigned’s Interest.";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);

            ReportHeader = "\n(Signature page to follow)";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 1));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);


            //ReportHeader = "\n\nPage 1 of 2";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            // loCell.AddElement(lochunk);


            //ReportHeader = "IN ANY OF THE FOREGOING INFORMATION, REPRESENTATIONS OR WARRANTIES";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            //lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            //loCell.AddElement(lochunk);


            //lochunk = new Paragraph(" ", setFontsAllVerdana(FontSize, 1, 0));
            //lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            //loCell.AddElement(lochunk);
            // loHeader.AddCell(loCell);
            //  loCell.Border = 1;
            loCell.Colspan = 2;
            loHeader.AddCell(loCell);

            document.Add(loHeader);

            // document.Add(new Phrase("\n"));
            #endregion

            #region Footer
            string Address = "Gresham Partners LLC    333 W Wacker Drive, Suite 700 Chicago, IL 60606    (312) 960-0200  Fax (312) 960-0204";
            PdfPTable TabFooter = addFooterAddress(Address, 1, 2);
            TabFooter.HorizontalAlignment = Element.ALIGN_CENTER;
            TabFooter.WidthPercentage = 100f;
            //  TabFooter.TotalWidth = 100f;
            TabFooter.TotalWidth = 600;
            TabFooter.WriteSelectedRows(0, 4, 0, 40, writer.DirectContent);
            #endregion

            document.Close();
        }
        catch (Exception ex)
        {
            file = "";
            Response.Write("ERROR :" + ex.Message.ToString());
            lblError.Visible = true;
            lblError.Text = "ERROR: " + ex.Message.ToString();
        }
        finally
        {
            document.Close();

        }

        return file;

    }//Gras
    public string GPAdditionalSubscriptionForm(string LegalEntity, string CloseDate, string FundName, string Amount, string[] SignerNames, string[] SignerTitle)
    {
        iTextSharp.text.Document document = null;
        string file = string.Empty;
        try
        {
            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 5, 2, 30, 10);//10,10
            string FolderPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\";
            string fileName = System.DateTime.Now.ToString("MMddyyhhmmss") + "GPAdditionalSubscriptionForm.pdf";
            file = Path.Combine(FolderPath, fileName);
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(file, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            // AddFooter(document, "Gresham Partners LLC    333 W Wacker Drive, Suite 700 Chicago, IL 60606    (312) 960-0200  Fax (312) 960-0204" );
            // AddFooter(document, "Page 1 of 2");
            document.Open();

            PdfPTable loHeader = new PdfPTable(1);
            PdfPCell loCell = new PdfPCell();
            Paragraph lochunk = new Paragraph();
            Chunk imge = new Chunk();
            #region Logo
            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
            png.SetAbsolutePosition(65, 815);//540
            png.ScalePercent(8.75f);
            //  document.Add(png);
            imge = new Chunk(png, 0, 0, true);
            loCell.AddElement(imge);
            #endregion

            #region Header

            float FontSize = 9;

            string ReportHeader = "Subscription Date: " + CloseDate;
            ReportHeader = "\n" + "\n" + ReportHeader;
            lochunk = new Paragraph(ReportHeader, setFontsAll(9, 1, 0, new iTextSharp.text.Color(150, 150, 150)));
            lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            loCell.AddElement(lochunk);
            loCell.Border = 0;

            FontSize = 12;
            document.Add(Chunk.NEWLINE);
            ReportHeader = "GP Additional Subscription Form";
            ReportHeader = "\n" + ReportHeader;
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            loCell.AddElement(lochunk);
            loCell.Border = 0;
            //loHeader.AddCell(loCell);

            FontSize = 9;
            ReportHeader = "\n" + FundName + " \n333 West Wacker Drive\nSuite 700 \nChicago, IL 60606 \n\nDear Sir or Madam:";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.SetLeading(13, 0);
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            loCell.AddElement(lochunk);

            ReportHeader = "\nThe undersigned hereby wishes to make an additional capital contribution to " + FundName + " (the \"Partnership\").  The amount to be contributed (\"Additional Capital Contribution\") is: " + Amount + ".";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);

            ReportHeader = "\nThe undersigned acknowledges and agrees:  (i) that the undersigned is making the Additional Subscription on the terms and conditions contained in the subscription agreement previously executed by the undersigned and accepted by the General Partner (the \"Subscription Agreement\"); (ii) that the representations and covenants of the undersigned contained in the Subscription Agreement and the anti-money laundering supplement thereto are true and correct in all material respects as of the date set forth below; (iii) the information provided on the Investor Profile Form in the Subscription Agreement is correct as of the date set forth below; and (iv) the background information provided to the General Partner is true and correct in all material respects as of the date set forth below.";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);

            ReportHeader = "\nTHE UNDERSIGNED AGREES TO NOTIFY THE GENERAL PARTNER \nPROMPTLY IN WRITING SHOULD THERE BE ANY CHANGE \nIN ANY OF THE FOREGOING INFORMATION.\n\n";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);

            //ReportHeader = "\nDated:                                  , 20\n\n";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");

            //loCell.AddElement(lochunk);


            PdfPTable tblEntity = new PdfPTable(2);
            tblEntity.TotalWidth = 100f;
            int[] headerwidths = { 30, 70 };
            tblEntity.SetWidths(headerwidths);

            PdfPCell cellhead = new PdfPCell(new Phrase("Dated:", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead.Border = 0;
            //cellhead.PaddingLeft = 7;
            tblEntity.AddCell(cellhead);

            PdfPCell cellhead1 = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = PdfCell.BOTTOM_BORDER;
            tblEntity.AddCell(cellhead1);

            cellhead = new PdfPCell(new Phrase("Investor/Entity:", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead.Border = 0;
            cellhead.PaddingTop = 5;
            //cellhead.PaddingLeft = 7;
            tblEntity.AddCell(cellhead);

            cellhead1 = new PdfPCell(new Phrase(LegalEntity, setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = PdfCell.BOTTOM_BORDER;
            cellhead1.PaddingTop = 5;
            tblEntity.AddCell(cellhead1);

            cellhead1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = 0;
            tblEntity.AddCell(cellhead1);

            cellhead1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = 0;
            tblEntity.AddCell(cellhead1);

            #region Table Signer(old)
            //PdfPTable tblSigner1 = new PdfPTable(2);

            //// iTextSharp.text.Font fontTable = FontFactory.GetFont("Arial", FontSize, iTextSharp.text.Font.UNDERLINE);
            //PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
            //cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //cell.Border = 0;
            //tblSigner1.AddCell(cell);

            //cell = new PdfPCell(new Phrase("value from Database", setUnderline(FontSize, 1)));
            //cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //cell.Border = 0;
            //tblSigner1.AddCell(cell);

            //cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
            //cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //cell.Border = 0;
            //tblSigner1.AddCell(cell);

            //cell = new PdfPCell(new Phrase("value from Database", setUnderline(FontSize, 1)));
            //cell.HorizontalAlignment = Element.ALIGN_LEFT;
            //cell.Border = 0;
            //tblSigner1.AddCell(cell);
            #endregion

            #region Table Signer
            PdfPTable tblSigner1 = new PdfPTable(2);
            tblSigner1.TotalWidth = 100f;
            int[] headerwidths1 = { 30, 70 };
            tblSigner1.SetWidths(headerwidths1);

            int countSigner = 0;
            for (int i = 0; i < 3; i++)
            {
                string signer = SignerNames[i].ToString();
                string Title = SignerTitle[i].ToString();
                if (signer != "")
                //if (signer.Contains("Neild, Edward") || signer.Contains("Beavers, Ben") || signer.Contains("Salsburg, David"))
                {
                    countSigner++;

                    #region not in use
                    //if (i == 0)
                    //{
                    //    // iTextSharp.text.Font fontTable = FontFactory.GetFont("Arial", FontSize, iTextSharp.text.Font.UNDERLINE);
                    //    PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    cell.PaddingLeft = 7;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    //    //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                    //    cell.Border = PdfCell.BOTTOM_BORDER;
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    //cell.Border = 0;            
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    cell.PaddingLeft = 7;
                    //    tblSigner1.AddCell(cell);

                    //    if (Title != "")
                    //    {
                    //        cell = new PdfPCell(new Phrase(signer + ", " + Title, setFontsAllVerdana(FontSize, 0, 0)));
                    //    }
                    //    else
                    //    {
                    //        cell = new PdfPCell(new Phrase(signer, setFontsAllVerdana(FontSize, 0, 0)));
                    //    }
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = PdfCell.BOTTOM_BORDER;
                    //    //cell.Colspan = 2;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);

                    //}


                    //if (i == 1)
                    //{
                    //    PdfPCell cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    cell.PaddingLeft = 7;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    //    //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                    //    cell.Border = PdfCell.BOTTOM_BORDER;
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    //cell.Border = 0;            
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);



                    //    cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    cell.PaddingLeft = 7;
                    //    tblSigner1.AddCell(cell);

                    //    if (Title != "")
                    //    {
                    //        cell = new PdfPCell(new Phrase(signer + ", " + Title, setFontsAllVerdana(FontSize, 0, 0)));
                    //    }
                    //    else
                    //    {
                    //        cell = new PdfPCell(new Phrase(signer, setFontsAllVerdana(FontSize, 0, 0)));
                    //    }
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = PdfCell.BOTTOM_BORDER;
                    //    //cell.Colspan = 2;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);

                    //}

                    //if (i == 2)
                    //{
                    //    PdfPCell cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    cell.PaddingLeft = 7;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    //    //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                    //    cell.Border = PdfCell.BOTTOM_BORDER;
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    //cell.Border = 0;            
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);



                    //    cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    cell.PaddingLeft = 7;
                    //    tblSigner1.AddCell(cell);

                    //    if (Title != "")
                    //    {
                    //        cell = new PdfPCell(new Phrase(signer + ", " + Title, setFontsAllVerdana(FontSize, 0, 0)));
                    //    }
                    //    else
                    //    {
                    //        cell = new PdfPCell(new Phrase(signer, setFontsAllVerdana(FontSize, 0, 0)));
                    //    }
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = PdfCell.BOTTOM_BORDER;
                    //    //cell.Colspan = 2;
                    //    tblSigner1.AddCell(cell);

                    //    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //    cell.Border = 0;
                    //    tblSigner1.AddCell(cell);

                    //}
                    #endregion

                    //countSigner++;
                    // iTextSharp.text.Font fontTable = FontFactory.GetFont("Arial", FontSize, iTextSharp.text.Font.UNDERLINE);
                    PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    //cell.PaddingLeft = 7;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                    cell.Border = PdfCell.BOTTOM_BORDER;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;            
                    tblSigner1.AddCell(cell);

                    //cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;
                    //tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    cell.PaddingTop = 5;
                    //cell.PaddingLeft = 7;
                    tblSigner1.AddCell(cell);

                    if (Title != "")
                    {
                        cell = new PdfPCell(new Phrase(signer + ", " + Title, setFontsAllVerdana(FontSize, 0, 0)));
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase(signer, setFontsAllVerdana(FontSize, 0, 0)));
                    }
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = PdfCell.BOTTOM_BORDER;
                    //cell.Colspan = 2;
                    cell.PaddingTop = 5;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    tblSigner1.AddCell(cell);

                    // tblSigner1.SpacingAfter =50;

                    //cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;
                    //tblSigner1.AddCell(cell);

                    //cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;
                    //tblSigner1.AddCell(cell);

                }




            }
            if (countSigner == 0)
            {
                PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                //cell.PaddingLeft = 7;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                cell.Border = PdfCell.BOTTOM_BORDER;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.Border = 0;            
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                cell.PaddingTop = 5;
                //cell.PaddingLeft = 7;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                cell.Border = PdfCell.BOTTOM_BORDER;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.PaddingTop = 5;
                //cell.Border = 0;            
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                tblSigner1.AddCell(cell);
            }
            #endregion
            #region intenalUse
            FontSize = 8;
            PdfPTable tblIntenal = new PdfPTable(2);
            tblIntenal.TotalWidth = 100f;
            int[] headerwidths5 = { 82, 18 };
            tblIntenal.SetWidths(headerwidths5);

            // iTextSharp.text.Font fontTable = FontFactory.GetFont("Arial", FontSize, iTextSharp.text.Font.UNDERLINE);
            PdfPCell cell1 = new PdfPCell(new Phrase("\n--------------------------------------------------------------------------------------------------------------------------------", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            FontSize = 10;
            cell1 = new PdfPCell(new Phrase("FOR INTERNAL USE ONLY\nTo be completed by Gresham Advisors, L.L.C.", setFontsAllVerdana(FontSize, 1, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            FontSize = 9;
            cell1 = new PdfPCell(new Phrase("ADDITIONAL CAPITAL CONTRIBUTION ACCEPTED\nAS TO $___________________________", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("\n" + FundName, setFontsAllVerdana(10, 1, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("By:  Gresham Advisors, L.L.C.\n\nBy:  _______________________\nDate: __________________, 20____", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            cell1.PaddingLeft = 85;
            tblIntenal.AddCell(cell1);

            iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\AdvisorInitial.png");
            dashjpg.ScalePercent(35);
            cell1 = new PdfPCell(dashjpg);
            // cell1.Colspan = 2;
            cell1.Border = 0;
            cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
            tblIntenal.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            tblIntenal.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Advisor Initials", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            tblIntenal.AddCell(cell1);

            #endregion
            #region footer
            //string line1 = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------\nFOR INTERNAL USE ONLY\nTo be completed by Gresham Advisors, L.L.C.";
            //string line2 = "\nCOMMITTMENT ACCEPTED\nAS TO $___________________________";
            //string line3 = "VALUE FROM DATABASE";
            //string line4 = "\nBy:  Gresham Advisors, L.L.C.\nBy:  _______________________\nDate: __________________, 20____";
            //string line5 = "Advisor Initials";
            //PdfPTable TabFooter = addFooterInternal(line1, line2, line3, line4, line5, true);
            //TabFooter.HorizontalAlignment = Element.ALIGN_CENTER;
            //TabFooter.WidthPercentage = 100f;
            ////  TabFooter.TotalWidth = 100f;
            //TabFooter.TotalWidth = 600;
            //TabFooter.WriteSelectedRows(0, 4, 0, 190, writer.DirectContent);

            string Address = "Gresham Partners LLC    333 W Wacker Drive, Suite 700 Chicago, IL 60606    (312) 960-0200  Fax (312) 960-0204";
            PdfPTable TabFooter1 = addFooterAddress(Address, 1, 1);
            TabFooter1.HorizontalAlignment = Element.ALIGN_CENTER;
            TabFooter1.WidthPercentage = 100f;
            //  TabFooter.TotalWidth = 100f;
            TabFooter1.TotalWidth = 600;

            TabFooter1.WriteSelectedRows(0, 4, 0, 38, writer.DirectContent);

            #endregion
            #region Advisor Initial Box
            //iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\AdvisorInitial.png");
            ////dashjpg.ScaleToFit(50f, 50f);
            ////dashjpg.SetAbsolutePosition(250f, 252);
            //if (countSigner == 0)
            //{
            //    dashjpg.SetAbsolutePosition(472, 311);
            //}
            //if (countSigner == 1)
            //{
            //    dashjpg.SetAbsolutePosition(472, 274);
            //}
            //else if (countSigner == 2)
            //{
            //    dashjpg.SetAbsolutePosition(472, 236);
            //}
            //else if (countSigner == 3)
            //{
            //    dashjpg.SetAbsolutePosition(472, 195);
            //}
            ////540
            //dashjpg.ScalePercent(35);
            ////dashjpg.IndentationLeft = 9f;
            ////dashjpg.SpacingAfter = 9f;
            //document.Add(dashjpg);
            #endregion
            loHeader.AddCell(loCell);

            document.Add(loHeader);
            document.Add(tblEntity);
            document.Add(tblSigner1);
            document.Add(tblIntenal);
            // document.Add(new Phrase("\n"));
            #endregion



            document.Close();

        }
        catch (Exception ex)
        {
            file = "";
            lblError.Visible = true;
            lblError.Text = "ERROR :" + ex.Message.ToString();
            Response.Write("ERROR :" + ex.Message.ToString());
        }
        finally
        {
            document.Close();

        }

        return file;

    }
    public string GPRedemeptionRequestForm(DataSet ds_gresham, string[] SignerNames, string[] SignerTitle)
    {
        iTextSharp.text.Document document = null;
        string file = string.Empty;
        DataTable dtRedemptionGrid = null;
        DataTable dtRedemption = null;
        // bool bCrossFlg = false;
        try
        {
            dtRedemption = ds_gresham.Tables[0];
            if (ds_gresham.Tables.Count > 1)
            {
                dtRedemptionGrid = ds_gresham.Tables[1];
                //   bCrossFlg = true;
            }


            string CloseDate = Convert.ToString(dtRedemption.Rows[0]["CloseDate"]);
            string FundName = Convert.ToString(dtRedemption.Rows[0]["Fund"]);
            string LegalEntity = Convert.ToString(dtRedemption.Rows[0]["LegalEntity"]);
            LegalEntity = LegalEntity.Replace("&#39;", "'").ToString();
            string TransactionType = Convert.ToString(dtRedemption.Rows[0]["TransactionType"]);
            string Amount = Convert.ToString(dtRedemption.Rows[0]["Amount"]);
            if (Amount != "")
            {
                Amount = string.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(Amount));
            }
            string FurtherCreditTo = Convert.ToString(dtRedemption.Rows[0]["RetFurtherCredit"]);
            string Withdrawaltype = Convert.ToString(dtRedemption.Rows[0]["Withdrawaltype"]);
            string Custodian = Convert.ToString(dtRedemption.Rows[0]["Custodian"]);
            string CustodianAccNumber = Convert.ToString(dtRedemption.Rows[0]["CustodianAccNumber"]);
            // string WireInstrFlg = Convert.ToString(dtRedemption.Rows[0]["WireInstrFlg"]);
            string ABA = Convert.ToString(dtRedemption.Rows[0]["ABA"]);
            string CreditName = Convert.ToString(dtRedemption.Rows[0]["CreditName"]);
            string CreditAcct = Convert.ToString(dtRedemption.Rows[0]["CreditAcct"]);
            string InitialNotes = Convert.ToString(dtRedemption.Rows[0]["InitialNotes"]);

            #region Commented - 12_20_2018 Logic Change
            //string Signer1Name = Convert.ToString(dtRedemption.Rows[0]["Signer1Name"]);
            //string Signer1Title = Convert.ToString(dtRedemption.Rows[0]["Signer1Title"]);
            //string Signer2Name = Convert.ToString(dtRedemption.Rows[0]["Signer2Name"]);
            //string Signer2Title = Convert.ToString(dtRedemption.Rows[0]["Signer2Title"]);
            //string Signer3Name = Convert.ToString(dtRedemption.Rows[0]["Signer3Name"]);
            //string Signer3Title = Convert.ToString(dtRedemption.Rows[0]["Signer3Title"]);

            //string LE_Signer1Name = Convert.ToString(dtRedemption.Rows[0]["LE_Signer1Name"]);
            //string LE_Signer1Title = Convert.ToString(dtRedemption.Rows[0]["LE_Signer1Title"]);
            //string LE_Signer2Name = Convert.ToString(dtRedemption.Rows[0]["LE_Signer2Name"]);
            //string LE_Signer2Title = Convert.ToString(dtRedemption.Rows[0]["LE_Signer2Title"]);
            //string LE_Signer3Name = Convert.ToString(dtRedemption.Rows[0]["LE_Signer3Name"]);
            //string LE_Signer3Title = Convert.ToString(dtRedemption.Rows[0]["LE_Signer3Title"]);

            //string ShowAccSignorFlg = Convert.ToString(dtRedemption.Rows[0]["ShowAccSignorFlg"]);
            #endregion

            string CrossFlg = Convert.ToString(dtRedemption.Rows[0]["CrossFlg"]);
            string BasicWireInfo = Convert.ToString(dtRedemption.Rows[0]["BasicWireInfo"]);//added on 27_06_2018 by brijesh

            #region Commented - 12_20_2018 Logic Change
            //string[] SignerNames = new string[3];
            //string[] SignerTitle = new string[3];

            //string[] LE_SignerNames = new string[3];
            //string[] LE_SignerTitle = new string[3];

            //SignerNames[0] = Signer1Name.Replace("&nbsp;", "");
            //SignerNames[1] = Signer2Name.Replace("&nbsp;", "");
            //SignerNames[2] = Signer3Name.Replace("&nbsp;", "");

            //LE_SignerNames[0] = LE_Signer1Name.Replace("&nbsp;", "");
            //LE_SignerNames[1] = LE_Signer2Name.Replace("&nbsp;", "");
            //LE_SignerNames[2] = LE_Signer3Name.Replace("&nbsp;", "");

            //SignerTitle[0] = Signer1Title.Replace("&nbsp;", "");
            //SignerTitle[1] = Signer2Title.Replace("&nbsp;", "");
            //SignerTitle[2] = Signer3Title.Replace("&nbsp;", "");

            //LE_SignerTitle[0] = LE_Signer1Title.Replace("&nbsp;", "");
            //LE_SignerTitle[1] = LE_Signer2Title.Replace("&nbsp;", "");
            //LE_SignerTitle[2] = LE_Signer3Title.Replace("&nbsp;", "");
            #endregion




            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 5, 2, 30, 10);//10,10
            string FolderPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\";
            string fileName = System.DateTime.Now.ToString("MMddyyhhmmss") + "GPRedemeptionRequestForm.pdf";
            file = Path.Combine(FolderPath, fileName);
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(file, FileMode.Create));

            document.Open();
            PdfPCell loCell = new PdfPCell();
            Chunk imge = new Chunk();
            PdfPTable tblLogo = new PdfPTable(1);

            #region Logo
            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
            png.SetAbsolutePosition(65, 815);//540
            png.ScalePercent(8.75f);
            //document.Add(png);
            imge = new Chunk(png, 0, 0, true);
            loCell.Border = 0;
            loCell.AddElement(imge);
            tblLogo.AddCell(loCell);
            #endregion

            float FontSize = 12f;
            Chunk headerChunk = new Chunk("\nGP Redemption Request Form\n", setFontsAllVerdana(FontSize, 1, 0));
            Paragraph p1 = new Paragraph();
            p1.Add(headerChunk);
            p1.IndentationLeft = 60;

            PdfPTable tblHeader = new PdfPTable(2);

            tblHeader.TotalWidth = 100f;
            int[] headerwidths = { 50, 50 };
            tblHeader.SetWidths(headerwidths);

            FontSize = 9f;

            PdfPCell cell1 = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            cell1.Colspan = 2;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Withdrawal Date:	", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase(CloseDate, setFontsAllVerdana(FontSize, 0, 0)));
            //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
            cell1.Border = PdfCell.BOTTOM_BORDER;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("GP Partnership (Fund): ", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase(FundName, setFontsAllVerdana(FontSize, 0, 0)));
            cell1.Border = PdfCell.BOTTOM_BORDER;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Investor Name (Legal Entity):", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase(LegalEntity, setFontsAllVerdana(FontSize, 0, 0)));
            cell1.Border = PdfCell.BOTTOM_BORDER;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Transaction Type:", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase(TransactionType, setFontsAllVerdana(FontSize, 0, 0)));
            cell1.Border = PdfCell.BOTTOM_BORDER;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Withdrawal Amount:", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);
            if (Amount != "" && Amount != null)
            {
                cell1 = new PdfPCell(new Phrase("$ " + Amount, setFontsAllVerdana(FontSize, 0, 0)));
            }
            else
            {
                cell1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));

            }
            cell1.Border = PdfCell.BOTTOM_BORDER;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Redemption Provision*:", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase(Withdrawaltype, setFontsAllVerdana(FontSize, 0, 0)));
            cell1.Border = PdfCell.BOTTOM_BORDER;
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Reinvestment to other GP Partnership (Cross):", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.PaddingLeft = 7;
            cell1.PaddingTop = 5;
            //cell1.Colspan = 2;
            tblHeader.AddCell(cell1);

            if (CrossFlg.ToLower() == "true")
            {
                cell1 = new PdfPCell(new Phrase("Yes", setFontsAllVerdana(FontSize, 0, 0)));
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            }
            else
            {
                cell1 = new PdfPCell(new Phrase("No", setFontsAllVerdana(FontSize, 0, 0)));
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            }


            //cell1.Border = PdfCell.BOTTOM_BORDER;
            cell1.Border = 0;
            //cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            // cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.PaddingTop = 5;
            tblHeader.AddCell(cell1);


            #region RelatedPartnership GRID
            PdfPTable tblRelatedPartnership = new PdfPTable(3);
            tblRelatedPartnership.TotalWidth = 100f;
            float[] headerwidths1 = { 1.5f, 48, 50 };
            tblRelatedPartnership.SetWidths(headerwidths1);

            if (CrossFlg.ToLower() == "true")
            {
                cell1 = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                cell1.Border = 0;
                cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                cell1.PaddingTop = 5;
                cell1.Colspan = 2;
                tblHeader.AddCell(cell1);

                for (int i = 0; i < dtRedemptionGrid.Rows.Count; i++)
                {
                    string Name = Convert.ToString(dtRedemptionGrid.Rows[i]["Name"]);
                    string Amount1 = Convert.ToString(dtRedemptionGrid.Rows[i]["Amount"]);
                    Amount1 = string.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(Amount1));


                    cell1 = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell1.Border = 0;
                    // cell1.PaddingTop = 5;
                    tblRelatedPartnership.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase(Name, setFontsAllVerdana(FontSize, 0, 0)));
                    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell1.PaddingLeft = 28;
                    //cell1.PaddingTop = 5;
                    cell1.SetLeading(1, 1);
                    tblRelatedPartnership.AddCell(cell1);

                    cell1 = new PdfPCell(new Phrase("$ " + Amount1, setFontsAllVerdana(FontSize, 0, 0)));
                    cell1.HorizontalAlignment = Element.ALIGN_LEFT;
                    // cell1.PaddingTop = 5;
                    tblRelatedPartnership.AddCell(cell1);
                }
                #endregion
            }


            #region Wire Instructions
            #region oldCommented
            //  PdfPTable tblWireInstructions = new PdfPTable(4);
            //  tblWireInstructions.TotalWidth = 100f;
            ////  int[] headerwidths4 = { 25, 75 };
            //  //int[] headerwidths4 = { 1, 55, 44 };
            //  //int[] headerwidths4 = { 1, 24, 75 };
            //  int[] headerwidths4 = { 1, 24, 30,45 };
            //  tblWireInstructions.SetWidths(headerwidths4);




            //            PdfPCell cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //            cellInstruction.Border = 0;
            //           // cellInstruction.Colspan = 2;
            //            cellInstruction.Colspan =4;
            //            tblWireInstructions.AddCell(cellInstruction);

            //            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.Border = 0;
            //            tblWireInstructions.AddCell(cellInstruction);

            //            cellInstruction = new PdfPCell(new Phrase("Wire Instructions", setFontsAllVerdana(FontSize, 1, 0)));
            //            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //            cellInstruction.Border = 0;
            //            //cellInstruction.PaddingLeft = 7;
            //            cellInstruction.Colspan = 3;
            //            tblWireInstructions.AddCell(cellInstruction);

            //            if (WireInstrFlg == "1" || WireInstrFlg == "2")
            //            {
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //#endregion
            //                if (BasicWireInfo != "")
            //                {
            //                    cellInstruction = new PdfPCell(new Phrase("" + BasicWireInfo, setFontsAllVerdana(FontSize, 0, 0)));
            //                    //cellInstruction = new PdfPCell(new Phrase("" + BasicWireInfo, setUnderline(FontSize,1)));
            //                    cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                    cellInstruction.Border = 0;
            //                    //cellInstruction.PaddingLeft = 7;
            //                    cellInstruction.SetLeading(15, 0);
            //                    cellInstruction.Colspan = 3;
            //                    tblWireInstructions.AddCell(cellInstruction);
            //                }
            //                else
            //                {
            //                    cellInstruction = new PdfPCell(new Phrase("" , setFontsAllVerdana(FontSize, 0, 0)));
            //                    //cellInstruction = new PdfPCell(new Phrase("" + BasicWireInfo, setUnderline(FontSize,1)));
            //                    cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                    cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                    //cellInstruction.PaddingLeft = 7;
            //                    cellInstruction.SetLeading(15, 0);
            //                    cellInstruction.Colspan = 3;
            //                    tblWireInstructions.AddCell(cellInstruction);


            //                    #region emptycell
            //                    cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                    cellInstruction.Border = 0;
            //                    tblWireInstructions.AddCell(cellInstruction);
            //                    #endregion

            //                    cellInstruction = new PdfPCell(new Phrase("" + BasicWireInfo, setFontsAllVerdana(FontSize, 0, 0)));
            //                    //cellInstruction = new PdfPCell(new Phrase("" + BasicWireInfo, setUnderline(FontSize,1)));
            //                    cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                    cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                    //cellInstruction.PaddingLeft = 7;
            //                    cellInstruction.SetLeading(15, 0);
            //                    cellInstruction.Colspan = 3;
            //                    tblWireInstructions.AddCell(cellInstruction);

            //                    #region emptycell
            //                    cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                    cellInstruction.Border = 0;
            //                    tblWireInstructions.AddCell(cellInstruction);
            //                    #endregion

            //                    cellInstruction = new PdfPCell(new Phrase("" + BasicWireInfo, setFontsAllVerdana(FontSize, 0, 0)));
            //                    //cellInstruction = new PdfPCell(new Phrase("" + BasicWireInfo, setUnderline(FontSize,1)));
            //                    cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                    cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                    //cellInstruction.PaddingLeft = 7;
            //                    cellInstruction.SetLeading(15, 0);
            //                    cellInstruction.Colspan = 3;
            //                    tblWireInstructions.AddCell(cellInstruction);
            //                } 

            //                #region ABA
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion

            //                cellInstruction = new PdfPCell(new Phrase("ABA: " , setFontsAllVerdana(FontSize, 0, 0)));                               
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                cellInstruction.Border = 0;                
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //               // cellInstruction.Colspan = 2;
            //                tblWireInstructions.AddCell(cellInstruction);

            //                cellInstruction = new PdfPCell(new Phrase( ABA, setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                //cellInstruction.Border = 0;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                cellInstruction.Colspan = 2;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion

            //                #region rowAccount Name
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                cellInstruction = new PdfPCell(new Phrase("Account Name: ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                //cellInstruction.Colspan = 2;
            //                tblWireInstructions.AddCell(cellInstruction);

            //                cellInstruction = new PdfPCell(new Phrase( CreditName, setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.SetLeading(12, 0);              
            //                tblWireInstructions.AddCell(cellInstruction);


            //#endregion
            //                #region Account#
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                cellInstruction = new PdfPCell(new Phrase("Acct#: ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                //cellInstruction.Colspan = 2;
            //                tblWireInstructions.AddCell(cellInstruction);

            //                cellInstruction = new PdfPCell(new Phrase( CreditAcct, setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.SetLeading(12, 0);               
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //            }
            //            if (WireInstrFlg == "2")
            //            {
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //#endregion
            //                cellInstruction = new PdfPCell(new Phrase("Further Credit to: ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                //cellInstruction.Colspan = 2;
            //                tblWireInstructions.AddCell(cellInstruction);

            //                cellInstruction = new PdfPCell(new Phrase("A/C #930-086957 ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.SetLeading(12, 0);               
            //                tblWireInstructions.AddCell(cellInstruction);
            //            }
            //            if (WireInstrFlg == "3")
            //            {
            //                //cellInstruction = new PdfPCell(new Phrase("\n\n\n\n" + CreditAcct, setFontsAllVerdana(FontSize, 0, 0)));
            //                //cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                //cellInstruction.Border = 0;
            //                //cellInstruction.Colspan = 2;
            //                //tblWireInstructions.AddCell(cellInstruction);
            //                #region row1
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion

            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                //cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                tblWireInstructions.AddCell(cellInstruction);

            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                #endregion
            //                #region row2
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                //cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                #endregion
            //                #region row3
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                //cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                #endregion
            //                #region row4
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion

            //                cellInstruction = new PdfPCell(new Phrase("ABA:", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                //cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                tblWireInstructions.AddCell(cellInstruction);

            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                #endregion
            //                #region row5
            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion

            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //                //cellInstruction.Border = 0;
            //                //cellInstruction.PaddingLeft = 7;
            //                cellInstruction.SetLeading(12, 0);
            //                cellInstruction.Colspan = 2;
            //                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //                tblWireInstructions.AddCell(cellInstruction);

            //                #region emptycell
            //                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //                cellInstruction.Border = 0;
            //                tblWireInstructions.AddCell(cellInstruction);
            //                #endregion
            //                #endregion
            //            }
            //#region row for Benefit of
            //            #region emptycell
            //            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.Border = 0;
            //            tblWireInstructions.AddCell(cellInstruction);
            //            #endregion
            //            cellInstruction = new PdfPCell(new Phrase("For Benefit of: " , setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //            cellInstruction.Border = 0;
            //            cellInstruction.SetLeading(12, 0);
            //            //cellInstruction.PaddingLeft = 7;
            //            //cellInstruction.Colspan = 2;
            //            tblWireInstructions.AddCell(cellInstruction);

            //            cellInstruction = new PdfPCell(new Phrase( LegalEntity, setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //            cellInstruction.SetLeading(12, 0);
            //            cellInstruction.Colspan = 2;     
            //            tblWireInstructions.AddCell(cellInstruction);
            //#endregion
            //            #region emptycell
            //            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.Border = 0;
            //            tblWireInstructions.AddCell(cellInstruction);
            //            #endregion

            //            #region row Account Number
            //            cellInstruction = new PdfPCell(new Phrase("Account Number: ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //            cellInstruction.Border = 0;
            //            cellInstruction.SetLeading(12, 0);
            //            //cellInstruction.PaddingLeft = 7;
            //            //cellInstruction.Colspan = 2;
            //            tblWireInstructions.AddCell(cellInstruction);

            //            cellInstruction = new PdfPCell(new Phrase( CustodianAccNumber, setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            //            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            //            cellInstruction.SetLeading(12, 0);
            //            cellInstruction.Colspan = 2;
            //            //cellInstruction.Colspan = 2;
            //            tblWireInstructions.AddCell(cellInstruction);
            //            #endregion
            //            #region emptycell
            //            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellInstruction.Border = 0;
            //            tblWireInstructions.AddCell(cellInstruction);
            //            #endregion

            //            #region Additional Instructions
            //            PdfPTable tblAdditionalInstructions = new PdfPTable(3);
            //            tblAdditionalInstructions.TotalWidth = 100f;
            //            //  int[] headerwidths4 = { 25, 75 };
            //            int[] headerwidths5 = { 1, 24, 75 };
            //            tblAdditionalInstructions.SetWidths(headerwidths5);

            //            #region emptycell
            //            PdfPCell cellAdditionalInstructions = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellAdditionalInstructions.Border = 0;
            //            tblAdditionalInstructions.AddCell(cellAdditionalInstructions);
            //            #endregion

            //            cellAdditionalInstructions = new PdfPCell(new Phrase("Additional Instructions: ", setFontsAllVerdana(FontSize, 0, 0)));
            //            cellAdditionalInstructions.HorizontalAlignment = Element.ALIGN_LEFT;
            //            cellAdditionalInstructions.Border = 0;
            //            cellAdditionalInstructions.SetLeading(12, 0);
            //            //cellInstruction.PaddingLeft = 7;
            //            tblAdditionalInstructions.AddCell(cellAdditionalInstructions);                 

            //           // cellAdditionalInstructions = new PdfPCell(new Phrase(InitialNotes, setFontsAllVerdana(FontSize, 0, 0)));
            //            cellAdditionalInstructions = new PdfPCell(new Phrase(InitialNotes,setUnderline(FontSize,1)));
            //            cellAdditionalInstructions.HorizontalAlignment = Element.ALIGN_LEFT;
            //            //cellAdditionalInstructions.Border = PdfCell.BOTTOM_BORDER;
            //            cellAdditionalInstructions.Border = 0;
            //            cellAdditionalInstructions.SetLeading(15, 0);
            //            //cellInstruction.SetLeading(13, 0);
            //            // cellInstruction.PaddingLeft = 7;
            //            tblAdditionalInstructions.AddCell(cellAdditionalInstructions);
            //            #endregion
            #endregion

            PdfPTable tblWireInstructions = new PdfPTable(3);
            tblWireInstructions.TotalWidth = 100f;
            float[] headerwidths4 = { 1.5f, 30, 69 };
            tblWireInstructions.SetWidths(headerwidths4);

            PdfPCell cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            // cellInstruction.Colspan = 2;
            cellInstruction.Colspan = 3;
            tblWireInstructions.AddCell(cellInstruction);

            # region Row Wire Instruction Heading
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            // cellInstruction.Colspan = 2;
            //cellInstruction.Colspan = 4;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("Wire Instructions", setFontsAllVerdana(FontSize, 1, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            //cellInstruction.PaddingLeft = 7;
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion

            #region BlankRow
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            // cellInstruction.Colspan = 2;
            cellInstruction.Colspan = 3;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion

            //if (WireInstrFlg == "1" || WireInstrFlg == "2")
            //{
            #region Basic Instruction Box
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("Bank Info:", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.TOP_BORDER | PdfCell.LEFT_BORDER | PdfCell.RIGHT_BORDER;
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            if (BasicWireInfo != "")
            {
                cellInstruction = new PdfPCell(new Phrase(BasicWireInfo, setFontsAllVerdana(FontSize, 0, 0)));
            }
            else
            {
                cellInstruction = new PdfPCell(new Phrase("\n\n\n\n\n", setFontsAllVerdana(FontSize, 0, 0)));
            }

            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.LEFT_BORDER | PdfCell.RIGHT_BORDER | PdfCell.BOTTOM_BORDER;
            //cellInstruction.SetLeading(12, 0);
            cellInstruction.PaddingBottom = 3;
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion


            #region row ABA
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("ABA: ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            cellInstruction.SetLeading(12, 0);
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase(ABA, setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            cellInstruction.SetLeading(12, 0);
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion

            #region row Account Name
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("Account Name: ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            cellInstruction.SetLeading(12, 0);
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase(CreditName, setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            cellInstruction.SetLeading(12, 0);
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion

            #region row Acct#
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("Acct#: ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            cellInstruction.SetLeading(12, 0);
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase(CreditAcct, setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            cellInstruction.SetLeading(12, 0);
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion
            //}
            //if (WireInstrFlg == "2")
            //{
            #region row Acct#
            if (FurtherCreditTo != "")
            {
                cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                cellInstruction.Border = 0;
                tblWireInstructions.AddCell(cellInstruction);

                cellInstruction = new PdfPCell(new Phrase("Further Credit to: ", setFontsAllVerdana(FontSize, 0, 0)));
                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
                cellInstruction.Border = 0;
                cellInstruction.SetLeading(12, 0);
                tblWireInstructions.AddCell(cellInstruction);

                cellInstruction = new PdfPCell(new Phrase(FurtherCreditTo, setFontsAllVerdana(FontSize, 0, 0)));
                cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
                cellInstruction.Border = PdfCell.BOTTOM_BORDER;
                cellInstruction.SetLeading(12, 0);
                cellInstruction.Colspan = 2;
                tblWireInstructions.AddCell(cellInstruction);
            }
            #endregion
            //}
            //if (WireInstrFlg == "3")
            //{

            //}

            #region row For Benefit of
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("For Benefit of: ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            cellInstruction.PaddingLeft = 20;
            cellInstruction.SetLeading(12, 0);
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase(LegalEntity, setFontsAllVerdana(FontSize, 1, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            cellInstruction.SetLeading(12, 0);
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion

            #region row Account Number
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("Account Number: ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            cellInstruction.PaddingLeft = 20;
            cellInstruction.SetLeading(12, 0);
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase(CustodianAccNumber, setFontsAllVerdana(FontSize, 1, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            cellInstruction.SetLeading(12, 0);
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion

            #region row Additional Instructions:
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.Border = 0;
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase("Additional Instructions: ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            cellInstruction.PaddingLeft = 20;
            cellInstruction.SetLeading(12, 0);
            tblWireInstructions.AddCell(cellInstruction);

            cellInstruction = new PdfPCell(new Phrase(InitialNotes, setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = PdfCell.BOTTOM_BORDER;
            cellInstruction.SetLeading(12, 0);
            cellInstruction.Colspan = 2;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion
            #region BlankRow
            cellInstruction = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellInstruction.HorizontalAlignment = Element.ALIGN_LEFT;
            cellInstruction.Border = 0;
            // cellInstruction.Colspan = 2;
            cellInstruction.Colspan = 3;
            tblWireInstructions.AddCell(cellInstruction);
            #endregion
            #endregion

            #region Table Signer
            # region Comment
            PdfPTable tblEntity = new PdfPTable(2);
            tblEntity.TotalWidth = 100f;
            int[] headerwidths3 = { 50, 50 };
            tblEntity.SetWidths(headerwidths3);

            cell1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.Colspan = 2;
            tblEntity.AddCell(cell1);

            //PdfPCell cellhead = new PdfPCell(new Phrase("Entity:", setFontsAllVerdana(FontSize, 0, 0)));
            //cellhead.HorizontalAlignment = Element.ALIGN_LEFT;
            //cellhead.Border = 0;
            //cellhead.PaddingLeft = 7;
            //tblEntity.AddCell(cellhead);

            //PdfPCell cellhead1 = new PdfPCell(new Phrase(LegalEntity, setFontsAllVerdana(FontSize, 0, 0)));
            //cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            //cellhead1.Border = PdfCell.BOTTOM_BORDER;
            //tblEntity.AddCell(cellhead1);

            //cellhead1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            //cellhead1.Border = 0;
            ////cellhead1.Colspan = 3;
            //tblEntity.AddCell(cellhead1);

            //cellhead1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            //cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            //cellhead1.Border = 0;
            ////cellhead1.Colspan = 3;
            //tblEntity.AddCell(cellhead1);
            #endregion

            PdfPTable tblSigner1 = new PdfPTable(4);
            tblSigner1.TotalWidth = 100f;
            //  int[] headerwidths2 = { 50, 50 };
            int[] headerwidths2 = { 25, 50, 10, 15 };
            tblSigner1.SetWidths(headerwidths2);
            string signer = string.Empty;
            string Title = string.Empty;
            //string LE_signer = string.Empty;
            //string LE_title = string.Empty;
            int countSigner = 0;
            for (int i = 0; i < 3; i++)
            {
                #region Commented - 12_20_2018 Logic Change
                //if (ShowAccSignorFlg == "0")
                //{
                //    signer = LE_SignerNames[i].ToString();
                //    Title = LE_SignerTitle[i].ToString();
                //}
                //else
                //{
                //    signer = SignerNames[i].ToString();
                //    Title = SignerTitle[i].ToString();
                //}
                #endregion
                signer = SignerNames[i].ToString();
                Title = SignerTitle[i].ToString();
                if (signer != "")
                {

                    countSigner++;
                    #region row Authorised Signature
                    // iTextSharp.text.Font fontTable = FontFactory.GetFont("Arial", FontSize, iTextSharp.text.Font.UNDERLINE);
                    PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    cell.PaddingLeft = 7;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                    cell.Border = PdfCell.BOTTOM_BORDER;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.PaddingLeft = 7;
                    //cell.Border = 0;            
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Dated:", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    cell.PaddingLeft = 7;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                    cell.Border = PdfCell.BOTTOM_BORDER;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.PaddingLeft = 7;
                    //cell.Border = 0;            
                    tblSigner1.AddCell(cell);

                    #endregion
                    #region row PrintName
                    cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    cell.PaddingTop = 5;
                    cell.PaddingLeft = 7;
                    tblSigner1.AddCell(cell);

                    if (Title != "")
                    {
                        cell = new PdfPCell(new Phrase(signer + ", " + Title, setFontsAllVerdana(FontSize, 0, 0)));
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase(signer, setFontsAllVerdana(FontSize, 0, 0)));
                    }
                    cell.PaddingTop = 5;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = PdfCell.BOTTOM_BORDER;
                    //cell.Colspan = 2;
                    tblSigner1.AddCell(cell);
                    #endregion

                    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    cell.Colspan = 2;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    cell.Colspan = 4;
                    tblSigner1.AddCell(cell);

                }




            }
            if (countSigner == 0)
            {
                PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                cell.PaddingLeft = 7;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                cell.Border = PdfCell.BOTTOM_BORDER;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.Border = 0;            
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("Dated:", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                cell.PaddingLeft = 7;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                cell.Border = PdfCell.BOTTOM_BORDER;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.PaddingLeft = 7;
                //cell.Border = 0;            
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                cell.PaddingTop = 5;
                cell.PaddingLeft = 7;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                cell.Border = PdfCell.BOTTOM_BORDER;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.PaddingTop = 5;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                cell.Border = 0;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Colspan = 2;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                cell.Colspan = 4;
                tblSigner1.AddCell(cell);

            }
            #endregion

            #region Footer Note
            PdfPTable tblFooter = new PdfPTable(2);
            tblFooter.TotalWidth = 100f;
            int[] headerwidths12 = { 50, 50 };
            tblFooter.SetWidths(headerwidths12);
            FontSize = 8;

            Chunk chunk1 = new Chunk("\n*Full withdrawals will be paid out over two quarters 50% on 1st quarter end date and the remaining \namount on ", setFontsAllVerdana(FontSize, 0, 0));
            Chunk chunk2 = new Chunk("2nd quarter end date excluding the audit reserve. ", setFontsAllVerdana(FontSize, 0, 1));
            Chunk chunk3 = new Chunk("Investors requesting specific dollar \nwithdrawals are subject to withdrawal limits for a given period. Investors may not receive the exact \ndollar requested as market fluctuations could result in a lower amount being paid based on the \npartnership value as of the withdrawal date.", setFontsAllVerdana(FontSize, 0, 0));

            Phrase phrase1 = new Phrase();
            phrase1.Add(chunk1);
            phrase1.Add(chunk2);
            phrase1.Add(chunk3);

            Paragraph para1 = new Paragraph(10);
            para1.Add(phrase1);
            cell1.AddElement(para1);
            //   cell1 = new PdfPCell(new Phrase("\n*Full withdrawals will be paid out over two quarters 50% on 1st quarter end date and the remaining \namount on 2nd quarter end date excluding the audit reserve. Investors requesting specific dollar \nwithdrawals are subject to withdrawal limits for a given period. Investors may not receive the exact \ndollar requested as market fluctuations could result in a lower amount being paid based on the \npartnership value as of the withdrawal date.", setFontsAllVerdana(FontSize, 0, 0)));

            cell1.HorizontalAlignment = Element.ALIGN_LEFT;
            cell1.Border = 0;
            cell1.Colspan = 2;
            //cell1.SetLeading(25, 3);
            tblFooter.AddCell(cell1);


            iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\AdvisorInitial.png");
            dashjpg.ScalePercent(35);
            cell1 = new PdfPCell(dashjpg);
            cell1.Colspan = 2;
            cell1.Border = 0;
            cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
            tblFooter.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Advisor Initials", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
            //cell1.PaddingRight = 10;
            cell1.Border = 0;
            cell1.Colspan = 2;
            tblFooter.AddCell(cell1);
            #endregion

            #region Footer

            string Address = "Gresham Partners LLC    333 W Wacker Drive, Suite 700 Chicago, IL 60606    (312) 960-0200  Fax (312) 960-0204";
            PdfPTable TabFooter1 = addFooterAddress(Address, 1, 1);
            TabFooter1.HorizontalAlignment = Element.ALIGN_CENTER;
            TabFooter1.WidthPercentage = 100f;
            //  TabFooter.TotalWidth = 100f;
            TabFooter1.TotalWidth = 600;

            TabFooter1.WriteSelectedRows(0, 4, 0, 38, writer.DirectContent);
            #endregion
            document.Add(tblLogo);
            document.Add(p1);
            document.Add(tblHeader);
            if (CrossFlg.ToLower() == "true")
            {
                document.Add(tblRelatedPartnership);
            }


            document.Add(tblWireInstructions);
            // document.Add(tblAdditionalInstructions);
            document.Add(tblEntity);
            document.Add(tblSigner1);
            document.Add(tblFooter);
            document.Close();
        }
        catch (Exception ex)
        {
            file = "";
            Response.Write("ERROR :" + ex.Message.ToString());
            lblError.Visible = true;
            lblError.Text = "ERROR: " + ex.Message.ToString();
        }
        finally
        {
            document.Close();

        }

        return file;
    }
    public string SignerPage(string LegalEntity, string[] SignerNames, string[] SignerTitle, string doctype, string CloseDate)
    {
        iTextSharp.text.Document document = null;
        string file = string.Empty;
        try
        {
            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 5, 2, 20, 10);//10,10
            string FolderPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\";
            string fileName = System.DateTime.Now.ToString("MMddyyhhmmss") + "SignerPAge.pdf";
            file = Path.Combine(FolderPath, fileName);
            PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(file, FileMode.Create));

            // string lsFooterText = FooterText;//footer text is in below method
            document.Open();

            PdfContentByte cb = writer.DirectContent;

            PdfPTable loHeader = new PdfPTable(2);
            loHeader.TotalWidth = 100f;
            int[] headerwidths111 = { 30, 70 };
            loHeader.SetWidths(headerwidths111);


            #region Header
            Paragraph lochunk = new Paragraph();
            PdfPCell loCell = new PdfPCell();

            float FontSize = 9f;
            string ReportHeader1 = "Close Date: " + CloseDate;
            lochunk = new Paragraph(ReportHeader1, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_RIGHT;// SetAlignment("center");
            // lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);
            loCell.Colspan = 2;
            loCell.Border = 0;
            loHeader.AddCell(loCell);

            loCell = new PdfPCell();
            string ReportHeader = "\nTHE UNDERSIGNED AGREES TO NOTIFY THE GENERAL PARTNER\nPROMPTLY IN WRITING SHOULD THERE BE ANY CHANGE\nIN ANY OF THE FOREGOING INFORMATION, REPRESENTATIONS OR WARRANTIES.";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            // lochunk.SetLeading(13, 0);
            loCell.AddElement(lochunk);
            loCell.Colspan = 2;
            loCell.Border = 0;

            loHeader.AddCell(loCell);

            loCell = new PdfPCell();
            ReportHeader = "\n";
            lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            loCell.AddElement(lochunk);
            loCell.Colspan = 2;
            loCell.Border = 0;
            loHeader.AddCell(loCell);

            //loCell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
            //loCell.Border = PdfCell.BOTTOM_BORDER;
            //loCell.HorizontalAlignment = Element.ALIGN_LEFT;
            //loHeader.AddCell(loCell);

            loCell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
            loCell.HorizontalAlignment = Element.ALIGN_LEFT;
            loCell.Colspan = 2;
            loCell.Border = 0;
            loHeader.AddCell(loCell);

            //ReportHeader = "\n\nEntity:";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);
            //loHeader.AddCell(loCell);

            //ReportHeader = "Entity Value from DATABSAE";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);
            //loHeader.AddCell(loCell);


            PdfPTable tblEntity = new PdfPTable(2);
            tblEntity.TotalWidth = 100f;
            int[] headerwidths = { 30, 70 };
            tblEntity.SetWidths(headerwidths);

            PdfPCell cellhead = new PdfPCell(new Phrase("Dated:", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead.Border = 0;
            //cellhead.PaddingLeft = 7;
            tblEntity.AddCell(cellhead);

            PdfPCell cellhead1 = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = PdfCell.BOTTOM_BORDER;
            tblEntity.AddCell(cellhead1);


            cellhead = new PdfPCell(new Phrase("Investor/Entity:", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead.Border = 0;
            cellhead.PaddingTop = 5;
            //cellhead.PaddingLeft = 7;
            tblEntity.AddCell(cellhead);

            cellhead1 = new PdfPCell(new Phrase(LegalEntity, setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = PdfCell.BOTTOM_BORDER;
            cellhead1.PaddingTop = 5;
            tblEntity.AddCell(cellhead1);

            cellhead1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = 0;
            tblEntity.AddCell(cellhead1);

            cellhead1 = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
            cellhead1.HorizontalAlignment = Element.ALIGN_LEFT;
            cellhead1.Border = 0;
            tblEntity.AddCell(cellhead1);

            #region Table Signer
            PdfPTable tblSigner1 = new PdfPTable(2);
            tblSigner1.TotalWidth = 100f;
            int[] headerwidths1 = { 30, 70 };
            tblSigner1.SetWidths(headerwidths1);

            int countSigner = 0;
            for (int i = 0; i < 3; i++)
            {
                string signer = SignerNames[i].ToString();
                string Title = SignerTitle[i].ToString();
                if (signer != "")
                {
                    countSigner++;
                    // iTextSharp.text.Font fontTable = FontFactory.GetFont("Arial", FontSize, iTextSharp.text.Font.UNDERLINE);
                    PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    //cell.PaddingLeft = 7;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                    //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                    cell.Border = PdfCell.BOTTOM_BORDER;
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;            
                    tblSigner1.AddCell(cell);

                    //cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;
                    //tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    cell.PaddingTop = 5;
                    //cell.PaddingLeft = 7;
                    tblSigner1.AddCell(cell);

                    if (Title != "")
                    {
                        cell = new PdfPCell(new Phrase(signer + ", " + Title, setFontsAllVerdana(FontSize, 0, 0)));
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase(signer, setFontsAllVerdana(FontSize, 0, 0)));
                    }
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = PdfCell.BOTTOM_BORDER;
                    //cell.Colspan = 2;
                    cell.PaddingTop = 5;
                    tblSigner1.AddCell(cell);

                    #region Blank Row
                    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    tblSigner1.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    cell.Border = 0;
                    tblSigner1.AddCell(cell);
                    #endregion
                    //cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;
                    //tblSigner1.AddCell(cell);

                    //cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                    //cell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //cell.Border = 0;
                    //tblSigner1.AddCell(cell);

                }




            }
            if (countSigner == 0)
            {
                PdfPCell cell = new PdfPCell(new Phrase("Authorized Signature:", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                //cell.PaddingLeft = 7;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                cell.Border = PdfCell.BOTTOM_BORDER;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.Border = 0;            
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("Print Name:", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                cell.PaddingTop = 5;
                //cell.PaddingLeft = 7;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
                //  cell = new PdfPCell(new Phrase("__________________",  setFontsAll(11, 1, 0, new iTextSharp.text.Color(255, 255, 255))));
                cell.Border = PdfCell.BOTTOM_BORDER;
                cell.PaddingTop = 5;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                //cell.Border = 0;            
                tblSigner1.AddCell(cell);

                #region Blank Row
                cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                tblSigner1.AddCell(cell);

                cell = new PdfPCell(new Phrase(" ", setFontsAllVerdana(FontSize, 0, 0)));
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = 0;
                tblSigner1.AddCell(cell);
                #endregion
            }
            #endregion


            #region unused
            //ReportHeader = "Authorized Signature:";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell1.AddElement(lochunk);

            //ReportHeader = "Print Name:";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell1.AddElement(lochunk);

            //FontSize = 8.5f;
            //ReportHeader = "\n\nGresham Private Equity Strategies, L.P. \nc/o Gresham Advisors, L.L.C.\n333 West Wacker Drive\nSuite 700 \nChicago, Illinois 60606 \n \nDear Sir or Madam:";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            //ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Private Equity Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP 2018 Private Equity Strategies (the “Series”) in the amount of:  $ [insert isnull(confirmed amount, proposed amount)].";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            //ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the 2011/2012 Series, 2013 Series, 2014 Series, 2015 Series, 2016 Series and/or 2017 Series and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein. ";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            //ReportHeader = "\nThe undersigned hereby (in addition and not by way of limitation of the power of attorney as set forth in the Partnership Agreement) irrevocably constitutes and appoints the General Partner, its successors and assigns, and the officers of the foregoing, as the undersigned’s true and lawful Attorney-in-Fact, with full power of substitution, in the undersigned’s name, place and stead, to: (a) file, prosecute, defend, settle or compromise litigation, claims or arbitrations on behalf of the Series and/or the Partnership; (b) make, execute, sign, acknowledge, swear to, deliver, record and file any documents or instruments, including, without limitation, Certificates of Limited Partnership and amendments thereto, the Partnership Agreement and amendments thereto, that may be considered necessary or desirable by the General Partner to carry out fully the provisions of the Partnership Agreement, including, without limitation, those (if any) necessary or desirable to effect the undersigned’s admission as a Limited Partner; and (c) to perform all other acts contemplated by the Partnership Agreement.  This Power of Attorney shall be deemed to be coupled with an interest and shall be irrevocable and survive and not be affected by the undersigned’s subsequent death, incapacity, disability, insolvency or dissolution or any delivery by the undersigned of an assignment of the whole or any portion of the undersigned’s Interest.\nThis Agreement shall be governed in accordance with the laws of the State of Delaware (without regard to conflict of law principles).";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            //ReportHeader = "\n(signature page to follow)";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            //lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            //ReportHeader = "IN ANY OF THE FOREGOING INFORMATION, REPRESENTATIONS OR WARRANTIES";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 1, 0));
            //lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            //loCell.AddElement(lochunk);


            //lochunk = new Paragraph(" ", setFontsAllVerdana(FontSize, 1, 0));
            //lochunk.Alignment = Element.ALIGN_CENTER;// SetAlignment("center");
            //loCell.AddElement(lochunk);
            // loHeader.AddCell(loCell);
            // loCell.Border = 1;



            #endregion
            //  PdfPTable tblFooter = new PdfPTable(1);
            //   Paragraph lochunk1 = new Paragraph();
            //   PdfPCell loCell1 = new PdfPCell();

            //string  FooterHeader = "----------------------------------------------------------------------------";
            //lochunk1 = new Paragraph(FooterHeader, setFontsAllVerdana(FontSize, 1, 0));
            //lochunk1.Alignment = Element.ALIGN_BOTTOM;// SetAlignment("center");
            //loCell1.AddElement(lochunk1);
            //loCell1.Border = 0;

            #region intenalUse
            FontSize = 8;
            PdfPTable tblIntenal = new PdfPTable(2);

            tblIntenal.TotalWidth = 100f;
            int[] headerwidths5 = { 82, 18 };
            tblIntenal.SetWidths(headerwidths5);

            // iTextSharp.text.Font fontTable = FontFactory.GetFont("Arial", FontSize, iTextSharp.text.Font.UNDERLINE);
            PdfPCell cell1 = new PdfPCell(new Phrase("\n\n--------------------------------------------------------------------------------------------------------------------------------", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            FontSize = 10;
            cell1 = new PdfPCell(new Phrase("FOR INTERNAL USE ONLY\nTo be completed by Gresham Advisors, L.L.C.", setFontsAllVerdana(FontSize, 1, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            FontSize = 9;
            cell1 = new PdfPCell(new Phrase("COMMITTMENT ACCEPTED\nAS TO $___________________________", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            if (doctype.ToLower() == "gpes")
            {
                cell1 = new PdfPCell(new Phrase("\nGresham Private Equity Strategies, L.P.", setFontsAllVerdana(10, 1, 0)));
            }
            else if (doctype.ToLower() == "gras")
            {
                cell1 = new PdfPCell(new Phrase("\nGresham Real Asset Strategies, L.P.", setFontsAllVerdana(10, 1, 0)));
            }
            // cell1 = new PdfPCell(new Phrase("\nGresham Private Equity Strategies, L.P.", setFontsAllVerdana(10, 1, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            cell1.Colspan = 2;
            tblIntenal.AddCell(cell1);

            //cell1 = new PdfPCell(new Phrase("By:  Gresham Advisors, L.L.C.\n\nBy:  _______________________\nDate: __________________, 20____", setFontsAllVerdana(FontSize, 0, 0)));
            cell1 = new PdfPCell(new Phrase("By:  Gresham Advisors, L.L.C.\n\nBy:  _______________________\nDate: _______________________", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            cell1.PaddingLeft = 85;
            tblIntenal.AddCell(cell1);


            iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\AdvisorInitial.png");
            dashjpg.ScalePercent(35);
            cell1 = new PdfPCell(dashjpg);
            cell1.Border = 0;
            cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
            tblIntenal.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.Border = 0;
            cell1.SetLeading(12, 0);
            tblIntenal.AddCell(cell1);

            cell1 = new PdfPCell(new Phrase("Advisor Initials", setFontsAllVerdana(FontSize, 0, 0)));
            cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
            cell1.Border = 0;
            //cell1.SetLeading(12, 0);
            tblIntenal.AddCell(cell1);

            #endregion
            #region Footer
            //string line1 = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------\nFOR INTERNAL USE ONLY\nTo be completed by Gresham Advisors, L.L.C.";
            //string line2 = "\nCOMMITTMENT ACCEPTED\nAS TO $___________________________";
            //string line3 = "Gresham Private Equity Strategies, L.P.";
            //string line4 = "\nBy:  Gresham Advisors, L.L.C.\nBy:  _______________________\nDate: __________________, 20____";
            //string line5 = "Advisor Initials";
            //PdfPTable TabFooter = addFooterInternal(line1, line2, line3, line4, line5, true);
            //TabFooter.HorizontalAlignment = Element.ALIGN_CENTER;
            //TabFooter.WidthPercentage = 100f;
            ////  TabFooter.TotalWidth = 100f;
            //TabFooter.TotalWidth = 600;
            //TabFooter.WriteSelectedRows(0, 4, 0, 190, writer.DirectContent);

            string Address = "Gresham Partners LLC    333 W Wacker Drive, Suite 700 Chicago, IL 60606    (312) 960-0200  Fax (312) 960-0204";
            PdfPTable TabFooter1 = addFooterAddress(Address, 2, 2);
            TabFooter1.HorizontalAlignment = Element.ALIGN_CENTER;
            TabFooter1.WidthPercentage = 100f;
            //  TabFooter.TotalWidth = 100f;
            TabFooter1.TotalWidth = 600;

            TabFooter1.WriteSelectedRows(0, 4, 0, 38, writer.DirectContent);
            #endregion
            #region Advisor Initial Box
            //iTextSharp.text.Image dashjpg = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\AdvisorInitial.png");
            ////dashjpg.ScaleToFit(50f, 50f);
            ////dashjpg.SetAbsolutePosition(250f, 252);
            //if (countSigner == 0)
            //{
            //    dashjpg.SetAbsolutePosition(472, 558);
            //}
            //if (countSigner == 1)
            //{
            //    dashjpg.SetAbsolutePosition(472, 518);
            //}
            //else if (countSigner == 2)
            //{
            //    dashjpg.SetAbsolutePosition(472, 475);
            //}
            //else if (countSigner == 3)
            //{
            //    dashjpg.SetAbsolutePosition(472, 435);
            //}
            ////540
            //dashjpg.ScalePercent(35);
            ////dashjpg.IndentationLeft = 9f;
            ////dashjpg.SpacingAfter = 9f;
            //document.Add(dashjpg);
            #endregion
            document.Add(loHeader);
            document.Add(tblEntity);
            document.Add(tblSigner1);
            document.Add(tblIntenal);
            //document.Add(tblInternalUse);

            //  tblFooter.AddCell(loCell1);

            //FontSize = 8.5f;
            //ReportHeader = "\nGresham Private Equity Strategies, L.P. \nc/o Gresham Advisors, L.L.C.\n333 West Wacker Drive\nSuite 700 \nChicago, Illinois 60606 \n \nDear Sir or Madam:";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            //ReportHeader = "\nThe undersigned agrees to become a limited partner (a “Limited Partner”) of Gresham Private Equity Strategies, L.P. (the “Partnership”) and, in connection therewith, subscribes for and agrees to purchase an Interest in and to make a capital commitment (a “Commitment”) to GP 2018 Private Equity Strategies (the “Series”) in the amount of:  $ [insert isnull(confirmed amount, proposed amount)].";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            //ReportHeader = "\nThe undersigned acknowledges and agrees that: (i) the undersigned has carefully read and understands the Confidential Offering Memorandum for the Partnership dated August 2015 (the “Memorandum”), the Series Resolution creating the Series, any Series supplement and the Amended and Restated Limited Partnership Agreement of the Partnership (the “Partnership Agreement”) and agrees to each and every term therein; (ii) the representations, warranties, agreements, undertakings and acknowledgments made by the undersigned in the Commitment Agreement to the Partnership with respect to the 2011/2012 Series, 2013 Series, 2014 Series, 2015 Series, 2016 Series and/or 2017 Series and the previously completed Investor Profile and General Eligibility Form (“Investor Profile Form”) (including, without limitation, the undersigned’s purchaser suitability and benefit plan investor representations, anti-money laundering representations, indemnity and agreement to receive documents electronically) are true and correct in all material respects and are hereby confirmed for the benefit of the Series named above as of the date set forth below and may be used as a defense in any actions relating to the Partnership, the Series, any other series or the General Partner, and that it is only on the basis of such representations and warranties that the General Partner may be willing to accept the undersigned’s Commitment to the Series; (iii) the undersigned agrees to be bound to the terms and provisions of the Memorandum, the Series Resolution creating the Series, any Series supplement and the Partnership Agreement and that its signature below constitutes the execution and receipt of this Commitment Agreement and the execution and receipt of the Partnership Agreement; (iv) if the undersigned fails to make a required capital contribution, the Partnership, the Series and the General Partner will have all of their legal remedies as set forth in the Partnership Agreement; and (v) it shall do all acts and execute all additional documentation necessary for the purpose of making the Commitment as described herein. ";
            //lochunk = new Paragraph(ReportHeader, setFontsAllVerdana(FontSize, 0, 0));
            //lochunk.Alignment = Element.ALIGN_LEFT;// SetAlignment("center");
            //loCell.AddElement(lochunk);

            // document.Add(tblFooter);

            // document.Add(new Phrase("\n"));
            #endregion

            document.Close();
            return file;
        }
        catch (Exception ex)
        {

            Response.Write("ERROR :" + ex.Message.ToString());
            lblError.Visible = true;
            lblError.Text = "Error :" + ex.Message.ToString();
            file = "";
        }
        finally
        {
            document.Close();
        }
        return file;
    }
    public string GPRedemptionRequest(string RecommendationId, string[] SignerNames, string[] SignerTitle)
    {
        string FileName = string.Empty;
        string greshamquery = string.Empty;
        int totalCount = 0;
        try
        {
            string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 400;
            SqlDataAdapter dagersham = new SqlDataAdapter();
            DataSet ds_gresham = new DataSet();

            // greshamquery = "EXEC SP_S_SUBSCRIPTION_REDEMPTION_REQUEST @RecommendationId = 'DE4EF09B-8F02-E211-9A89-0019B9E7EE05'";
            greshamquery = "EXEC SP_S_SUBSCRIPTION_REDEMPTION_REQUEST  @RecommendationId = '" + RecommendationId + "'";

            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;
            if (totalCount > 0)
            {
                FileName = GPRedemeptionRequestForm(ds_gresham, SignerNames, SignerTitle);
            }
        }
        catch (Exception ex)
        {
            FileName = "";
            Response.Write("ERROR: " + ex.Message.ToString());
            lblError.Visible = true;
            lblError.Text = "ERROR: " + ex.Message.ToString();

        }
        return FileName;

    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        ClearControls();
        string ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.SubscriptionLetters);
        // string ReportOpFolder = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\";
        try
        {
            #region OLDCODE
            //foreach (GridViewRow row in GridView1.Rows)
            //{
            //    bool bProceed = false;
            //    bool bGPES = false;
            //    bool bGRAS = false;
            //    bool bSubscription = false;
            //    string sql = string.Empty;
            //    DataTable dtRecommendation = null;
            //    //string ssi_batchid = row.Cells[1].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            //    //string BatchStatusID = row.Cells[2].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");

            //    string RecommendationId = row.Cells[9].Text;
            //    string FinalFileName = row.Cells[10].Text;

            //    dtRecommendation = GetDataTable(RecommendationId, "GPES");
            //    {
            //        if (dtRecommendation.Rows.Count > 0)
            //        {
            //            bGPES = true;
            //        }
            //        else
            //        {
            //            dtRecommendation = GetDataTable(RecommendationId, "GRAS");
            //            if (dtRecommendation.Rows.Count > 0)
            //            {
            //                bGRAS = true;
            //            }
            //            else
            //            {
            //                dtRecommendation = GetDataTable(RecommendationId, "SUBSCRIPTION");
            //                if (dtRecommendation.Rows.Count > 0)
            //                {
            //                    bSubscription = true;
            //                }
            //            }
            //        }
            //    }

            //    if (dtRecommendation.Rows.Count > 0)
            //    {
            //        bProceed = true;
            //        string Amount = Convert.ToString(dtRecommendation.Rows[0]["Amount"]);
            //        string LegalEntity = Convert.ToString(dtRecommendation.Rows[0]["LegalEntity"]);
            //        string Signer1Name = Convert.ToString(dtRecommendation.Rows[0]["Signor1Name"]);
            //        string Signer1Title = Convert.ToString(dtRecommendation.Rows[0]["Signor1Title"]);
            //        string Signer2Name = Convert.ToString(dtRecommendation.Rows[0]["Signor2Name"]);
            //        string Signer2Title = Convert.ToString(dtRecommendation.Rows[0]["Signor2Title"]);
            //        string Signer3Name = Convert.ToString(dtRecommendation.Rows[0]["Signor3Name"]);
            //        string Signer3Title = Convert.ToString(dtRecommendation.Rows[0]["Signor3Title"]);

            //        string[] SignerNames = new string[3];
            //        string[] SignerTitle = new string[3];

            //        SignerNames[0] = Signer1Name;
            //        SignerNames[1] = Signer2Name;
            //        SignerNames[2] = Signer3Name;

            //        SignerTitle[0] = Signer1Title;
            //        SignerTitle[1] = Signer2Title;
            //        SignerTitle[2] = Signer3Title;

            //        //string[] SourceFiles = new string[2];
            //        if (bGPES)
            //        {
            //            //string PESCFile=  PrivateEquityStrategiesCommitment(Amount);
            //            //SourceFiles[0] = PESCFile;
            //        }
            //        else if (bGRAS)
            //        {
            //            //string RASCFile=  RealAssetsStrategiesCommitment(Amount);
            //            //SourceFiles[0] = RASCFile;
            //        }
            //        // string SignerFile = SignerPage(LegalEntity, SignerNames, SignerTitle);                         

            //        //SourceFiles[1] = SignerFile;

            //        //pdfMerge.MergeFiles(FolderPath+FinalFileName, SourceFiles);
            //        if (bGPES)
            //        {
            //            //File.Delete(PESCFile);
            //            //File.Delete(SignerFile);
            //        }
            //        else if (bGRAS)
            //        {
            //            //File.Delete(RASCFile);
            //            //File.Delete(SignerFile);
            //        }
            //        if (bSubscription)
            //        {
            //            string CloseDate = Convert.ToString(dtRecommendation.Rows[0]["Close Date"]);
            //            string FundName = Convert.ToString(dtRecommendation.Rows[0]["FundName"]);
            //            // string GPAdditionalSubscriptionFile = GPAdditionalSubscriptionForm(LegalEntity,CloseDate, FundName, Amount, SignerNames, SignerTitle);
            //        }
            //    }

            //}
            #endregion
            foreach (GridViewRow row in GridView1.Rows)
            {
                recommendationCount = GridView1.Rows.Count;
                bool Success = false;
                string RecommendationId = row.Cells[9].Text;

                string HouseholdId = row.Cells[1].Text;
                string LegalEntity = row.Cells[2].Text;
                LegalEntity = LegalEntity.Replace("&#39;", "'").ToString();
                string CloseDate = row.Cells[3].Text;
                string Fund = row.Cells[4].Text;
                string Amount = row.Cells[10].Text;
                string Year = row.Cells[11].Text.Replace("&nbsp;", "");
                string GRASSeries = row.Cells[12].Text.Replace("&nbsp;", "");
                string GPESSeries = row.Cells[13].Text.Replace("&nbsp;", "");
                Amount = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(Amount.Replace("(", "").Replace(")", "")));
                Amount = currencyFormat(Amount);
                // Amount = RoundToZeroDecimal(Amount);
                #region Commented - 12_20_2018 Logic Change
                ///*Account Signer*/
                //string Signer1Name = row.Cells[11].Text;
                //string Signer1Title = row.Cells[12].Text;
                //string Signer2Name = row.Cells[13].Text;
                //string Signer2Title = row.Cells[14].Text;
                //string Signer3Name = row.Cells[15].Text;
                //string Signer3Title = row.Cells[16].Text;

                ///*Legal Entity Signer*/
                //string LE_Signer1Name = row.Cells[17].Text;
                //string LE_Signer1Title = row.Cells[18].Text;
                //string LE_Signer2Name = row.Cells[19].Text;
                //string LE_Signer2Title = row.Cells[20].Text;
                //string LE_Signer3Name = row.Cells[21].Text;
                //string LE_Signer3Title = row.Cells[22].Text;

                //string ShowAccSignorFlg = row.Cells[23].Text;//show flg

                //string[] SignerNames = new string[3];
                //string[] SignerTitle = new string[3];

                //string[] LE_SignerNames = new string[3];
                //string[] LE_SignerTitle = new string[3];

                //SignerNames[0] = Signer1Name.Replace("&nbsp;", "");
                //SignerNames[1] = Signer2Name.Replace("&nbsp;", "");
                //SignerNames[2] = Signer3Name.Replace("&nbsp;", "");

                //LE_SignerNames[0] = LE_Signer1Name.Replace("&nbsp;", "");
                //LE_SignerNames[1] = LE_Signer2Name.Replace("&nbsp;", "");
                //LE_SignerNames[2] = LE_Signer3Name.Replace("&nbsp;", "");


                //SignerTitle[0] = Signer1Title.Replace("&nbsp;", "");
                //SignerTitle[1] = Signer2Title.Replace("&nbsp;", "");
                //SignerTitle[2] = Signer3Title.Replace("&nbsp;", "");


                //LE_SignerTitle[0] = LE_Signer1Title.Replace("&nbsp;", "");
                //LE_SignerTitle[1] = LE_Signer2Title.Replace("&nbsp;", "");
                //LE_SignerTitle[2] = LE_Signer3Title.Replace("&nbsp;", "");

                #endregion

                // DataTable dtReportType = GetDataTable(RecommendationId);
                //string ReportType = Convert.ToString(dtReportType.Rows[0]["Report"]);
                //string FinalFileName = Convert.ToString(dtReportType.Rows[0]["FileName"]);
                DataSet dtReportType = GetDataSet(RecommendationId);
                string ReportType = Convert.ToString(dtReportType.Tables[0].Rows[0]["Report"]);
                string FinalFileName = Convert.ToString(dtReportType.Tables[0].Rows[0]["FileName"]);

                string SignerFile_geps = string.Empty;
                string SignerFile_gras = string.Empty;
                //if (FinalFileName != "")
                //{
                //    FinalFileName = RemoveSpecialCharacters(FinalFileName);
                //}

                //Added 12_20_2018 Logic Change(Signor)
                string Signer1Name = Convert.ToString(dtReportType.Tables[1].Rows[0]["Signer1Name"]);
                string Signer1Title = Convert.ToString(dtReportType.Tables[1].Rows[0]["Signer1Title"]);
                string Signer2Name = Convert.ToString(dtReportType.Tables[1].Rows[0]["Signer2Name"]);
                string Signer2Title = Convert.ToString(dtReportType.Tables[1].Rows[0]["Signer2Title"]);
                string Signer3Name = Convert.ToString(dtReportType.Tables[1].Rows[0]["Signer3Name"]);
                string Signer3Title = Convert.ToString(dtReportType.Tables[1].Rows[0]["Signer3Title"]);

                string[] SignerNames = new string[3];
                string[] SignerTitle = new string[3];

                SignerNames[0] = Signer1Name.Replace("&nbsp;", "");
                SignerNames[1] = Signer2Name.Replace("&nbsp;", "");
                SignerNames[2] = Signer3Name.Replace("&nbsp;", "");
                SignerTitle[0] = Signer1Title.Replace("&nbsp;", "");
                SignerTitle[1] = Signer2Title.Replace("&nbsp;", "");
                SignerTitle[2] = Signer3Title.Replace("&nbsp;", "");


                if (ReportType.ToLower() == "gpes")
                {

                    string[] SourceFiles = new string[2];

                    string PESCFile = PrivateEquityStrategiesCommitment(Amount, LegalEntity, CloseDate,Year,GPESSeries);
                    if (PESCFile != "")
                    {
                        SourceFiles[0] = PESCFile;
                    }
                    #region Commented 12_20_2018 Logic Change
                    //if (ShowAccSignorFlg == "0")
                    //    SignerFile_geps = SignerPage(LegalEntity, LE_SignerNames, LE_SignerTitle, "gpes", CloseDate);
                    //else
                    //    SignerFile_geps = SignerPage(LegalEntity, SignerNames, SignerTitle, "gpes", CloseDate);
                    #endregion

                    //GPES Signer LEtter
                    SignerFile_geps = SignerPage(LegalEntity, SignerNames, SignerTitle, "gpes", CloseDate);

                    if (SignerFile_geps != "")
                    {
                        SourceFiles[1] = SignerFile_geps;
                    }

                    Success = MergeFiles(ReportOpFolder + FinalFileName, SourceFiles);

                    if (Success)
                        SuccessCount++;

                    if (PESCFile != "")
                    {
                        File.Delete(PESCFile);
                    }
                    if (SignerFile_geps != "")
                    {
                        File.Delete(SignerFile_geps);
                    }
                }
                else if (ReportType.ToLower() == "gras")
                {
                    string[] SourceFiles = new string[2];

                    string RASCFile = RealAssetsStrategiesCommitment(Amount, LegalEntity, CloseDate,Year,GRASSeries);
                    if (RASCFile != "")
                    {
                        SourceFiles[0] = RASCFile;
                    }
                    #region Commented 12_20_2018 Logic Change
                    //if (ShowAccSignorFlg == "0")
                    //    SignerFile_gras = SignerPage(LegalEntity, LE_SignerNames, LE_SignerTitle, "gras", CloseDate);
                    //else
                    //    SignerFile_gras = SignerPage(LegalEntity, SignerNames, SignerTitle, "gras", CloseDate);
                    #endregion

                    //GRA Signer LEtter
                    SignerFile_gras = SignerPage(LegalEntity, SignerNames, SignerTitle, "gras", CloseDate);

                    if (SignerFile_gras != "")
                    {
                        SourceFiles[1] = SignerFile_gras;
                    }

                    Success = MergeFiles(ReportOpFolder + FinalFileName, SourceFiles);
                    if (Success)
                        SuccessCount++;

                    if (RASCFile != "")
                    {
                        File.Delete(RASCFile);
                    }
                    if (SignerFile_gras != "")
                    {
                        File.Delete(SignerFile_gras);
                    }
                }
                else if (ReportType.ToLower() == "additionalcont")
                {
                    string FileName = string.Empty;
                    #region Commented 12_20_2018 Logic Change
                    //if (ShowAccSignorFlg == "0")
                    //    FileName = GPAdditionalSubscriptionForm(LegalEntity, CloseDate, Fund, Amount, LE_SignerNames, LE_SignerTitle);
                    //else
                    //    FileName = GPAdditionalSubscriptionForm(LegalEntity, CloseDate, Fund, Amount, SignerNames, SignerTitle);
                    #endregion
                    FileName = GPAdditionalSubscriptionForm(LegalEntity, CloseDate, Fund, Amount, SignerNames, SignerTitle);

                    if (FileName != "")
                    {
                        //string FinalFileName1= FinalFileName.Replace(".pdf", "_SUB.pdf");
                        try
                        {
                            File.Copy(FileName, ReportOpFolder + FinalFileName, true);
                            SuccessCount++;
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.ToString().Contains("because it is being used by another process"))
                            {
                                lblError.Visible = true;
                                lblError.Text = "File is open ,Kindly Close the file and try again.";
                                //Response.Write("ERROR :" + ex.Message.ToString());
                            }
                            else
                            {
                                lblError.Visible = true;
                                lblError.Text = "ERROR :" + ex.Message.ToString();
                                Response.Write("ERROR :" + ex.Message.ToString());
                            }
                        }
                        File.Delete(FileName);
                    }

                }
                else if (ReportType.ToLower() == "redemptionrequest")
                {
                    string FileNAme = GPRedemptionRequest(RecommendationId, SignerNames, SignerTitle);

                    if (FileNAme != "")
                    {
                        try
                        {
                            //string FinalFileName1 = FinalFileName.Replace(".pdf", "_REQ.pdf");
                            File.Copy(FileNAme, ReportOpFolder + FinalFileName, true);
                            SuccessCount++;
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.ToString().Contains("because it is being used by another process"))
                            {
                                lblError.Visible = true;
                                lblError.Text = "File is open ,Kindly Close the file and try again.";
                                //Response.Write("ERROR :" + ex.Message.ToString());
                            }
                            else
                            {
                                lblError.Visible = true;
                                lblError.Text = "ERROR :" + ex.Message.ToString();
                                Response.Write("ERROR :" + ex.Message.ToString());
                            }
                        }
                        File.Delete(FileNAme);
                    }
                }

            }

            // lblError.Text ="Letters Created Successfully";
        }
        catch (Exception ex)
        {
            if (ex.Message.ToString().Contains("because it is being used by another process"))
            {
                lblError.Visible = true;
                lblError.Text = "File is open ,Kindly Close the file and try again.";
                //Response.Write("ERROR :" + ex.Message.ToString());
            }
            else
            {
                lblError.Visible = true;
                lblError.Text = "ERROR :" + ex.Message.ToString();
                Response.Write("ERROR :" + ex.Message.ToString());
            }
        }
        finally
        {
            lblSuccess.Visible = true;
            lblSuccess.Text = SuccessCount + " Out of " + recommendationCount + " Created Successfully";
        }
    }
    public bool MergeFiles(string destinationFile, string[] sourceFiles)
    {
        bool Success = false;
        Document document = null;
        try
        {
            int f = 0;
            // we create a reader for a certain document
            PdfReader reader = new PdfReader(sourceFiles[f]);
            // we retrieve the total number of pages
            int n = reader.NumberOfPages;
            //Console.WriteLine("There are " + n + " pages in the original file.");
            // step 1: creation of a document-object
            document = new Document(reader.GetPageSizeWithRotation(1));
            //    document = new Document(reader.GetPageSizeWithRotation(1));

            // step 2: we create a writer that listens to the document
            //FileInfo file = new FileInfo();
            //file.Create();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(destinationFile, FileMode.Create));
            // step 3: we open the document
            document.Open();
            PdfContentByte cb = writer.DirectContent;
            PdfImportedPage page;
            int rotation;
            // step 4: we add content
            while (f < sourceFiles.Length)
            {
                int i = 0;
                while (i < n)
                {
                    i++;
                    document.SetPageSize(reader.GetPageSizeWithRotation(i));
                    document.NewPage();
                    page = writer.GetImportedPage(reader, i);
                    rotation = reader.GetPageRotation(i);
                    if (rotation == 90 || rotation == 270)
                    {
                        cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
                    }
                    else
                    {
                        cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                    }
                    //Console.WriteLine("Processed page " + i);
                }
                f++;
                if (f < sourceFiles.Length)
                {
                    if (sourceFiles[f] != null && Convert.ToString(sourceFiles[f]) != "")
                    {
                        reader = new PdfReader(sourceFiles[f]);
                        // we retrieve the total number of pages
                        n = reader.NumberOfPages;
                        //Console.WriteLine("There are " + n + " pages in the original file.");
                    }
                    else
                    {
                        //f++;
                        n = 0;
                    }
                }
            }
            // step 5: we close the document
            document.Close();
            Success = true;
        }
        catch (Exception ex)
        {
            Success = false;
            if (ex.Message.ToString().Contains("because it is being used by another process"))
            {
                lblError.Visible = true;
                lblError.Text = "File is open ,Kindly Close the file and try again.";
                //Response.Write("ERROR :" + ex.Message.ToString());
            }
            else
            {
                lblError.Visible = true;
                lblError.Text = "ERROR :" + ex.Message.ToString();
                Response.Write("ERROR :" + ex.Message.ToString());
            }
        }
        finally
        {
            document.Close();

        }
        return Success;
    }
    //private DataTable GetDataTable(String RecommendationId)
    private DataSet GetDataSet(String RecommendationId)
    {
        string greshamquery = string.Empty;
        int totalCount = 0;

        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {
            //  greshamquery = "EXEC SP_S_SUBSCRIPTION_GET_REPORT_TYPE  @RecommendationId = '9F6D56F4-28BB-E711-80C4-005056A04CD7'"; //--- GPES
            // greshamquery = "EXEC SP_S_SUBSCRIPTION_GET_REPORT_TYPE  @RecommendationId = 'CFAEC5CA-7FAF-E711-80C4-005056A04CD7'"; // ---GRAS 
            // greshamquery = "EXEC SP_S_SUBSCRIPTION_GET_REPORT_TYPE  @RecommendationId = '8EF3C151-8BFF-E711-80D7-005056A04CD7'"; // ---additional subscription
            // greshamquery = "EXEC SP_S_SUBSCRIPTION_GET_REPORT_TYPE  @RecommendationId = '0E689426-DAAD-E711-80C4-005056A04CD7'";

            greshamquery = "EXEC SP_S_SUBSCRIPTION_GET_REPORT_TYPE  @RecommendationId = '" + RecommendationId + "'";
            //  greshamquery = "EXEC SP_S_SUBSCRIPTION_GET_REPORT_TYPE  @RecommendationId = '8EF3C151-8BFF-E711-80D7-005056A04CD7'"; // ---additional subscription
            //if (LetterType == "GPES")
            //{
            //    greshamquery = "EXEC SP_S_SUBSCRIPTION_GPES_GRAS @GpesFlg = 1 , @RecommendationId = '9F6D56F4-28BB-E711-80C4-005056A04CD7'";
            //    // greshamquery = "EXEC SP_S_SUBSCRIPTION_GPES_GRAS @GpesFlg = 1 , @RecommendationId = '" + RecommendationId + "'";
            //}
            //else if (LetterType == "GRAS")
            //{
            //    greshamquery = "EXEC SP_S_SUBSCRIPTION_GPES_GRAS @GpesFlg = 0 , @RecommendationId = 'CFAEC5CA-7FAF-E711-80C4-005056A04CD7'";
            //    //greshamquery = "EXEC SP_S_SUBSCRIPTION_GPES_GRAS @GpesFlg = 0 , @RecommendationId = '" + RecommendationId + "'";
            //}
            //else if (LetterType == "SUBSCRIPTION")
            //{
            //    greshamquery = "EXEC SP_S_SUBSCRIPTION_ADDITIONAL_DOC @RecommendationId = '8EF3C151-8BFF-E711-80D7-005056A04CD7'";
            //    // greshamquery = "EXEC SP_S_SUBSCRIPTION_ADDITIONAL_DOC  @RecommendationId = '" + RecommendationId + "'";
            //}
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;
            return ds_gresham;
        }


        catch (Exception exc)
        {
            totalCount = 0;
            Response.Write("Stored PRocedure fails error desc:" + exc.Message);
            return null;
        }

        // return ds_gresham.Tables[0];
    }
    public PdfPTable addFooterAddress(string address, int PageNo, int LastPageNo)
    {
        PdfPTable fotable = new PdfPTable(2);

        fotable.HorizontalAlignment = Element.ALIGN_CENTER;
        fotable.TotalWidth = 100f;
        int[] headerwidths = { 80, 20 };
        fotable.SetWidths(headerwidths);

        PdfPCell loCell = new PdfPCell();
        Paragraph loparagrapgh = new Paragraph();

        int FontSize = 8;
        loCell = new PdfPCell();
        loparagrapgh = new Paragraph(address, setFontsAll(FontSize, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        loCell.HorizontalAlignment = Element.ALIGN_CENTER;
        loparagrapgh.SetAlignment("center");
        loCell.AddElement(loparagrapgh);
        loCell.BorderWidth = 0;
        loCell.Colspan = 2;
        fotable.AddCell(loCell);

        loCell = new PdfPCell();
        loparagrapgh = new Paragraph("", setFontsAll(FontSize, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        loCell.HorizontalAlignment = Element.ALIGN_CENTER;
        loparagrapgh.SetAlignment("right");
        loCell.AddElement(loparagrapgh);
        loCell.BorderWidth = 0;
        fotable.AddCell(loCell);

        loCell = new PdfPCell();
        Phrase phrase1 = new Phrase();
        Chunk chunk1 = new Chunk("Page ", setFontsAll(FontSize, 0, 1, new iTextSharp.text.Color(150, 150, 150)));
        Chunk chunk2 = new Chunk("" + PageNo, setFontsAll(FontSize, 1, 1, new iTextSharp.text.Color(150, 150, 150)));
        Chunk chunk3 = new Chunk(" of ", setFontsAll(FontSize, 0, 1, new iTextSharp.text.Color(150, 150, 150)));
        Chunk chunk4 = new Chunk("" + LastPageNo, setFontsAll(FontSize, 1, 1, new iTextSharp.text.Color(150, 150, 150)));
        //loChunk = new Paragraph("Page " + PageNo + " of " + LastPageNo, setFontsAll(8,1, 0, new iTextSharp.text.Color(150, 150, 150)));
        loCell.Colspan = 2;
        phrase1.Add(chunk1);
        phrase1.Add(chunk2);
        phrase1.Add(chunk3);
        phrase1.Add(chunk4);
        loparagrapgh = new Paragraph();
        loparagrapgh.Add(phrase1);

        loCell.HorizontalAlignment = Element.ALIGN_LEFT;
        loparagrapgh.SetAlignment("left");
        loCell.AddElement(loparagrapgh);
        loCell.BorderWidth = 0;
        fotable.AddCell(loCell);

        return fotable;
    }
    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic, iTextSharp.text.Color foColor)
    {

        string fontpath = HttpContext.Current.Server.MapPath(".") + "\\Verdana";
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\verdana.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\verdanab.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        }
        if (italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\verdanai.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\verdanaz.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        }
        return font;

    }
    public iTextSharp.text.Font setFontsAllVerdana(float size, int bold, int italic)
    {
        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = HttpContext.Current.Server.MapPath(".") + "\\Verdana";

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\verdana.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\verdanab.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\verdanai.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\verdanaz.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        }

        return font;
        #endregion
    }
    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic)
    {

        string fontpath = Server.MapPath(".") + "\\Verdana";
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\verdana.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\verdanab.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\verdanai.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\verdanaz.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        }
        return font;

    }
    public iTextSharp.text.Font setUnderline(float size, int underline)
    {
        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = Server.MapPath(".") + "\\Verdana";
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\verdana.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (underline == 1)
        {
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.UNDERLINE);
        }
        return font;
        #endregion
    }
    static public string RemoveSpecialCharacters(string str)
    {

        //System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex("[;\\/:*?\"<>|&']");
        System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex("[;\\/:*?\"<>|&',-]");
        string outputString = re.Replace(str, "");
        return outputString;
    }

    private string RoundToZeroDecimal(string lsFormatedString)
    {
        if (lsFormatedString != "")
        {
            lsFormatedString = String.Format("$ {0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));
            return lsFormatedString;
        }
        else
            return "";
    }

    public string currencyFormat(string Value)
    {
        string value = Value.Replace(",", "").Replace("$", "").Replace("%", "").Replace("(", "-").Replace(")", "");

        decimal ul = 0;
        if (value == "")
            ul = 0;//text.Text = "";
        else
            ul = Convert.ToDecimal(value);

        value = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C0}", ul);

        return value;
    }
}