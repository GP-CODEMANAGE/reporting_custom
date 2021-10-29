using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp.text.pdf.draw;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Globalization;
/// <summary>
/// Summary description for clsReportTemplate
/// </summary>
public class clsReportTemplate
{
    #region General Declaration

    public string STYLE_DEFAULT_TYPE = "style";
    public string DOCUMENT_HTML_START = "<html><head></head><body>";
    public string DOCUMENT_HTML_END = "</body></html>";
    public string REGEX_GROUP_SELECTOR = "selector";
    public string REGEX_GROUP_STYLE = "style";

    private StyleSheet _Styles;
    Boolean fbCheckExcel = false;
    public StreamWriter sw = null;
    public string strDescription = string.Empty;
    //bool bProceed = true;
    public int liPageSize = 10;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, Orientation, DynamicFlg, strDetails, Details;

    #endregion

    #region Properties
    private string _MailID = string.Empty;
    private string _LegalEntityID = string.Empty;
    private string _ContactNameID = string.Empty;
    private string _TemplateID = string.Empty;
    private string AsOf_Date = string.Empty;
    private string _AUMAsOfDate = string.Empty;
    private string _HHId = string.Empty;
    private string _BillingName = string.Empty;


    public string MailID
    {
        get
        {
            return _MailID;
        }
        set
        {
            _MailID = value;
        }
    }
    public string LegalEntityID
    {
        get
        {
            return _LegalEntityID;
        }
        set
        {
            _LegalEntityID = value;
        }
    }
    public string ContactNameID
    {
        get
        {
            return _ContactNameID;
        }
        set
        {
            _ContactNameID = value;
        }
    }
    public string TemplateID
    {
        get
        {
            return _TemplateID;
        }
        set
        {
            _TemplateID = value;
        }
    }
    public string AsOfDate
    {
        get
        {
            return AsOf_Date;
        }
        set
        {
            AsOf_Date = value;
        }
    }

    public string AUMAsOfDate
    {
        get
        {
            return _AUMAsOfDate;
        }
        set
        {
            _AUMAsOfDate = value;
        }
    }
    public string HHId
    {
        get
        {
            return _HHId;
        }
        set
        {
            _HHId = value;
        }
    }

    public string BillingName
    {
        get
        {
            return _BillingName;
        }
        set
        {
            _BillingName = value;
        }
    }


    #endregion


    #region General Methods

    private string GetDateSuperScript(string Num)
    {
        string retval = string.Empty;
        if (!string.IsNullOrEmpty(Num))
        {
            int no = Convert.ToInt32(Num);
            if (no == 1 || no == 21 || no == 31)
                retval = "st";
            else if (no == 2 || no == 22)
                retval = "nd";
            else if (no == 3 || no == 23)
                retval = "rd";
            else if (no > 3 && no < 21)
                retval = "th";
            else
                retval = "th";
        }
        return retval;
    }


    private string RoundUp(string lsFormatedString)
    {
        if (lsFormatedString != "")
        {
            lsFormatedString = String.Format("$ {0:#,###0.00;(#,###0.00)}", Convert.ToDecimal(lsFormatedString));
            return lsFormatedString;
        }
        else
            return "";
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


    private string GetLEName(string LegalEntity)
    {

        if (LegalEntity.Length > 150)
        {
            LegalEntity = LegalEntity.Substring(0, 150);
        }
        return LegalEntity;
    }

    private string Percentage(string lsFormatedString)
    {
        if (lsFormatedString != "")
        {
            string CurrentAllocation = String.Format("{0:#,###0.00;(#,###0.00)}%", Convert.ToDecimal(lsFormatedString));

            // string[] strFormat = lsFormatedString.Split('.');
            lsFormatedString = CurrentAllocation;// strFormat[0] + "%"; //String.Format("{0:#.0%;(#.0)%}", 		Convert.ToDecimal(lsFormatedString));
        }

        return lsFormatedString;
    }


    public enum ReportType
    {
        FidelityWire = 1,
        OtherWire = 2,
        NonStandard = 3,
        CapitalCallStatement = 4,
        DistributionStatement = 5,
        DistributionWireInstruction = 6,
        CapitalCallStatementCustom = 7,
        DistributionLetterCustom = 8,
        GreshamAdvisorsGLRLegalEntity = 9,
        GreshamAdvisorsGLRFund = 10,
        GreshamAdvisorsNonSpecific = 11,
        FundMemorandum = 12,
        UploadPdf = 13,
        SLOA = 14,
        Invoice = 15
    }

    private string getFinalSp(ReportType Type)
    {
        String lsSQL = String.Empty;
        object objHHId = HHId == "" ? "null" : "'" + HHId.ToString() + "'";

        if (Type == ReportType.FidelityWire)
        {
            lsSQL = "SP_S_FidelityWireInstructions " + MailID + ", '" + LegalEntityID + "', '" + ContactNameID + "'";
        }
        else if (Type == ReportType.OtherWire)
        {
            lsSQL = "SP_S_Other_Wire_Letter " + MailID + ", '" + LegalEntityID + "', '" + ContactNameID + "'";
        }
        else if (Type == ReportType.NonStandard)
        {
            lsSQL = "SP_S_NonStandardWireLetter @MailID = " + MailID + ", @legalentitynameid = '" + LegalEntityID + "', @ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.CapitalCallStatement)
        {
            lsSQL = "SP_S_CapitalCallStatements @MailID = " + MailID + ", @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.DistributionStatement)
        {
            lsSQL = "SP_S_DistributionStatements @MailID = " + MailID + ", @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.DistributionWireInstruction)
        {
            lsSQL = "SP_S_DistributionWireInstructions @MailID = " + MailID + ", @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.CapitalCallStatementCustom)
        {
            lsSQL = "SP_S_CapitalCallLetterCustom @MailID = " + MailID + ",@TemplateID='" + TemplateID + "', @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }

        else if (Type == ReportType.DistributionLetterCustom)
        {
            lsSQL = "SP_S_DistributionLetterCustom @MailID = " + MailID + ",@TemplateID='" + TemplateID + "', @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.GreshamAdvisorsGLRLegalEntity)
        {
            lsSQL = "SP_S_GreshamAdvisorsGLRLegalEntity @MailID = " + MailID + ",@TemplateID='" + TemplateID + "', @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.GreshamAdvisorsGLRFund)
        {
            lsSQL = "SP_S_GreshamAdvisorsGLRFund @MailID = " + MailID + ",@TemplateID='" + TemplateID + "', @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.GreshamAdvisorsNonSpecific)
        {
            lsSQL = "SP_S_GreshamAdvisorsGLNSRecipient @MailID = " + MailID + ",@TemplateID='" + TemplateID + "', @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.FundMemorandum)
        {
            lsSQL = "SP_S_MemorandumRegardingFund @MailID = " + MailID + ",@TemplateID='" + TemplateID + "', @legalentitynameid = '" + LegalEntityID + "',@ContactFullnameID = '" + ContactNameID + "'";
        }
        else if (Type == ReportType.UploadPdf)
        {
            lsSQL = "SP_S_Template_File_Name @MailID=" + MailID + ",@TemplateID='" + TemplateID + "',@LegalEntityNameID='" + LegalEntityID + "',@ContactFullnameID='" + ContactNameID + "'";
        }
        else if (Type == ReportType.SLOA)
        {
            lsSQL = "SP_S_FundMailingSignReqd  @MailID=" + MailID + ",@LegalEntityNameID='" + LegalEntityID + "',@ContactFullnameID='" + ContactNameID + "'";
        }
        else if (Type == ReportType.Invoice)
        {
            lsSQL = "SP_S_PDF_BILLING @MailID=" + MailID + ",@AsOfDate='" + AUMAsOfDate + "',@ContactID='" + ContactNameID + "',@HouseHoldID = " + objHHId + ",@BillingName='" + BillingName + "'";
        }

        return lsSQL;
    }
    #endregion

    #region Predefined Standard Report Templates

    #region Capital Call Wire Instructions

    public string CapitalCallWireInstruction()
    {
        Random rand = new Random();
        rand.Next();
        string FilePath = string.Empty;
        string DestinationFileName = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + DateTime.Now.ToString("ddMMMyyhhmmss") + System.Guid.NewGuid().ToString() + "CapitalCallWireInstruction.pdf";
        string[] SourceFileName = new string[4];

        SourceFileName[0] = GetFidelityWire();
        SourceFileName[1] = GetFidelityWireRetirement();
        SourceFileName[2] = GetOtherWire();
        SourceFileName[3] = GetNonStandardWire();

        for (int i = 0; i < SourceFileName.Length; i++)
        {
            if (FilePath != "")
            {
                if (SourceFileName[i] != "")
                {
                    FilePath = FilePath + "|" + SourceFileName[i];
                }
            }
            else
            {
                if (SourceFileName[i] != "")
                {
                    FilePath = "|" + SourceFileName[i];
                }
            }
        }


        if (FilePath != "")
        {
            FilePath = FilePath.Substring(1, FilePath.Length - 1);
            string[] strPath = FilePath.Split('|');

            PDFMerge PDF = new PDFMerge();
            PDF.MergeFiles(DestinationFileName, strPath);
        }
        else
        {
            DestinationFileName = "";
        }

        return DestinationFileName;
    }

    #region Fidelity Wire
    public string GetFidelityWire()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = "";//"See notes for this illustration located in the Appendix under Commitment Schedule for important information.";
        String lsSQL = getFinalSp(ReportType.FidelityWire);//Store Procedure call

        newdataset = clsDB.getDataSet(lsSQL);
        string str1 = string.Empty;

        var dv = newdataset.Tables[0].DefaultView;
        dv.RowFilter = "RetirementFlg = 0";
        var newDS = new DataSet();
        var newDT = dv.ToTable();
        newDS.Tables.Add(newDT);


        int DSCount = newDS.Tables[0].Rows.Count;

        string[] SourceFileName = new string[10];
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();
        rand.Next();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "Fidelity.xls";

        //String ls = HttpContext.Current.Server.MapPath("\\") + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        string Retirementflag = "0";
        for (int l = 0; l < DSCount; l++)
        {
            Retirementflag = Convert.ToString(newdataset.Tables[0].Rows[l]["RetirementFlg"]);
            if (DSCount > 0 && Retirementflag != "1")
            {
                DataTable table = newdataset.Tables[0].Copy();

                int liBlankCounter = 0;

                DataSet lodataset = new DataSet();
                lodataset.Tables.Add(table);
                DataSet loInsertdataset = lodataset.Copy();

                loInsertdataset.AcceptChanges();
                iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 31, 8);//10,10
                iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

                //AddHeader(document);//UnComment this to add header logo on each pdf sheet
                //AddFooter();//

                document.Open();

                lsTotalNumberofColumns = "";// loInsertdataset.Tables[0].Columns.Count + "";
                iTextSharp.text.Table loTable = new iTextSharp.text.Table(table.Columns.Count, table.Rows.Count);   // 2 rows, 2 columns           
                iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(1, 3);   // 2 rows, 2 columns           
                loTable.Width = 90;
                iTextSharp.text.Cell loCell = new Cell();
                Cell loCell1 = new Cell();
                Cell loCell2 = new Cell();
                Cell loCell21 = new Cell();
                Cell loCell22 = new Cell();
                Cell loCell3 = new Cell();
                Cell Cellsigner1 = new Cell();
                Cell Cellsigner2 = new Cell();
                Cell Cellsigner3 = new Cell();
                loTable.Cellpadding = 0f;
                loTable.Cellspacing = 0f;


                setTableProperty(loTable, ReportType.FidelityWire);
                String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

                int liTotalPage = 1;
                int liCurrentPage = 0;
                liPageSize = 38;



                #region Header

                Chunk lochunk = null;
                Chunk lochunk111 = null;
                Chunk lochunk222 = null;
                string Retirementflg = "0";
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    Retirementflg = Convert.ToString(table.Rows[i]["RetirementFlg"]);
                    if (Retirementflg != "1")
                    {
                        if (Convert.ToString(table.Rows[i]["ssi_signor1idname"]) != "")
                        {
                            lochunk = new Chunk("\n" + Convert.ToString(table.Rows[i]["ssi_signor1idname"]), setFontsAll(11, 0, 0));
                        }

                        if (Convert.ToString(table.Rows[i]["ssi_signor2idname"]) != "")
                        {
                            lochunk111 = new Chunk("\n" + Convert.ToString(table.Rows[i]["ssi_signor2idname"]), setFontsAll(11, 0, 0));
                        }

                        if (Convert.ToString(table.Rows[i]["ssi_signor3idname"]) != "")
                        {
                            lochunk222 = new Chunk("\n" + Convert.ToString(table.Rows[i]["ssi_signor3idname"]), setFontsAll(11, 0, 0));
                        }
                    }
                }

                //Chunk lochunk = new Chunk("\n" + SignerClientAccount, setFontsAll(11, 1, 0));
                if (lochunk != null)
                {
                    loCell.Add(lochunk);
                }
                if (lochunk111 != null)
                {
                    loCell.Add(lochunk111);
                }
                if (lochunk222 != null)
                {
                    loCell.Add(lochunk222);
                }
                loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                loCell.HorizontalAlignment = 1;
                loCell.Border = 0;

                if (table.Rows.Count > 0)
                {
                    if (Convert.ToString(table.Rows[0]["ssi_addressline1_mail"]) != "")
                    {
                        Chunk lochunk1 = new Chunk("\n" + table.Rows[0]["ssi_addressline1_mail"].ToString(), setFontsAll(11, 0, 0));
                        loCell.Add(lochunk1);
                        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                        loCell.HorizontalAlignment = 1;
                        loCell.Border = 0;
                    }
                    if (Convert.ToString(table.Rows[0]["ssi_addressline2_mail"]) != "")
                    {
                        Chunk lochunk2 = new Chunk("\n" + table.Rows[0]["ssi_addressline2_mail"].ToString(), setFontsAll(11, 0, 0));
                        loCell.Add(lochunk2);
                        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                        loCell.HorizontalAlignment = 1;
                        loCell.Border = 0;
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_addressline3_mail"]) != "")
                    {
                        Chunk lochunk3 = new Chunk("\n" + table.Rows[0]["ssi_addressline3_mail"].ToString(), setFontsAll(11, 0, 0));
                        loCell.Add(lochunk3);
                        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                        loCell.HorizontalAlignment = 1;
                        loCell.Border = 0;
                    }


                    string strAddress = "";
                    if (Convert.ToString(table.Rows[0]["ssi_city_mail"]) != "")
                    {
                        strAddress = Convert.ToString(table.Rows[0]["ssi_city_mail"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_stateprovince_mail"]) != "")
                    {
                        strAddress = strAddress + ", " + Convert.ToString(table.Rows[0]["ssi_stateprovince_mail"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_zipcode_mail"]) != "")
                    {
                        strAddress = strAddress + " " + Convert.ToString(table.Rows[0]["ssi_zipcode_mail"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_countryregion_mail"]) != "")
                    {
                        strAddress = strAddress + "\n" + Convert.ToString(table.Rows[0]["ssi_countryregion_mail"]);
                    }


                    Chunk lochunkAddress = new Chunk("\n" + strAddress, setFontsAll(11, 0, 0));
                    loCell.Add(lochunkAddress);
                    loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    loCell.HorizontalAlignment = 1;
                    loCell.EnableBorderSide(2);
                }




                #endregion

                #region As Of Date


                if (table.Rows.Count > 0)
                {
                    Chunk lochunkDate = null;
                    Chunk lochunkDate111 = null;

                    if (Convert.ToString(table.Rows[0]["Ssi_LetterDate"]) != "")
                    {
                        lochunkDate = new Chunk("\n" + Convert.ToString(table.Rows[0]["Ssi_LetterDate"]), setFontsAll(9, 0, 0));
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_wireinstructionletteraddress"]) != "")
                    {
                        lochunkDate111 = new Chunk("\n\n" + Convert.ToString(table.Rows[0]["ssi_wireinstructionletteraddress"]) + "\n\n", setFontsAll(9, 0, 0));
                    }

                    //Chunk lochunkDate = new Chunk("\n" + strAsOfDate, setFontsAll(11, 0, 0));
                    if (lochunkDate != null)
                    {
                        loCell1.Add(lochunkDate);
                    }
                    if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) != true)
                    {
                        if (lochunkDate111 != null)
                        {
                            //   loCell1.Add(lochunkDate111);
                        }
                    }
                    loCell1.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    loCell1.Leading = 13f;
                    loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    loCell1.Border = 0;
                }



                #endregion

                #region Legal Entity

                if (table.Rows.Count > 0)
                {
                    string strLegalEntity = "";
                    Chunk LegalEntity = null;
                    Chunk LegalEntity11 = null;
                    Chunk strLegalEntity12 = null;
                    if (Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"]) != "")
                    {
                        LegalEntity = new Chunk("\nRE: Legal Entity - " + Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"]), setFontsAll(9, 1, 0));
                        //strLegalEntity = "\n" + LegalEntity + Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_accountname1"]) != "")
                    {
                        LegalEntity11 = new Chunk("\n " + "      " + "Fidelity Account - " + Convert.ToString(table.Rows[0]["ssi_accountname1"]), setFontsAll(9, 1, 0));
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_accountnumber"]) != "")
                    {
                        strLegalEntity12 = new Chunk("\n " + "      " + Convert.ToString(table.Rows[0]["ssi_accountnumber"]), setFontsAll(9, 1, 0));
                    }


                    //if (Convert.ToString(table.Rows[0]["ssi_wireinstructionletterdear"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n Dear " + Convert.ToString(table.Rows[0]["ssi_wireinstructionletterdear"]) + ":";
                    //}


                    loCell2.Add(LegalEntity);
                    loCell2.Add(LegalEntity11);
                    loCell2.Add(strLegalEntity12);

                    loCell2.Border = 0;
                    loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    loCell2.Colspan = table.Columns.Count;


                }


                #endregion

                #region Fund

                if (table.Rows.Count > 0)
                {
                    int num = 1;

                    Chunk lochunkFundRemark = null;

                    // Commented by abhi 01/24/2018   Changes 
                    //if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) == false)
                    //{
                    //    lochunkFundRemark = new Chunk("\nDear " + Convert.ToString(table.Rows[0]["ssi_wireinstructionletterdear"]) + ":\n\n" + "Please process the following wire transfers on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));
                    //}
                    //else
                    //{
                    //    if (table.Rows.Count == 1)
                    //        lochunkFundRemark = new Chunk("\nThe following amount will be debited from your Fidelity account on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));
                    //    else
                    //        lochunkFundRemark = new Chunk("\nThe following amounts will be debited from your Fidelity account on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));
                    //}

                    if (table.Rows.Count == 1)
                        lochunkFundRemark = new Chunk("\nThe following amount will be debited from the above Fidelity account per the standing instructions on file on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));
                    else
                        lochunkFundRemark = new Chunk("\nThe following amounts will be debited from the above Fidelity account per the standing instructions on file on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));


                    Cell cellRemark = new Cell();
                    cellRemark.Add(lochunkFundRemark);
                    cellRemark.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    cellRemark.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    cellRemark.Border = 0;
                    string strRetirementFlg = "0";
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        strRetirementFlg = Convert.ToString(table.Rows[i]["RetirementFlg"]);
                        if (strRetirementFlg != "1")
                        {
                            if (Convert.ToString(table.Rows[i]["ssi_fundname"]) != "")
                            {
                                Chunk lochunkFundName1;
                                Chunk lochunkFundName2;
                                Chunk lochunkFundName3;

                                if (i == 0)
                                {
                                    lochunkFundName1 = new Chunk("    " + Convert.ToString(num++) + ".  " + "Wire ", setFontsAll(9, 0, 0));
                                    lochunkFundName2 = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])), setFontsAll(10, 1, 0));
                                    lochunkFundName3 = new Chunk(" to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                                    //lochunkFundName = new Chunk("    " + Convert.ToString(num++) + ".  " + "Wire " + RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])) + " to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                                }
                                else
                                {
                                    lochunkFundName1 = new Chunk("\n    " + Convert.ToString(num++) + ".  " + "Wire ", setFontsAll(9, 0, 0));
                                    lochunkFundName2 = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])), setFontsAll(10, 1, 0));
                                    lochunkFundName3 = new Chunk(" to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                                    //lochunkFundName = new Chunk("\n    " + Convert.ToString(num++) + ".  " + "Wire " + RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])) + " to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                                }
                                loCell3.Add(lochunkFundName1);
                                loCell3.Add(lochunkFundName2);
                                loCell3.Add(lochunkFundName3);
                                loCell3.Colspan = loInsertdataset.Tables[0].Columns.Count;
                                loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell3.Border = 0;
                            }
                        }
                    }




                #endregion

                    #region Instructions

                    Chunk lochunkinstr = new Chunk("Standing instructions are on file for each wire above. ", setFontsAll(9, 1, 0));
                    Cell cellinstr = new Cell();
                    cellinstr.Add(lochunkinstr);
                    cellinstr.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    cellinstr.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    cellinstr.Border = 0;

                    Chunk lochunkinstr1 = null;
                    if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) == false)
                    {
                        // Commented by abhi 01/24/2018
                        //  lochunkinstr1 = new Chunk("Please contact " + Convert.ToString(table.Rows[0]["ssi_secownerfname_mail"]) + " " + Convert.ToString(table.Rows[0]["ssi_secownerlname_mail"]) + " of Gresham Partners, LLC at (312)960-0200 if you have any questions.\n\nSincerely,\n", setFontsAll(9, 0, 0));
                        lochunkinstr1 = new Chunk("Please contact " + Convert.ToString(table.Rows[0]["ssi_secownerfname_mail"]) + " " + Convert.ToString(table.Rows[0]["ssi_secownerlname_mail"]) + " of Gresham Partners, LLC at (312)960-0200 if you have any questions.\n", setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        lochunkinstr1 = new Chunk("Please contact " + Convert.ToString(table.Rows[0]["ssi_secownerfname_mail"]) + " " + Convert.ToString(table.Rows[0]["ssi_secownerlname_mail"]) + " of Gresham Partners, LLC at (312)960-0200 if you have any questions.\n", setFontsAll(9, 0, 0));
                    }
                    //string SignatureLine = "\n_______________________________________\n";

                    Chunk signer = null;
                    Chunk signer1 = null;
                    Chunk signer2 = null;
                    if (Convert.ToString(table.Rows[l]["ssi_signor1idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                    {
                        if (Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                        {
                            signer = new Chunk("\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor1idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer1title"]), setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            signer = new Chunk("\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor1idname"]), setFontsAll(9, 0, 0));
                        }
                    }

                    if (Convert.ToString(table.Rows[l]["ssi_signor2idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                    {
                        if (Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                        {
                            signer1 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor2idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer2title"]), setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            signer1 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor2idname"]), setFontsAll(9, 0, 0));
                        }
                    }

                    if (Convert.ToString(table.Rows[l]["ssi_signor3idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                    {
                        if (Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                        {
                            signer2 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor3idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer3title"]), setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            signer2 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor3idname"]), setFontsAll(9, 0, 0));
                        }
                    }

                    //Chunk lochunkinstr1 = new Chunk("\n" + Instr + signer, setFontsAll(11, 0, 0));
                    Cell cellinstr1 = new Cell();
                    cellinstr1.Add(lochunkinstr1);

                    if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) == false)
                    {
                        //// Commented by abhi 01/24/2018
                        //if (signer != null)
                        //{
                        //    cellinstr1.Add(signer);
                        //}
                        //if (signer1 != null)
                        //{
                        //    cellinstr1.Add(signer1);
                        //}
                        //if (signer2 != null)
                        //{
                        //    cellinstr1.Add(signer2);
                        //}
                    }
                    cellinstr1.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    cellinstr1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    cellinstr1.Border = 0;

                    #endregion

                    #region Signer Client Account

                    //string SignerClient = string.Empty;

                    //for (int i = 0; i < table.Rows.Count; i++)
                    //{
                    //    if (Convert.ToString(table.Rows[i]["ssi_signor1idname"]) != "" && Convert.ToString(table.Rows[i]["ssi_signer1title"]) != "")
                    //    {
                    //        SignerClient = "\n\n" + Convert.ToString(table.Rows[i]["ssi_signor1idname"]) + " " + Convert.ToString(table.Rows[i]["ssi_signer1title"]);
                    //    }


                    //    if (Convert.ToString(table.Rows[i]["ssi_signor2idname"]) != "" && Convert.ToString(table.Rows[i]["ssi_signer2title"]) != "")
                    //    {
                    //        SignerClient = SignerClient + "\n" + Convert.ToString(table.Rows[i]["ssi_signor2idname"]) + " " + Convert.ToString(table.Rows[i]["ssi_signer2title"]);
                    //    }


                    //    if (Convert.ToString(table.Rows[i]["ssi_signor3idname"]) != "" && Convert.ToString(table.Rows[i]["ssi_signer3title"]) != "")
                    //    {
                    //        SignerClient = SignerClient + "\n" + Convert.ToString(table.Rows[i]["ssi_signor3idname"]) + " " + Convert.ToString(table.Rows[i]["ssi_signer3title"]);
                    //    }
                    //}

                    //if (SignerClient != "")
                    //{
                    //    Chunk Signer1 = new Chunk("\n\n" + SignerClient, setFontsAll(11, 0, 0));

                    //    Cellsigner1.Add(Signer1);
                    //    Cellsigner1.Colspan = loInsertdataset.Tables[0].Columns.Count - 10;
                    //    Cellsigner1.HorizontalAlignment = 1;
                    //    Cellsigner1.EnableBorderSide(3);
                    //}


                    #endregion


                    loTable.AddCell(loCell); //header
                    loTable.AddCell(loCell1); // as of date
                    loTable.AddCell(loCell2); // Lagal Entity
                    //loTable.AddCell(loCell21); // Lagal Entity
                    //loTable.AddCell(loCell22); // Lagal Entity
                    loTable.AddCell(cellRemark); // Fund Remark
                    loTable.AddCell(loCell3); // Fund Details
                    //   loTable.AddCell(cellinstr); // Instructions  Commented by abhi 01/24/2018
                    loTable.AddCell(cellinstr1); // Instructions 1
                    //if (SignerClient != "")
                    //{
                    //    loTable.AddCell(Cellsigner1); // Signer1
                    //}

                }
                document.Add(loTable);

                document.Close();
            }

        }

        if (DSCount > 0)
        {
            try
            {
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }

        return fsFinalLocation.Replace(".xls", ".pdf");
    }
    #endregion

    #region Fidelity Wire (Retirement)
    public string GetFidelityWireRetirement()
    {
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = "";//"See notes for this illustration located in the Appendix under Commitment Schedule for important information.";
        String lsSQL = getFinalSp(ReportType.FidelityWire);//Store Procedure call

        newdataset = clsDB.getDataSet(lsSQL);
        string str1 = string.Empty;

        var dv = newdataset.Tables[0].DefaultView;
        dv.RowFilter = "RetirementFlg = 1";
        var newDS = new DataSet();
        var newDT = dv.ToTable();
        newDS.Tables.Add(newDT);

        int DSCount = newDS.Tables[0].Rows.Count;

        string[] SourceFileName = new string[10];
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();
        rand.Next();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "FidelityRetirement.xls";

        //String ls = HttpContext.Current.Server.MapPath("\\") + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        string Retirementflag = "0";
        for (int l = 0; l < DSCount; l++)
        {
            Retirementflag = Convert.ToString(newdataset.Tables[0].Rows[l]["RetirementFlg"]);
            if (DSCount > 0 && Retirementflag == "1")
            {
                DataTable table = newdataset.Tables[0].Copy();

                int liBlankCounter = 0;

                DataSet lodataset = new DataSet();
                lodataset.Tables.Add(table);
                DataSet loInsertdataset = lodataset.Copy();

                loInsertdataset.AcceptChanges();
                iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 31, 8);//10,10
                iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

                //AddHeader(document);//UnComment this to add header logo on each pdf sheet
                //AddFooter();//

                document.Open();

                lsTotalNumberofColumns = "";// loInsertdataset.Tables[0].Columns.Count + "";
                iTextSharp.text.Table loTable = new iTextSharp.text.Table(table.Columns.Count, table.Rows.Count);   // 2 rows, 2 columns           
                iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(1, 3);   // 2 rows, 2 columns           
                loTable.Width = 90;
                iTextSharp.text.Cell loCell = new Cell();
                Cell loCell1 = new Cell();
                Cell loCell2 = new Cell();
                Cell loCell21 = new Cell();
                Cell loCell22 = new Cell();
                Cell loCell3 = new Cell();
                Cell Cellsigner1 = new Cell();
                Cell Cellsigner2 = new Cell();
                Cell Cellsigner3 = new Cell();
                loTable.Cellpadding = 0f;
                loTable.Cellspacing = 0f;


                setTableProperty(loTable, ReportType.FidelityWire);
                String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();

                int liTotalPage = 1;
                int liCurrentPage = 0;
                liPageSize = 38;



                #region Header

                Chunk lochunk = null;
                Chunk lochunk111 = null;
                Chunk lochunk222 = null;

                for (int i = 0; i < table.Rows.Count; i++)
                {
                    string Retirmentflg = Convert.ToString(table.Rows[i]["RetirementFlg"]);
                    if (Retirmentflg == "1")
                    {
                        if (Convert.ToString(table.Rows[i]["ssi_signor1idname"]) != "")
                        {
                            lochunk = new Chunk("\n" + Convert.ToString(table.Rows[i]["ssi_signor1idname"]), setFontsAll(11, 0, 0));
                        }

                        if (Convert.ToString(table.Rows[i]["ssi_signor2idname"]) != "")
                        {
                            lochunk111 = new Chunk("\n" + Convert.ToString(table.Rows[i]["ssi_signor2idname"]), setFontsAll(11, 0, 0));
                        }

                        if (Convert.ToString(table.Rows[i]["ssi_signor3idname"]) != "")
                        {
                            lochunk222 = new Chunk("\n" + Convert.ToString(table.Rows[i]["ssi_signor3idname"]), setFontsAll(11, 0, 0));
                        }
                    }
                }

                //Chunk lochunk = new Chunk("\n" + SignerClientAccount, setFontsAll(11, 1, 0));
                if (lochunk != null)
                {
                    loCell.Add(lochunk);
                }
                if (lochunk111 != null)
                {
                    loCell.Add(lochunk111);
                }
                if (lochunk222 != null)
                {
                    loCell.Add(lochunk222);
                }
                loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                loCell.HorizontalAlignment = 1;
                loCell.Border = 0;

                if (table.Rows.Count > 0)
                {
                    if (Convert.ToString(table.Rows[0]["ssi_addressline1_mail"]) != "")
                    {
                        Chunk lochunk1 = new Chunk("\n" + table.Rows[0]["ssi_addressline1_mail"].ToString(), setFontsAll(11, 0, 0));
                        loCell.Add(lochunk1);
                        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                        loCell.HorizontalAlignment = 1;
                        loCell.Border = 0;
                    }
                    if (Convert.ToString(table.Rows[0]["ssi_addressline2_mail"]) != "")
                    {
                        Chunk lochunk2 = new Chunk("\n" + table.Rows[0]["ssi_addressline2_mail"].ToString(), setFontsAll(11, 0, 0));
                        loCell.Add(lochunk2);
                        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                        loCell.HorizontalAlignment = 1;
                        loCell.Border = 0;
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_addressline3_mail"]) != "")
                    {
                        Chunk lochunk3 = new Chunk("\n" + table.Rows[0]["ssi_addressline3_mail"].ToString(), setFontsAll(11, 0, 0));
                        loCell.Add(lochunk3);
                        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                        loCell.HorizontalAlignment = 1;
                        loCell.Border = 0;
                    }


                    string strAddress = "";
                    if (Convert.ToString(table.Rows[0]["ssi_city_mail"]) != "")
                    {
                        strAddress = Convert.ToString(table.Rows[0]["ssi_city_mail"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_stateprovince_mail"]) != "")
                    {
                        strAddress = strAddress + ", " + Convert.ToString(table.Rows[0]["ssi_stateprovince_mail"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_zipcode_mail"]) != "")
                    {
                        strAddress = strAddress + " " + Convert.ToString(table.Rows[0]["ssi_zipcode_mail"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_countryregion_mail"]) != "")
                    {
                        strAddress = strAddress + "\n" + Convert.ToString(table.Rows[0]["ssi_countryregion_mail"]);
                    }


                    Chunk lochunkAddress = new Chunk("\n" + strAddress, setFontsAll(11, 0, 0));
                    loCell.Add(lochunkAddress);
                    loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    loCell.HorizontalAlignment = 1;
                    loCell.EnableBorderSide(2);
                }




                #endregion

                #region As Of Date


                if (table.Rows.Count > 0)
                {
                    Chunk lochunkDate = null;
                    Chunk lochunkDate111 = null;

                    if (Convert.ToString(table.Rows[0]["Ssi_LetterDate"]) != "")
                    {
                        lochunkDate = new Chunk("\n" + Convert.ToString(table.Rows[0]["Ssi_LetterDate"]), setFontsAll(9, 0, 0));
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_wireinstructionletteraddress"]) != "")
                    {
                        lochunkDate111 = new Chunk("\n\n" + Convert.ToString(table.Rows[0]["ssi_wireinstructionletteraddress"]) + "\n\n", setFontsAll(9, 0, 0));
                    }

                    //Chunk lochunkDate = new Chunk("\n" + strAsOfDate, setFontsAll(11, 0, 0));
                    if (lochunkDate != null)
                    {
                        loCell1.Add(lochunkDate);
                    }
                    if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) != true)
                    {
                        if (lochunkDate111 != null)
                        {
                            loCell1.Add(lochunkDate111);
                        }
                    }
                    loCell1.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    loCell1.Leading = 13f;
                    loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    loCell1.Border = 0;
                }



                #endregion

                #region Legal Entity

                if (table.Rows.Count > 0)
                {
                    string strLegalEntity = "";
                    Chunk LegalEntity = null;
                    Chunk LegalEntity11 = null;
                    Chunk strLegalEntity12 = null;
                    if (Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"]) != "")
                    {
                        LegalEntity = new Chunk("\nRE: Legal Entity - " + Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"]), setFontsAll(9, 1, 0));
                        //strLegalEntity = "\n" + LegalEntity + Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"]);
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_accountname1"]) != "")
                    {
                        LegalEntity11 = new Chunk("\n " + "      " + "Fidelity Account - " + Convert.ToString(table.Rows[0]["ssi_accountname1"]), setFontsAll(9, 1, 0));
                    }

                    if (Convert.ToString(table.Rows[0]["ssi_accountnumber"]) != "")
                    {
                        strLegalEntity12 = new Chunk("\n " + "      " + Convert.ToString(table.Rows[0]["ssi_accountnumber"]), setFontsAll(9, 1, 0));
                    }

                    //if (Convert.ToString(table.Rows[0]["ssi_wireinstructionletterdear"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n Dear " + Convert.ToString(table.Rows[0]["ssi_wireinstructionletterdear"]) + ":";
                    //}


                    loCell2.Add(LegalEntity);
                    loCell2.Add(LegalEntity11);
                    loCell2.Add(strLegalEntity12);

                    loCell2.Border = 0;
                    loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    loCell2.Colspan = table.Columns.Count;


                }


                #endregion

                #region Fund

                if (table.Rows.Count > 0)
                {
                    int num = 1;

                    Chunk lochunkFundRemark = null;
                    if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) == false)
                    {
                        lochunkFundRemark = new Chunk("\nDear " + Convert.ToString(table.Rows[0]["ssi_wireinstructionletterdear"]) + ":\n\n" + "Please process the following wire transfers on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        if (table.Rows.Count == 1)
                            lochunkFundRemark = new Chunk("\nThe following amount will be debited from your Fidelity account on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));
                        else
                            lochunkFundRemark = new Chunk("\nThe following amounts will be debited from your Fidelity account on " + Convert.ToString(table.Rows[0]["As Of Date"]) + ".", setFontsAll(9, 0, 0));
                    }
                    Cell cellRemark = new Cell();
                    cellRemark.Add(lochunkFundRemark);
                    cellRemark.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    cellRemark.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    cellRemark.Border = 0;
                    string strRetirementFlg = "0";
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        strRetirementFlg = Convert.ToString(table.Rows[i]["RetirementFlg"]);

                        if (Convert.ToString(table.Rows[i]["ssi_fundname"]) != "" && strRetirementFlg == "1")
                        {
                            Chunk lochunkFundName1;
                            Chunk lochunkFundName2;
                            Chunk lochunkFundName3;

                            if (i == 0)
                            {
                                lochunkFundName1 = new Chunk("    " + Convert.ToString(num++) + ".  " + "Wire ", setFontsAll(9, 0, 0));
                                lochunkFundName2 = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])), setFontsAll(10, 1, 0));
                                lochunkFundName3 = new Chunk(" to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                                //lochunkFundName = new Chunk("    " + Convert.ToString(num++) + ".  " + "Wire " + RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])) + " to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                            }
                            else
                            {
                                lochunkFundName1 = new Chunk("\n    " + Convert.ToString(num++) + ".  " + "Wire ", setFontsAll(9, 0, 0));
                                lochunkFundName2 = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])), setFontsAll(10, 1, 0));
                                lochunkFundName3 = new Chunk(" to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                                //lochunkFundName = new Chunk("\n    " + Convert.ToString(num++) + ".  " + "Wire " + RoundUp(Convert.ToString(table.Rows[i]["ssi_currentcall_ccsf"])) + " to " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                            }
                            loCell3.Add(lochunkFundName1);
                            loCell3.Add(lochunkFundName2);
                            loCell3.Add(lochunkFundName3);
                            loCell3.Colspan = loInsertdataset.Tables[0].Columns.Count;
                            loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            loCell3.Border = 0;
                        }
                    }




                #endregion

                    #region Instructions

                    // Chunk lochunkinstr = new Chunk("Standing instructions are on file for each wire above. ", setFontsAll(9, 1, 0));
                    //  Cell cellinstr = new Cell();
                    // cellinstr.Add(lochunkinstr);
                    // cellinstr.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    // cellinstr.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //  cellinstr.Border = 0;

                    Chunk lochunkinstr1 = null;
                    if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) == false)
                    {
                        lochunkinstr1 = new Chunk("Please contact " + Convert.ToString(table.Rows[0]["ssi_secownerfname_mail"]) + " " + Convert.ToString(table.Rows[0]["ssi_secownerlname_mail"]) + " of Gresham Partners, LLC at (312)960-0200 if you have any questions.\n\nSincerely,\n", setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        lochunkinstr1 = new Chunk("Please contact " + Convert.ToString(table.Rows[0]["ssi_secownerfname_mail"]) + " " + Convert.ToString(table.Rows[0]["ssi_secownerlname_mail"]) + " of Gresham Partners, LLC at (312)960-0200 if you have any questions.\n", setFontsAll(9, 0, 0));
                    }
                    //string SignatureLine = "\n_______________________________________\n";

                    Chunk signer = null;
                    Chunk signer1 = null;
                    Chunk signer2 = null;
                    if (Convert.ToString(table.Rows[l]["ssi_signor1idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                    {
                        if (Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                        {
                            signer = new Chunk("\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor1idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer1title"]), setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            signer = new Chunk("\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor1idname"]), setFontsAll(9, 0, 0));
                        }
                    }

                    if (Convert.ToString(table.Rows[l]["ssi_signor2idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                    {
                        if (Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                        {
                            signer1 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor2idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer2title"]), setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            signer1 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor2idname"]), setFontsAll(9, 0, 0));
                        }
                    }

                    if (Convert.ToString(table.Rows[l]["ssi_signor3idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                    {
                        if (Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                        {
                            signer2 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor3idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer3title"]), setFontsAll(9, 0, 0));
                        }
                        else
                        {
                            signer2 = new Chunk("\n\n" + "\n_______________________________________\n" + Convert.ToString(table.Rows[l]["ssi_signor3idname"]), setFontsAll(9, 0, 0));
                        }
                    }

                    //Chunk lochunkinstr1 = new Chunk("\n" + Instr + signer, setFontsAll(11, 0, 0));
                    Cell cellinstr1 = new Cell();
                    cellinstr1.Add(lochunkinstr1);

                    if (Convert.ToBoolean(table.Rows[0]["ssi_discretionaryflg"]) == false)
                    {
                        if (signer != null)
                        {
                            cellinstr1.Add(signer);
                        }
                        if (signer1 != null)
                        {
                            cellinstr1.Add(signer1);
                        }
                        if (signer2 != null)
                        {
                            cellinstr1.Add(signer2);
                        }
                    }
                    cellinstr1.Colspan = loInsertdataset.Tables[0].Columns.Count;
                    cellinstr1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    cellinstr1.Border = 0;

                    #endregion

                    #region Signer Client Account

                    //string SignerClient = string.Empty;

                    //for (int i = 0; i < table.Rows.Count; i++)
                    //{
                    //    if (Convert.ToString(table.Rows[i]["ssi_signor1idname"]) != "" && Convert.ToString(table.Rows[i]["ssi_signer1title"]) != "")
                    //    {
                    //        SignerClient = "\n\n" + Convert.ToString(table.Rows[i]["ssi_signor1idname"]) + " " + Convert.ToString(table.Rows[i]["ssi_signer1title"]);
                    //    }


                    //    if (Convert.ToString(table.Rows[i]["ssi_signor2idname"]) != "" && Convert.ToString(table.Rows[i]["ssi_signer2title"]) != "")
                    //    {
                    //        SignerClient = SignerClient + "\n" + Convert.ToString(table.Rows[i]["ssi_signor2idname"]) + " " + Convert.ToString(table.Rows[i]["ssi_signer2title"]);
                    //    }


                    //    if (Convert.ToString(table.Rows[i]["ssi_signor3idname"]) != "" && Convert.ToString(table.Rows[i]["ssi_signer3title"]) != "")
                    //    {
                    //        SignerClient = SignerClient + "\n" + Convert.ToString(table.Rows[i]["ssi_signor3idname"]) + " " + Convert.ToString(table.Rows[i]["ssi_signer3title"]);
                    //    }
                    //}

                    //if (SignerClient != "")
                    //{
                    //    Chunk Signer1 = new Chunk("\n\n" + SignerClient, setFontsAll(11, 0, 0));

                    //    Cellsigner1.Add(Signer1);
                    //    Cellsigner1.Colspan = loInsertdataset.Tables[0].Columns.Count - 10;
                    //    Cellsigner1.HorizontalAlignment = 1;
                    //    Cellsigner1.EnableBorderSide(3);
                    //}


                    #endregion

                    loTable.AddCell(loCell); //header
                    loTable.AddCell(loCell1); // as of date
                    loTable.AddCell(loCell2); // Lagal Entity
                    //loTable.AddCell(loCell21); // Lagal Entity
                    //loTable.AddCell(loCell22); // Lagal Entity
                    loTable.AddCell(cellRemark); // Fund Remark
                    loTable.AddCell(loCell3); // Fund Details
                    //  loTable.AddCell(cellinstr); // Instructions
                    loTable.AddCell(cellinstr1); // Instructions 1
                    //if (SignerClient != "")
                    //{
                    //    loTable.AddCell(Cellsigner1); // Signer1
                    //}

                }
                document.Add(loTable);

                document.Close();
            }

        }

        if (DSCount > 0)
        {
            try
            {
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }

        return fsFinalLocation.Replace(".xls", ".pdf");
    }
    #endregion


    #region Other Wire
    public string GetOtherWire()
    {
        int i = 1;
        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = "";// "See notes for this illustration located in the Appendix under Commitment Schedule for important information.";
        String lsSQL = getFinalSp(ReportType.OtherWire);
        //String lsSQL = "SP_S_Other_Wire_Letter 4, 'D5DC3BE5-6D15-DE11-8391-001D09665E8F', 'E9063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        string str1 = string.Empty;
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();
        DataSet lodataset = new DataSet();
        lodataset.Tables.Add(table);
        string str = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + DateTime.Now.ToString("ddMMMyyhhmmssss") + "OtherWire" + System.Guid.NewGuid().ToString() + ".pdf"; //FileName

        if (File.Exists(str))
        {
            File.Delete(str);
        }

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();
        rand.Next();

        //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 10);
        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));



        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "OtherWire.xls";

        string[] SourceFileName = new string[DSCount];


        if (DSCount > 0)
        {
            for (int l = 0; l < DSCount; l++)
            {
                if (l != 0)
                {
                    document.NewPage();
                }

                int liBlankCounter = 0;

                DataSet loInsertdataset = lodataset.Copy();

                loInsertdataset.AcceptChanges();

                //AddHeader(document);
                //AddFooter();

                document.Open();

                //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                //png.SetAbsolutePosition(45, 800);//540
                ////png.ScaleToFit(288f, 42f);
                //png.ScalePercent(8);
                //document.Add(png);


                lsTotalNumberofColumns = loInsertdataset.Tables[0].Columns.Count + "";
                iTextSharp.text.Table loTable = new iTextSharp.text.Table(table.Columns.Count, table.Rows.Count);   // 2 rows, 2 columns           
                iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(28, 3);   // 2 rows, 2 columns           
                loTable.Width = 90;
                loTable1.Width = 90;
                iTextSharp.text.Cell loCell = new Cell();
                Cell loCell1 = new Cell();
                Cell loCell2 = new Cell();
                Cell loCell3 = new Cell();
                Cell loCell4 = new Cell();
                Cell loCell44 = new Cell();
                Cell loCell5 = new Cell();
                Cell Cellsigner1 = new Cell();
                Cell Cellsigner2 = new Cell();
                Cell Cellsigner3 = new Cell();
                loTable.Cellpadding = 0f;
                loTable.Cellspacing = 0f;


                setTableProperty(loTable, ReportType.OtherWire);
                String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
                int liTotalPage = (loInsertdataset.Tables[0].Rows.Count / liPageSize);
                int liCurrentPage = 0;
                if (loInsertdataset.Tables[0].Rows.Count % liPageSize != 0)
                {
                    liTotalPage = liTotalPage + 1;
                }
                else
                {
                    liPageSize = 28;
                    liTotalPage = liTotalPage + 1;
                }

                //check the length of the column name to set the pagesize.
                for (int j = 0; j < loInsertdataset.Tables[0].Columns.Count; j++)
                {
                    if (loInsertdataset.Tables[0].Columns[j].ColumnName.Length > 30)
                    {
                        liPageSize = 28;
                    }
                }


                for (int liRowCount = 0; liRowCount < loInsertdataset.Tables[0].Rows.Count; liRowCount++)
                {
                    if (liRowCount % liPageSize == 0)
                    {
                        document.Add(loTable);

                        if (liRowCount != 0)
                        {
                            liCurrentPage = liCurrentPage + 1;
                            //document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                            document.NewPage();
                        }
                        loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, loInsertdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                        setTableProperty(loTable, ReportType.OtherWire);
                    }
                }




                #region Header

                Chunk SignerClientAccount = null;
                Chunk SignerClientAccount1 = null;
                Chunk SignerClientAccount2 = null;

                if (Convert.ToString(table.Rows[l]["ssi_signor1idname"]) != "")
                {
                    SignerClientAccount = new Chunk(Convert.ToString(table.Rows[l]["ssi_signor1idname"]), setFontsAll(11, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_signor2idname"]) != "")
                {
                    SignerClientAccount1 = new Chunk("\n" + Convert.ToString(table.Rows[l]["ssi_signor2idname"]), setFontsAll(11, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_signor3idname"]) != "")
                {
                    SignerClientAccount2 = new Chunk("\n" + Convert.ToString(table.Rows[l]["ssi_signor3idname"]), setFontsAll(11, 0, 0));
                }

                //Chunk lochunk = new Chunk("\n" + SignerClientAccount, setFontsAll(11, 1, 0));
                //loParagraph.Chunks.Add(lochunk);
                if (SignerClientAccount != null)
                {
                    loCell.Add(SignerClientAccount);
                }
                if (SignerClientAccount1 != null)
                {
                    loCell.Add(SignerClientAccount1);
                }
                if (SignerClientAccount2 != null)
                {
                    loCell.Add(SignerClientAccount2);
                }
                loCell.Colspan = table.Columns.Count;
                loCell.HorizontalAlignment = 1;
                loCell.Border = 0;


                if (Convert.ToString(table.Rows[l]["ssi_addressline1_mail"]) != "")
                {
                    Chunk lochunk1 = new Chunk("\n" + table.Rows[l]["ssi_addressline1_mail"].ToString(), setFontsAll(11, 0, 0));
                    loCell.Add(lochunk1);
                    loCell.Colspan = table.Columns.Count;
                    loCell.HorizontalAlignment = 1;
                    loCell.Border = 0;
                }

                if (Convert.ToString(table.Rows[l]["ssi_addressline2_mail"]) != "")
                {
                    Chunk lochunk2 = new Chunk("\n" + table.Rows[l]["ssi_addressline2_mail"].ToString(), setFontsAll(11, 0, 0));
                    loCell.Add(lochunk2);
                    loCell.Colspan = table.Columns.Count;
                    loCell.HorizontalAlignment = 1;
                    loCell.Border = 0;
                }

                if (Convert.ToString(table.Rows[l]["ssi_addressline3_mail"]) != "")
                {

                    Chunk lochunk3 = new Chunk("\n" + table.Rows[l]["ssi_addressline3_mail"].ToString(), setFontsAll(11, 0, 0));
                    loCell.Add(lochunk3);
                    loCell.Colspan = table.Columns.Count;
                    loCell.HorizontalAlignment = 1;
                    loCell.Border = 0;
                }


                Chunk lochunkAddress = null;
                Chunk lochunkAddress1 = null;
                Chunk lochunkAddress2 = null;
                Chunk lochunkAddress3 = null;
                if (Convert.ToString(table.Rows[l]["ssi_city_mail"]) != "")
                {
                    lochunkAddress = new Chunk("\n" + Convert.ToString(table.Rows[l]["ssi_city_mail"]), setFontsAll(11, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_stateprovince_mail"]) != "")
                {
                    lochunkAddress1 = new Chunk(", " + Convert.ToString(table.Rows[l]["ssi_stateprovince_mail"]), setFontsAll(11, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_zipcode_mail"]) != "")
                {
                    lochunkAddress2 = new Chunk(" " + Convert.ToString(table.Rows[l]["ssi_zipcode_mail"]), setFontsAll(11, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_countryregion_mail"]) != "")
                {
                    lochunkAddress3 = new Chunk("\n" + Convert.ToString(table.Rows[l]["ssi_countryregion_mail"]), setFontsAll(11, 0, 0));
                }

                //Chunk lochunkAddress = new Chunk("\n" + strAddress, setFontsAll(11, 0, 0));
                if (lochunkAddress != null)
                {
                    loCell.Add(lochunkAddress);
                }
                if (lochunkAddress1 != null)
                {
                    loCell.Add(lochunkAddress1);
                }
                if (lochunkAddress2 != null)
                {
                    loCell.Add(lochunkAddress2);
                }
                if (lochunkAddress3 != null)
                {
                    loCell.Add(lochunkAddress3);
                }
                loCell.Colspan = table.Columns.Count;
                loCell.HorizontalAlignment = 1;
                loCell.EnableBorderSide(2);

                #endregion

                #region As Of Date

                Chunk lochunkDate = null;
                Chunk lochunkDate1 = null;
                Chunk lochunkDate2 = null;


                if (Convert.ToString(table.Rows[l]["Ssi_LetterDate"]) != "")
                {
                    lochunkDate = new Chunk("\n" + Convert.ToString(table.Rows[l]["Ssi_LetterDate"]) + "\n", setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_capcallbankaddressee"]) != "")
                {
                    lochunkDate1 = new Chunk("\n" + Convert.ToString(table.Rows[l]["ssi_capcallbankaddressee"]), setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["bankaddress"]) != "")
                {
                    lochunkDate2 = new Chunk("\n" + Convert.ToString(table.Rows[l]["bankaddress"]), setFontsAll(9, 0, 0));
                }

                //Chunk lochunkDate = new Chunk("\n" + strAsOfDate, setFontsAll(11, 0, 0));

                if (lochunkDate != null)
                {
                    loCell1.Add(lochunkDate);
                }
                if (lochunkDate1 != null)
                {
                    loCell1.Add(lochunkDate1);
                }
                if (lochunkDate2 != null)
                {
                    loCell1.Add(lochunkDate2);
                }
                loCell1.Colspan = table.Columns.Count;
                loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loCell1.Border = 0;

                #endregion

                #region Account
                string strLegalEntity = "";
                string strAccountNumber = "";
                string Info = "";
                Chunk lochunkLegalEntity11 = null;
                Chunk lochunkLegalEntity = null;
                Chunk lochunkAccountNumber = null;
                Chunk loBankName = null;
                Chunk lochunkInfo = null;
                Chunk lochunkInfo1 = null;
                if (Convert.ToString(table.Rows[l]["ssi_legalentitynameidname"]) != "")
                {
                    lochunkLegalEntity11 = new Chunk("\nRE: Legal Entity - " + Convert.ToString(table.Rows[l]["ssi_legalentitynameidname"]), setFontsAll(9, 0, 0));
                    //Chunk LegalEntity1 = new Chunk(Convert.ToString(table.Rows[l]["ssi_accountname1"]), setFontsAll(11, 0, 0));

                }

                if (Convert.ToString(table.Rows[l]["ssi_accountname1"]) != "")
                {
                    lochunkLegalEntity = new Chunk("\n      Account - " + Convert.ToString(table.Rows[l]["ssi_accountname1"]), setFontsAll(9, 0, 0));
                    //Chunk LegalEntity1 = new Chunk(Convert.ToString(table.Rows[l]["ssi_accountname1"]), setFontsAll(11, 0, 0));

                }

                if (Convert.ToString(table.Rows[l]["ssi_bankidname"]) != "")
                {
                    loBankName = new Chunk("\n      " + Convert.ToString(table.Rows[l]["ssi_bankidname"]), setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_accountnumber"]) != "")
                {
                    lochunkAccountNumber = new Chunk("\n      " + Convert.ToString(table.Rows[l]["ssi_accountnumber"]), setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["bankdear"]) != "")
                {
                    lochunkInfo = new Chunk("\nDear " + Convert.ToString(table.Rows[l]["bankdear"]) + ":\n", setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_currentcall_ccsf"]) != "")
                {
                    lochunkInfo1 = new Chunk("\nUpon receipt of this letter, please wire transfer " + RoundUp(Convert.ToString(table.Rows[l]["ssi_currentcall_ccsf"])) + " on " + Convert.ToString(table.Rows[l]["Ssi_WireAsOfDate"]) + " from the account named above according to the following instructions:", setFontsAll(9, 0, 0));
                }

                if (lochunkLegalEntity11 != null)
                    loCell2.Add(lochunkLegalEntity11);
                if (lochunkLegalEntity != null)
                    loCell2.Add(lochunkLegalEntity);
                if (loBankName != null)
                    loCell2.Add(loBankName);
                if (lochunkAccountNumber != null)
                    loCell2.Add(lochunkAccountNumber);
                loCell2.Colspan = table.Columns.Count;
                loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loCell2.Border = 0;

                //loCell44.Add(loBankName);
                //loCell44.Colspan = table.Columns.Count;
                //loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                //loCell44.Border = 0;

                //loCell4.Add(lochunkAccountNumber);
                //loCell4.Colspan = table.Columns.Count;
                //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                //loCell4.Border = 0;


                if (lochunkInfo != null)
                {
                    loCell5.Add(lochunkInfo);
                }
                if (lochunkInfo1 != null)
                {
                    loCell5.Add(lochunkInfo1);
                }
                loCell5.Colspan = table.Columns.Count;
                loCell5.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loCell5.Border = 0;
                #endregion

                #region Fund

                int num = 1;
                //Chunk lochunkFundRemark = new Chunk("\n" + "Please process the following wire transfers on " + Convert.ToString(table.Rows[0]["As Of Date"]), setFontsAll(11, 0, 0));
                //Cell cellRemark = new Cell();
                //cellRemark.Add(lochunkFundRemark);
                //cellRemark.Colspan = table.Columns.Count;
                //cellRemark.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                //cellRemark.Border = 0;
                string BasicWireInfo = "";
                Chunk lochunkFundName = null;
                Chunk lochunkFundName1 = null;
                Chunk lochunkFundName2 = null;
                Chunk lochunkFundName3 = null;
                Chunk lochunkFundName4 = null;
                if (Convert.ToString(table.Rows[l]["ssi_basicwireinfo"]) != "")
                {
                    lochunkFundName = new Chunk("\n" + Convert.ToString(table.Rows[l]["ssi_basicwireinfo"]), setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["wirefundname"]) != "")
                {
                    lochunkFundName1 = new Chunk("\nFFC: " + Convert.ToString(table.Rows[l]["wirefundname"]), setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_account"]) != "")
                {
                    lochunkFundName2 = new Chunk("\nAccount # " + Convert.ToString(table.Rows[l]["ssi_account"]), setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_attnforwire"]) != "")
                {
                    lochunkFundName3 = new Chunk("\n" + Convert.ToString(table.Rows[l]["ssi_attnforwire"]), setFontsAll(9, 0, 0));
                }

                if (Convert.ToString(table.Rows[l]["ssi_legalentitynameidname"]) != "")
                {
                    lochunkFundName4 = new Chunk("\nFor the Benefit of: " + Convert.ToString(table.Rows[l]["ssi_legalentitynameidname"]), setFontsAll(9, 0, 0));
                }




                Paragraph p = new Paragraph();
                //p.Chunks.Add(lochunkFundName);
                if (lochunkFundName != null) { p.Add(lochunkFundName); }
                if (lochunkFundName1 != null) { p.Add(lochunkFundName1); }
                if (lochunkFundName2 != null) { p.Add(lochunkFundName2); }
                if (lochunkFundName3 != null) { p.Add(lochunkFundName3); }
                if (lochunkFundName4 != null) { p.Add(lochunkFundName4); }
                p.IndentationLeft = 75.0f;
                loCell3.Add(p);
                loCell3.Colspan = table.Columns.Count;
                //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                loCell3.Border = 0;

                #endregion

                #region Instructions
                Cell cellinstr1 = new Cell();

                string SignatureLine = "\n_________________________________\n";
                Chunk Instr = new Chunk("\nPlease call " + Convert.ToString(table.Rows[0]["ssi_secownerfname_mail"]) + " " + Convert.ToString(table.Rows[0]["ssi_secownerlname_mail"]) + " of Gresham Partners, LLC at (312)960-0200 if you have any questions.\n\nSincerely,\n", setFontsAll(9, 0, 0));

                Chunk signer = null;
                Chunk signer1 = null;
                Chunk signer2 = null;
                if (Convert.ToString(table.Rows[l]["ssi_signor1idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                {
                    if (Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                    {
                        signer = new Chunk(SignatureLine + Convert.ToString(table.Rows[l]["ssi_signor1idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer1title"]), setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        signer = new Chunk(SignatureLine + Convert.ToString(table.Rows[l]["ssi_signor1idname"]), setFontsAll(9, 0, 0));
                    }
                }

                if (Convert.ToString(table.Rows[l]["ssi_signor2idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                {
                    if (Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                    {
                        signer1 = new Chunk("\n" + SignatureLine + Convert.ToString(table.Rows[l]["ssi_signor2idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer2title"]), setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        signer1 = new Chunk("\n" + SignatureLine + Convert.ToString(table.Rows[l]["ssi_signor2idname"]), setFontsAll(9, 0, 0));
                    }
                }

                if (Convert.ToString(table.Rows[l]["ssi_signor3idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                {
                    if (Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                    {
                        signer2 = new Chunk("\n" + SignatureLine + Convert.ToString(table.Rows[l]["ssi_signor3idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer3title"]), setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        signer2 = new Chunk("\n" + SignatureLine + Convert.ToString(table.Rows[l]["ssi_signor3idname"]), setFontsAll(9, 0, 0));
                    }
                }

                //Chunk lochunkinstr1 = new Chunk("\n" + Instr + signer, setFontsAll(11, 0, 0));

                cellinstr1.Add(Instr);
                if (signer != null)
                {
                    cellinstr1.Add(signer);
                }
                if (signer1 != null)
                {
                    cellinstr1.Add(signer1);
                }
                if (signer2 != null)
                {
                    cellinstr1.Add(signer2);
                }
                cellinstr1.Colspan = table.Columns.Count;
                cellinstr1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                cellinstr1.Border = 0;



                #endregion

                #region Signer Client Account

                Chunk SignerClient1 = null;
                Chunk SignerClient2 = null;
                Chunk SignerClient3 = null;

                if (Convert.ToString(table.Rows[l]["ssi_signor1idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                {
                    if (Convert.ToString(table.Rows[l]["ssi_signer1title"]) != "")
                    {
                        SignerClient1 = new Chunk("_________________________\n" + "\n____\n" + Convert.ToString(table.Rows[l]["ssi_signor1idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer1title"]), setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        SignerClient1 = new Chunk("_________________________\n" + "\n____\n" + Convert.ToString(table.Rows[l]["ssi_signor1idname"]), setFontsAll(9, 0, 0));
                    }
                }

                if (Convert.ToString(table.Rows[l]["ssi_signor2idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                {
                    if (Convert.ToString(table.Rows[l]["ssi_signer2title"]) != "")
                    {
                        SignerClient2 = new Chunk("_________________________\n" + "\n____\n" + Convert.ToString(table.Rows[l]["ssi_signor2idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer2title"]), setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        SignerClient2 = new Chunk("_________________________\n" + "\n____\n" + Convert.ToString(table.Rows[l]["ssi_signor2idname"]), setFontsAll(9, 0, 0));
                    }
                }

                if (Convert.ToString(table.Rows[l]["ssi_signor3idname"]) != "" || Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                {
                    if (Convert.ToString(table.Rows[l]["ssi_signer3title"]) != "")
                    {
                        SignerClient3 = new Chunk("_________________________\n" + "\n____\n" + Convert.ToString(table.Rows[l]["ssi_signor3idname"]) + ", " + Convert.ToString(table.Rows[l]["ssi_signer3title"]), setFontsAll(9, 0, 0));
                    }
                    else
                    {
                        SignerClient3 = new Chunk("_________________________\n" + "\n____\n" + Convert.ToString(table.Rows[l]["ssi_signor3idname"]), setFontsAll(9, 0, 0));
                    }
                }

                if (SignerClient1 != null)
                {
                    //Chunk Signer1 = new Chunk("_________________________\n" + SignerClient1, setFontsAll(11, 0, 0));

                    Cellsigner1.Add(SignerClient1);
                    Cellsigner1.Colspan = table.Columns.Count - 8;
                    Cellsigner1.HorizontalAlignment = 1;
                    Cellsigner1.DisableBorderSide(1);
                }


                if (SignerClient2 != null)
                {
                    //Chunk Signer2 = new Chunk("\n" + SignerClient2, setFontsAll(11, 0, 0));

                    Cellsigner2.Add(SignerClient2);
                    Cellsigner2.Colspan = table.Columns.Count;
                    Cellsigner2.HorizontalAlignment = 1;
                    //Cellsigner2.EnableBorderSide(0);
                }


                if (SignerClient3 != null)
                {
                    //Chunk Signer3 = new Chunk("\n" + SignerClient3, setFontsAll(11, 0, 0));

                    Cellsigner3.Add(SignerClient2);
                    Cellsigner3.Colspan = table.Columns.Count;
                    Cellsigner3.HorizontalAlignment = 1;
                    //Cellsigner3.EnableBorderSide(0);
                }



                #endregion

                loTable.AddCell(loCell); //header
                loTable.AddCell(loCell1); // as of date
                loTable.AddCell(loCell2); // Lagal Entity
                //loTable.AddCell(loCell44); // Bank Name
                //loTable.AddCell(loCell4); // Lagal Entity
                loTable.AddCell(loCell5); // Lagal Entity
                //loTable.AddCell(cellRemark); // Fund Remark
                loTable.AddCell(loCell3); // Fund Details
                //loTable.AddCell(cellinstr); // Instructions
                loTable.AddCell(cellinstr1); // Instructions 1
                //if (SignerClient1 != "")
                //{
                //    loTable1.AddCell(Cellsigner1); // Signer1
                //}

                //if (SignerClient2 != "")
                //{
                //    loTable1.AddCell(Cellsigner2); // Signer2
                //}
                //if (SignerClient3 != "")
                //{
                //    loTable1.AddCell(Cellsigner3); // Signer3 
                //}

                document.Add(loTable);
            }

        }


        if (DSCount > 0)
        {
            document.Close();
            try
            {
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                return fsFinalLocation.Replace(".xls", ".pdf");
            }
            catch
            {

            }
        }
        else
        {
            return fsFinalLocation = "";
        }


        return fsFinalLocation.Replace(".xls", ".pdf");
    }
    #endregion

    #region Non-Standard Wire
    private string GetNonStandardWire()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.NonStandard);
        //String lsSQL = "SP_S_NonStandardWireLetter @MailID = 4, @legalentitynameid = '1470F8CF-AC32-DF11-B686-001D09665E8F', @ContactFullnameID = '8FA4752F-822E-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        int DSCount1 = newdataset.Tables[1].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string strGUID = System.DateTime.Now.ToString("ddMMMyyhhmmssss");
        Random rand = new Random();
        rand.Next();




        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        //AddFooter();

        document.Open();

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //png.SetAbsolutePosition(45, 800);//540
        ////png.ScaleToFit(288f, 42f);
        //png.ScalePercent(8);
        //document.Add(png);


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "NonStandard.xls";


        #region  Dataset 1

        if (newdataset.Tables.Count > 0)
        {
            if (newdataset.Tables[0].Rows.Count > 0)
            {
                iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, 2);   // 2 rows, 2 columns           
                loTable.Width = 30;
                loTable.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                iTextSharp.text.Cell locell = new Cell();


                string SignerClientAccount = string.Empty;


                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]) != "")
                {
                    Chunk lochunk = new Chunk("\n" + "Payment Instructions for " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]) + "\n\n", setFontsAll(11, 1, 0));
                    Paragraph p1 = new Paragraph();
                    p1.Add(lochunk);
                    document.Add(p1);
                }

                if (newdataset.Tables[0].Rows.Count > 0)
                {
                    string fontpath = HttpContext.Current.Server.MapPath(".");

                    string payment = "Payment via Check";
                    Chunk lochunk1 = new Chunk("\n" + payment, setFontsAll(11, 0, 0));
                    Chunk underline = new Chunk("\n_______________", setFontsAll(11, 0, 0));

                    locell.Width = 15;
                    locell.Leading = 2f;
                    locell.Add(lochunk1);
                    locell.Add(underline);
                    //locell1.EnableBorderSide(2);
                    //locell1.BorderWidthTop = 1f;
                    //locell1.BorderColorTop = iTextSharp.text.Color.BLACK;
                    locell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    locell.Colspan = 2;
                    locell.Border = 0;
                    loTable.Border = 0;

                    loTable.AddCell(locell);
                    //loTable.AddCell(locell1);
                    document.Add(loTable);
                    //Paragraph p2 = new Paragraph();
                    //p2.Add(lochunk1);
                    //document.Add(p2);

                    Chunk lochunk2 = new Chunk("\n" + "Please make your check payable to:", setFontsAll(11, 0, 0));
                    Paragraph p3 = new Paragraph();
                    p3.Add(lochunk2);
                    document.Add(p3);

                    Chunk FundDetails = null;
                    List list = new List(List.UNORDERED, 10f);
                    list.SetListSymbol("\u2022");
                    list.IndentationLeft = 30f;
                    for (int m = 0; m < DSCount; m++)
                    {
                        if (Convert.ToString(table.Rows[m]["ssi_fundname"]) != "" || Convert.ToString(table.Rows[m]["ssi_currentcall_ccsf"]) != "")
                        {
                            FundDetails = new Chunk(Convert.ToString(table.Rows[m]["ssi_fundname"]) + " for " + RoundUp(Convert.ToString(table.Rows[m]["ssi_currentcall_ccsf"])) + ".", setFontsAll(11, 0, 0));
                            list.Add(new iTextSharp.text.ListItem(FundDetails));
                        }
                    }
                    document.Add(list);

                    int w = 0;
                    string strLegalEntity = "";
                    Chunk lochunk4 = new Chunk("\n" + "Mail the check to following address, by " + Convert.ToString(table.Rows[w]["Ssi_WireAsOfDate"]) + ", to insure processing by due date: ", setFontsAll(11, 0, 0));
                    Paragraph p4 = new Paragraph();
                    p4.Add(lochunk4);
                    document.Add(p4);

                    //if (Convert.ToString(table.Rows[w]["ssi_legalentitynameidname"]) != "")
                    //{
                    //    strLegalEntity = "\n\n" + Convert.ToString(table.Rows[w]["ssi_legalentitynameidname"]);
                    //}

                    //if (Convert.ToString(table.Rows[w]["ssi_addressline1_mail"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n" + Convert.ToString(table.Rows[w]["ssi_addressline1_mail"]);
                    //}

                    //if (Convert.ToString(table.Rows[w]["ssi_addressline2_mail"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n" + Convert.ToString(table.Rows[w]["ssi_addressline2_mail"]);
                    //}

                    //if (Convert.ToString(table.Rows[w]["ssi_addressline3_mail"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n" + Convert.ToString(table.Rows[w]["ssi_addressline3_mail"]);
                    //}

                    //if (Convert.ToString(table.Rows[w]["ssi_city_mail"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n" + Convert.ToString(table.Rows[w]["ssi_city_mail"]);
                    //}

                    //if (Convert.ToString(table.Rows[w]["ssi_stateprovince_mail"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n" + Convert.ToString(table.Rows[w]["ssi_stateprovince_mail"]);
                    //}

                    //if (Convert.ToString(table.Rows[w]["ssi_zipcode_mail"]) != "")
                    //{
                    //    strLegalEntity = strLegalEntity + "\n" + Convert.ToString(table.Rows[w]["ssi_zipcode_mail"]);
                    //}


                    strLegalEntity = "\n\n Kate Warner \n Gresham Partners, LLC \n 333 West Wacker Drive \n Suite 700 \n Chicago,IL 60606";

                    Chunk lochunk5 = new Chunk(strLegalEntity, setFontsAll(11, 0, 0));

                    Paragraph p5 = new Paragraph();
                    p5.Add(lochunk5);
                    p5.IndentationLeft = 75.0f;
                    document.Add(p5);
                }


            }
        }




        #endregion

        #region Dataset 2

        if (newdataset.Tables.Count > 0)
        {
            if (newdataset.Tables[1].Rows.Count > 0)
            {
                iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, 2);   // 2 rows, 2 columns           
                loTable.Width = 30;
                loTable.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                iTextSharp.text.Cell locell = new Cell();

                if (Convert.ToString(newdataset.Tables[1].Rows[0]["ssi_legalentitynameidname_1"]) != "")
                {
                    Chunk lochunk = new Chunk("\n" + Convert.ToString(newdataset.Tables[1].Rows[0]["ssi_legalentitynameidname_1"]) + "\n\n", setFontsAll(11, 1, 0));

                    Paragraph p2 = new Paragraph();
                    p2.Add(lochunk);
                    document.Add(p2);
                }
                else
                {
                    Chunk lochunk111 = new Chunk("\n\n", setFontsAll(9, 0, 0));
                    Paragraph p22 = new Paragraph();
                    p22.Add(lochunk111);
                    document.Add(p22);
                }


                DataTable table1 = newdataset.Tables[1].Copy();


                string payment1 = "Payment via Wire Transfer";
                Chunk lochunk11 = new Chunk("\n" + payment1, setFontsAll(11, 0, 0));
                //lochunk11.SetUnderline(1, -2);
                //Chunk lochunk1 = new Chunk("\n" + payment, setFontsAll(11, 0, 0));
                Chunk underline = new Chunk("\n_____________________", setFontsAll(11, 0, 0));

                locell.Width = 15;
                locell.Leading = 2f;
                locell.Add(lochunk11);
                locell.Add(underline);
                //locell1.EnableBorderSide(2);
                //locell1.BorderWidthTop = 1f;
                //locell1.BorderColorTop = iTextSharp.text.Color.BLACK;
                locell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                locell.Colspan = 2;
                locell.Border = 0;
                loTable.Border = 0;

                loTable.AddCell(locell);
                //loTable.AddCell(locell1);
                document.Add(loTable);

                Chunk lochunk12 = new Chunk("\n" + "Please use the following instructions:\n\n", setFontsAll(11, 0, 0));
                Paragraph p7 = new Paragraph();
                p7.Add(lochunk12);
                document.Add(p7);

                string FundDetails11 = "";
                string FundDetails111 = "";
                string FundDetails1 = "";
                string strFFC = "";
                string FundDetails2 = "";
                string FundName = "";
                Chunk loFundCheck = null;
                Chunk loFundName = null;
                Chunk loFundCheck1 = null;
                for (int n = 0; n < DSCount1; n++)
                {
                    if (Convert.ToString(table1.Rows[n]["FundFullName"]) != "" || Convert.ToString(table1.Rows[n]["ssi_currentcall_ccsf"]) != "")
                    {
                        if (n == 0)
                        {
                            loFundCheck = new Chunk("Wire ", setFontsAll(11, 0, 0));
                        }
                        else
                        {
                            loFundCheck = new Chunk("\n" + "Wire ", setFontsAll(11, 0, 0));
                        }
                        // loFundCheck = new Chunk("Wire ", setFontsAll(11, 0, 0));
                        loFundName = new Chunk(RoundUp(Convert.ToString(table1.Rows[n]["ssi_currentcall_ccsf"])), setFontsAll(11, 1, 0));
                        loFundCheck1 = new Chunk(" for " + Convert.ToString(table1.Rows[n]["FundFullName"]) + " to:", setFontsAll(11, 0, 0));
                    }


                    if (Convert.ToString(table1.Rows[n]["ssi_basicwireinfo"]) != "")
                    {
                        FundDetails1 = Convert.ToString(table1.Rows[n]["ssi_basicwireinfo"]);
                    }


                    if (Convert.ToString(table1.Rows[n]["FFC"]) != "")
                    {
                        strFFC = "FFC: " + Convert.ToString(table1.Rows[n]["FFC"]);
                    }

                    if (Convert.ToString(table1.Rows[n]["ssi_accountnumber"]) != "")
                    {
                        FundDetails2 = "Account # " + Convert.ToString(table1.Rows[n]["ssi_accountnumber"]);
                    }

                    if (Convert.ToString(table1.Rows[n]["ssi_Attnforwire"]) != "")
                    {
                        FundDetails2 = FundDetails2 + "\n" + Convert.ToString(table1.Rows[n]["ssi_Attnforwire"]) + "\n\n\n";
                    }

                    Chunk lochunkFundName1 = null;
                    Chunk lochunkFundName3 = null;
                    Chunk FFC = null;



                    if (FundDetails1 != "")
                    {
                        lochunkFundName1 = new Chunk("\n" + FundDetails1, setFontsAll(11, 0, 0));
                    }
                    if (strFFC != "")
                    {
                        FFC = new Chunk("\n" + strFFC, setFontsAll(11, 0, 0));
                    }
                    if (FundDetails2 != "")
                    {
                        lochunkFundName3 = new Chunk("\n" + FundDetails2, setFontsAll(11, 0, 0));
                    }


                    Paragraph p8 = new Paragraph();

                    if (loFundCheck != null)
                    { p8.Add(loFundCheck); }
                    if (loFundName != null)
                    { p8.Add(loFundName); }
                    if (loFundCheck1 != null)
                    { p8.Add(loFundCheck1); }


                    p8.Add(lochunkFundName1);
                    p8.Add(FFC);
                    p8.Add(lochunkFundName3);
                    document.Add(p8);

                }
            }


        }

        #endregion


        if (DSCount > 0 || DSCount1 > 0)
        {
            document.Close();
            try
            {
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch { }

        }
        else
        {
            fsFinalLocation = "";
        }

        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    #endregion

    #endregion

    #region Capital Call Statement

    public string GetCapitalCallStatement()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.CapitalCallStatement);
        //String lsSQL = "SP_S_CapitalCallStatements @MailID = 2, @legalentitynameid = '6055D1AA-6E15-DE11-8391-001D09665E8F',@ContactFullnameID = '4F073302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "CapitalCallStatement.pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        //AddFooter();

        document.Open();

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //png.SetAbsolutePosition(45, 800);//540
        ////png.ScaleToFit(288f, 42f);
        //png.ScalePercent(8);
        //document.Add(png);

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        try
        {
            if (DSCount > 0)
            {
                for (int i = 0; i < DSCount; i++)
                {
                    if (i != 0)
                    {
                        document.NewPage();
                    }
                    string fundname = Convert.ToString(table.Rows[i]["ssi_fundname"]);
                    Chunk lochunk1 = new Chunk("\n\n" + fundname.ToUpper(), setFontsAllFrutiger(12, 1, 0));

                    string AsOfDate = "Statement as of " + Convert.ToString(table.Rows[i]["Ssi_WireAsOfDate"]);
                    Chunk lochunk2 = new Chunk("\n" + AsOfDate, setFontsAllFrutiger(10, 1, 0));

                    string Legalentityname = Convert.ToString(table.Rows[i]["ssi_legalentitynameidname"]);
                    Chunk lochunk3 = new Chunk("\n\n" + Legalentityname + "\n\n\n", setFontsAllFrutiger(11, 1, 0));

                    Paragraph p1 = new Paragraph();
                    p1.Add(lochunk1);
                    p1.Add(lochunk2);
                    p1.Add(lochunk3);
                    p1.Alignment = 1;
                    document.Add(p1);

                    iTextSharp.text.Table loTable = new iTextSharp.text.Table(4, 10);   // 2 rows, 2 columns           
                    lsTotalNumberofColumns = "4";
                    iTextSharp.text.Cell loCell = new Cell();
                    setTableProperty(loTable, ReportType.CapitalCallStatement);
                    iTextSharp.text.Chunk lochunk = new Chunk();

                    int rowsize = 10;
                    int colsize = 4;
                    for (int j = 0; j < rowsize; j++)
                    {
                        for (int k = 0; k < colsize; k++)
                        {
                            if (j == 0 && k == 0)
                            {
                                lochunk = new Chunk("Capital Call", setFontsAll(10, 1, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Colspan = 2;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("Total", setFontsAll(10, 1, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Colspan = 2;
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                                loTable.AddCell(loCell);
                            }
                            if (j == 1 && k == 0)
                            {
                                lochunk = new Chunk("Current Capital Call - " + Convert.ToString(table.Rows[i]["Ssi_WireAsOfDate"]) + "", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Colspan = 4;
                                k = k + 3;

                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);
                            }
                            if (j == 2 && k == 0)
                            {
                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["percentcalled_ccsf"])) + " of Capital Commitment", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Colspan = 2;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["currentcall_ccsf"])), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Colspan = 2;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 3 && k == 0)
                            {
                                lochunk = new Chunk("\n% of      \nCommitted\n       Capital", setFontsAll(10, 1, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);

                                loCell.Colspan = 4;
                                k = k + 3;
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);
                            }
                            if (j == 4 && k == 0)
                            {
                                lochunk = new Chunk("Status of Commitment after Capital Call:", setFontsAll(10, 1, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);

                                loCell.Colspan = 4;
                                k = k + 3;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;
                                loTable.AddCell(loCell);
                            }
                            if (j == 5 && k == 0)
                            {
                                lochunk = new Chunk("Capital Commitment", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("$", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["ttladjcommit_ccsf"])).Replace("$", ""), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("100.00%", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 6 && k == 0)
                            {
                                lochunk = new Chunk("Current Capital Call", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["currentcall_ccsf"])).Replace("$", ""), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["percentcalled_ccsf"])), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 7 && k == 0)
                            {
                                lochunk = new Chunk("Prior Calls", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.BorderWidthBottom = 1F;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 1F;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["priorcalls_ccsf"])).Replace("$", ""), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 1F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["percentpriorcalls_ccsf"])), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 1F;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 8 && k == 0)
                            {
                                lochunk = new Chunk("Called to Date", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["calledtodate_ccsf"])).Replace("$", ""), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["calledtodateP_ccsf"])), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 9 && k == 0)
                            {
                                lochunk = new Chunk("Remaining Commitment", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("$", setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["remainingcommitment_ccsf"])).Replace("$", ""), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 2F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["remainingcommitmentP_ccsf"])), setFontsAll(10, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                        }
                    }
                    document.Add(loTable);
                }
            }
        }
        catch
        {
        }



        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    #endregion


    #region SLOA
    public string GetSLOA()
    {
        Random rand = new Random();
        rand.Next();
        string FilePath = string.Empty;
        string DestinationFileName = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + DateTime.Now.ToString("ddMMMyyhhmmss") + System.Guid.NewGuid().ToString() + "SLOA.pdf";
        string[] SourceFileName = new string[2];

        SourceFileName[0] = GetSLOA1();
        //SourceFileName[1] = GetSLOA2();
        SourceFileName[1] = GetSLOA3();

        for (int i = 0; i < SourceFileName.Length; i++)
        {
            if (FilePath != "")
            {
                if (SourceFileName[i] != "")
                {
                    FilePath = FilePath + "|" + SourceFileName[i];
                }
            }
            else
            {
                if (SourceFileName[i] != "")
                {
                    FilePath = "|" + SourceFileName[i];
                }
            }
        }


        if (FilePath != "")
        {
            FilePath = FilePath.Substring(1, FilePath.Length - 1);
            string[] strPath = FilePath.Split('|');

            PDFMerge PDF = new PDFMerge();
            //   PDF.MergeFiles(DestinationFileName, strPath);
            PDF.MergeNew(DestinationFileName, strPath);
        }
        else
        {
            DestinationFileName = "";
        }

        return DestinationFileName;
    }

    public string GetSLOA1()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.SLOA);


        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        //pdfTemplate -- With TextFields to Write values
        string pdfTemplate = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\Fidelity-SLOA.pdf";

        string newFile = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "-SLOA1.pdf";

        PdfReader pdfReader = new PdfReader(pdfTemplate);
        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                    newFile, FileMode.Create));

        AcroFields pdfFormFields = pdfStamper.AcroFields;

        // set form pdfFormFields
        string LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));
        string FidelityAcc = Convert.ToString(table.Rows[0]["ssi_accountnumber"]);
        string FidelityAcc2 = null;
        if (FidelityAcc != "")
        {
            FidelityAcc2 = FidelityAcc.Replace("-", "");
            int accountNoChars = FidelityAcc2.Length;

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < accountNoChars; i++)
            {
                sb.Append(FidelityAcc2[i] + " ");
            }
            string FidelityAcc3 = sb.ToString();
            string[] strAccNo = FidelityAcc3.Split(' ');
            if (accountNoChars >= 1)
            {
                pdfFormFields.SetField("AccNum1", strAccNo[0]);
                if (accountNoChars >= 2)
                {
                    pdfFormFields.SetField("AccNum2", strAccNo[1]);
                    if (accountNoChars >= 3)
                    {
                        pdfFormFields.SetField("AccNum3", strAccNo[2]);
                        if (accountNoChars >= 4)
                        {
                            pdfFormFields.SetField("AccNum4", strAccNo[3]);
                            if (accountNoChars >= 5)
                            {
                                pdfFormFields.SetField("AccNum5", strAccNo[4]);
                                if (accountNoChars >= 6)
                                {
                                    pdfFormFields.SetField("AccNum6", strAccNo[5]);
                                    if (accountNoChars >= 7)
                                    {
                                        pdfFormFields.SetField("AccNum7", strAccNo[6]);
                                        if (accountNoChars >= 8)
                                        {
                                            pdfFormFields.SetField("AccNum8", strAccNo[7]);
                                            if (accountNoChars >= 9)
                                            {
                                                pdfFormFields.SetField("AccNum9", strAccNo[8]);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        //string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        //string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        //string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);

        // The first worksheet
        //    pdfFormFields.SetField("LEName", LEname);

        string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);
        string AccFirstName = Convert.ToString(table.Rows[0]["Signor1FirstName"]);
        string AccLastName = Convert.ToString(table.Rows[0]["Signor1LastName"]);
        string SLOA_LE_Name = Convert.ToString(table.Rows[0]["SLOALegalEntityName"]);

        // The first worksheet
        //pdfFormFields.SetField("LEName", LEname);
        //pdfFormFields.SetField("AccountNumber", FidelityAcc);
        pdfFormFields.SetField("Signer1", strSigner1);
        pdfFormFields.SetField("Signer2", strSigner2);
        pdfFormFields.SetField("Signer3", strSigner3);
        //  pdfFormFields.SetField("AccFirstName", AccFirstName);
        pdfFormFields.SetField("AccFirstName", LEname);
        pdfFormFields.SetField("AccLastName", AccLastName);
        pdfFormFields.SetField("SLOA_LE_Name", SLOA_LE_Name);

        //pdfFormFields.SetField("AccountNumber", FidelityAcc);
        //pdfFormFields.SetField("Signer1", strSigner1);
        //pdfFormFields.SetField("Signer2", strSigner2);
        //pdfFormFields.SetField("Signer3", strSigner3);

        // flatten the form to remove editting options, set it to false
        // to leave the form open to subsequent manual edits
        pdfStamper.FormFlattening = true;
        pdfFormFields.GenerateAppearances = true;
        // close the pdf
        pdfStamper.Close();
        return newFile;
    }

    public string GetSLOA2()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.SLOA);


        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        //pdfTemplate -- With TextFields to Write values
        string pdfTemplate = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\Fidelity-SLOA2.pdf";

        string newFile = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "-SLOA2.pdf";

        PdfReader pdfReader = new PdfReader(pdfTemplate);
        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                    newFile, FileMode.Create));

        AcroFields pdfFormFields = pdfStamper.AcroFields;

        // set form pdfFormFields
        //string LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));
        //string FidelityAcc = Convert.ToString(table.Rows[0]["ssi_accountnumber"]);
        string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);

        // The first worksheet
        //pdfFormFields.SetField("LEName", LEname);
        //pdfFormFields.SetField("AccountNumber", FidelityAcc);
        pdfFormFields.SetField("Signer1", strSigner1);
        pdfFormFields.SetField("Signer2", strSigner2);
        pdfFormFields.SetField("Signer3", strSigner3);

        // flatten the form to remove editting options, set it to false
        // to leave the form open to subsequent manual edits
        pdfStamper.FormFlattening = true;
        pdfFormFields.GenerateAppearances = true;
        // close the pdf
        pdfStamper.Close();
        return newFile;
    }


    public string GetSLOA3()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.SLOA);


        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "SLOA.pdf";
        // iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        //AddHeader(document);
        //AddFooter();

        document.Open();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        string FidelityAcc = string.Empty;
        string LENameNEW = "";
        string LENamOLD = "";

        //Creating Table for Heading with two column 
        PdfPTable LoHeader = new PdfPTable(1);

        int[] width12 = { 100 };
        LoHeader.SetWidths(width12);

        //Defining Width 100%
        LoHeader.WidthPercentage = 100f;

        FidelityAcc = Convert.ToString(table.Rows[0]["ssi_accountnumber"]);

        PdfPCell LEName = new PdfPCell();

        Paragraph HeadingLEtName = new Paragraph(LEname, setFontsAllTimesNewRoman(14, 1, 0));
        Paragraph HeadingAccount = new Paragraph("Fidelity Account #" + FidelityAcc, setFontsAllTimesNewRoman(14, 1, 0));

        LEName.AddElement(HeadingLEtName);
        LEName.AddElement(HeadingAccount);
        LEName.Border = 0;


        string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);


        PdfPTable LoFooter = new PdfPTable(3);
        int[] width1 = { 35, 35, 35 };
        LoFooter.SetWidths(width1);

        LoFooter.WidthPercentage = 100f;

        PdfPCell CellFooter11 = new PdfPCell();
        PdfPCell CellFooter12 = new PdfPCell();
        PdfPCell CellFooter13 = new PdfPCell();

        PdfPCell CellFooter21 = new PdfPCell();
        PdfPCell CellFooter22 = new PdfPCell();
        PdfPCell CellFooter23 = new PdfPCell();

        PdfPCell CellFooter31 = new PdfPCell();
        PdfPCell CellFooter32 = new PdfPCell();
        PdfPCell CellFooter33 = new PdfPCell();

        PdfPCell CellFooter1 = new PdfPCell();
        PdfPCell CellFooter2 = new PdfPCell();
        PdfPCell CellFooter3 = new PdfPCell();


        Paragraph PBlank = new Paragraph(" ", setFontsAllTimesNewRoman(10, 1, 0));
        Paragraph PSigner1 = new Paragraph(strSigner1, setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PSigner2 = new Paragraph(strSigner2, setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PSigner3 = new Paragraph(strSigner3, setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PAccountHolderSignature = new Paragraph("Account Holder Signature", setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PPrintedName = new Paragraph("Printed Name", setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PDate = new Paragraph("Date", setFontsAllTimesNewRoman(11, 0, 0));

        PAccountHolderSignature.SetAlignment("center");
        PPrintedName.SetAlignment("center");
        PDate.SetAlignment("center");

        PAccountHolderSignature.Leading = 8f;
        PPrintedName.Leading = 8f;
        PDate.Leading = 8f;


        PSigner1.SetAlignment("center");
        PSigner2.SetAlignment("center");
        PSigner3.SetAlignment("center");

        CellFooter11.Border = 0;
        CellFooter12.Border = 0;
        CellFooter13.Border = 0;

        CellFooter21.Border = 0;
        CellFooter22.Border = 0;
        CellFooter23.Border = 0;

        CellFooter31.Border = 0;
        CellFooter32.Border = 0;
        CellFooter33.Border = 0;

        CellFooter1.Border = 0;
        CellFooter2.Border = 0;
        CellFooter3.Border = 0;

        CellFooter1.BorderWidthTop = 1f;
        CellFooter2.BorderWidthTop = 1f;
        CellFooter3.BorderWidthTop = 1f;


        CellFooter11.AddElement(PBlank);
        CellFooter12.AddElement(PSigner1);
        CellFooter13.AddElement(PBlank);

        CellFooter11.PaddingTop = 10f;
        CellFooter12.PaddingTop = 10f;
        CellFooter13.PaddingTop = 10f;

        LoFooter.AddCell(CellFooter11);
        LoFooter.AddCell(CellFooter12);
        LoFooter.AddCell(CellFooter13);

        CellFooter1.AddElement(PAccountHolderSignature);
        CellFooter2.AddElement(PPrintedName);
        CellFooter3.AddElement(PDate);

        LoFooter.AddCell(CellFooter1);
        LoFooter.AddCell(CellFooter2);
        LoFooter.AddCell(CellFooter3);

        CellFooter21.AddElement(PBlank);
        CellFooter22.AddElement(PSigner2);
        CellFooter23.AddElement(PBlank);

        CellFooter21.PaddingTop = 10f;
        CellFooter22.PaddingTop = 10f;
        CellFooter23.PaddingTop = 10f;

        LoFooter.AddCell(CellFooter21);
        LoFooter.AddCell(CellFooter22);
        LoFooter.AddCell(CellFooter23);

        LoFooter.AddCell(CellFooter1);
        LoFooter.AddCell(CellFooter2);
        LoFooter.AddCell(CellFooter3);

        CellFooter31.AddElement(PBlank);
        CellFooter32.AddElement(PSigner3);
        CellFooter33.AddElement(PBlank);

        CellFooter31.PaddingTop = 10f;
        CellFooter32.PaddingTop = 10f;
        CellFooter33.PaddingTop = 10f;

        LoFooter.AddCell(CellFooter31);
        LoFooter.AddCell(CellFooter32);
        LoFooter.AddCell(CellFooter33);

        LoFooter.AddCell(CellFooter1);
        LoFooter.AddCell(CellFooter2);
        LoFooter.AddCell(CellFooter3);


        LoFooter.TotalWidth = 510;

        LoHeader.AddCell(LEName);

        //   document.Add(LoFooter);

        //Page Size Capacity for Portfolio Perf report
        int PageSize1 = 6;
        //Calculate How much requires to display records
        int TotalPage = (table.Rows.Count / PageSize1);

        //Check If it requires more page
        if (table.Rows.Count % PageSize1 != 0)
        {
            TotalPage = TotalPage + 1;
        }

        //Calulates and display records with Header and footer
        for (int CurPage = 1; CurPage <= TotalPage; CurPage++)
        {
            //New Table 
            PdfPTable ptab1 = new PdfPTable(1);
            ptab1.WidthPercentage = 100f;
            //New Cell
            PdfPCell loCellSLOA = new PdfPCell();

            //New table - Temp 
            PdfPTable Ptab;

            //Add New Page 
            document.NewPage();

            //add Header starts from 2nd page
            document.Add(LoHeader);

            Ptab = BindDynamicSLOAData(CurPage, TotalPage, PageSize1, table);

            Ptab.WidthPercentage = 100f;

            loCellSLOA.AddElement(Ptab);

            //No border
            loCellSLOA.Border = 0;
            loCellSLOA.PaddingLeft = 0f;
            loCellSLOA.PaddingRight = 0f;

            ptab1.AddCell(loCellSLOA);

            document.Add(ptab1);
            LoFooter.WriteSelectedRows(0, -1, 35, 140, writer.DirectContent);

            // document.Add(LoFooter);
        }


        document.Close();
        FileInfo loFile = new FileInfo(ls);
        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

        return fsFinalLocation.Replace(".xls", ".pdf");

    }

    public PdfPTable BindDynamicSLOAData(int CurrentPage, int TotalPage, int PageSize, DataTable table)
    {
        PdfPTable LoData = new PdfPTable(2);


        int[] width2 = { 50, 50 };
        LoData.SetWidths(width2);


        LoData.WidthPercentage = 100f;

        int DSCount = table.Rows.Count;

        //Calculate End Value of Grid.
        //i.e End Value  = Current Page Number * Page Size Capacity
        int finalVal = (CurrentPage * PageSize);

        //Calculate Starting Value of Grid.
        //i.e Starting Value  = End Value - Page Size Capacity
        int intialVal = finalVal - PageSize;

        //As Values in Grid Starting from Zero, End Value will Decreased by 1
        finalVal = finalVal - 1;

        if (CurrentPage == TotalPage)
            finalVal = table.Rows.Count - 1;


        for (int x = intialVal; x <= finalVal; x++)
        {
            PdfPCell CellData = new PdfPCell();
            string strPartnership = Convert.ToString(table.Rows[x]["ssi_fundname"]);
            string strABA = Convert.ToString(table.Rows[x]["Ssi_ABARoutingin"]);
            // string strCustName = Convert.ToString(table.Rows[x]["CustodianName"]);
            string strCustName = "";
            string strPartAcc = Convert.ToString(table.Rows[x]["ssi_accountnumber"]);
            string strOtherWireInst = Convert.ToString(table.Rows[x]["OtherWireInstr"]);
            string DDANumber = Convert.ToString(table.Rows[x]["Ssi_DDAoutgoing"]);
            string LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));
            string strBasicWireInfo = Convert.ToString(table.Rows[x]["BasicWireInfo"]);
            string strbeneficiaryAccount = Convert.ToString(table.Rows[x]["ssi_beneficiaryAccount"]);

            Paragraph PPartnership = new Paragraph(strPartnership, setFontsAllTimesNewRoman(10, 1, 0));


            //ABA
            Paragraph PABA = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PABAText = new Chunk("ABA #: ", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PABAVal = new Chunk(strABA, setFontsAllTimesNewRoman(10, 1, 0));

            //Bank
            Paragraph PBank = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PBankText = new Chunk("Bank: ", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PBankCustName = new Chunk(strCustName, setFontsAllTimesNewRoman(10, 0, 0));

            //Other info
            //Paragraph POtherInfo = new Paragraph("Other Info*", setFontsAllTimesNewRoman(10, 0, 0));
            Paragraph POtherInfo = new Paragraph(strBasicWireInfo, setFontsAllTimesNewRoman(10, 0, 0));

            //Beneficiary            
            Paragraph PBeneficiaryText = new Paragraph("Beneficiary: ", setFontsAllTimesNewRoman(10, 1, 0));
            Paragraph PBeneficiaryVal = new Paragraph(strPartnership, setFontsAllTimesNewRoman(10, 0, 0));

            //Bene Acct
            Paragraph PBeneAcct = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PBeneAcctText = new Chunk("Bene Acct #: ", setFontsAllTimesNewRoman(10, 1, 0));
            //Chunk PBeneAcctVal = new Chunk("Partnership Acct #", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PBeneAcctVal = new Chunk(strbeneficiaryAccount, setFontsAllTimesNewRoman(10, 0, 0));
            //Chunk PBeneAcctVal = new Chunk(strPartAcc, setFontsAllTimesNewRoman(10, 0, 0));

            //Details
            Paragraph PDetails = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PDetailsText = new Chunk("Details: REF: ", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PDetailsTVal = new Chunk(LEname, setFontsAllTimesNewRoman(10, 0, 0));

            Paragraph POtherWireInstVal = new Paragraph(strOtherWireInst, setFontsAllTimesNewRoman(10, 0, 0));

            //DDA
            Paragraph PDDA = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PDDAText = new Chunk("DDA #:", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PDDAVal = new Chunk(DDANumber, setFontsAllTimesNewRoman(10, 0, 0));

            PABA.Add(PABAText);
            PABA.Add(PABAVal);

            PBank.Add(PBankText);
            PBank.Add(PBankCustName);

            PBeneAcct.Add(PBeneAcctText);
            PBeneAcct.Add(PBeneAcctVal);

            PDetails.Add(PDetailsText);
            PDetails.Add(PDetailsTVal);

            PDDA.Add(PDDAText);
            PDDA.Add(PDDAVal);



            CellData.AddElement(PPartnership);
            CellData.AddElement(PABA);
            CellData.AddElement(PBank);
            CellData.AddElement(POtherInfo);

            CellData.AddElement(PBeneficiaryText);
            CellData.AddElement(PBeneficiaryVal);
            CellData.AddElement(PBeneAcct);
            CellData.AddElement(PDetails);

            CellData.AddElement(POtherWireInstVal);
            if (!string.IsNullOrEmpty(DDANumber))
                CellData.AddElement(PDDA);

            CellData.Border = 0;
            CellData.PaddingTop = 12f;
            LoData.AddCell(CellData);
        }

        if (DSCount % 2 != 0)
        {
            PdfPCell Test1 = new PdfPCell();
            Paragraph Test2 = new Paragraph(" ", setFontsAllTimesNewRoman(10, 0, 0));
            Test1.Border = 0;
            Test1.AddElement(Test2);
            Test1.PaddingTop = 12f;
            LoData.AddCell(Test1);
        }
        LoData.SpacingBefore = 10f;
        return LoData;
    }
    #endregion

    #region SLOA Pdf
    public string Get_SLOA(DataTable lodataset, string FolderName, string FileName)
    {
        Random rand = new Random();
        rand.Next();
        string FilePath = string.Empty;
        //  string DestinationFileName = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + DateTime.Now.ToString("ddMMMyyhhmmss") + System.Guid.NewGuid().ToString() + "SLOA.pdf";

        string DestinationFileName = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FolderName + "\\" + FileName;

        string[] SourceFileName = new string[2];

        //  SourceFileName[0] = Get_SLOA1(lodataset);

        SourceFileName[0] = Get_SLOA1(lodataset);
        //SourceFileName[1] = GetSLOA2();
        // SourceFileName[1] = Get_SLOA3(lodataset);


        SourceFileName[1] = Get_SLOA3(lodataset);

        for (int i = 0; i < SourceFileName.Length; i++)
        {
            if (FilePath != "")
            {
                if (SourceFileName[i] != "")
                {
                    FilePath = FilePath + "|" + SourceFileName[i];
                }
            }
            else
            {
                if (SourceFileName[i] != "")
                {
                    FilePath = "|" + SourceFileName[i];
                }
            }
        }


        if (FilePath != "")
        {
            FilePath = FilePath.Substring(1, FilePath.Length - 1);
            string[] strPath = FilePath.Split('|');

            PDFMerge PDF = new PDFMerge();
            //   PDF.MergeFiles(DestinationFileName, strPath);
            PDF.MergeNew(DestinationFileName, strPath);
        }
        else
        {
            DestinationFileName = "";
        }

        return DestinationFileName;
    }

    public string Get_SLOA1(DataTable loDataSet)
    {
        DataTable newdataset;
        DB clsDB = new DB();
        newdataset = null;
        //  String lsSQL = getFinalSp(ReportType.SLOA);


        newdataset = loDataSet.Copy();
        newdataset.AcceptChanges();
        // int DSCount = newdataset.Tables[0].Rows.Count;
        int DSCount = newdataset.Rows.Count;
        DataTable table = newdataset.Copy();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        //pdfTemplate -- With TextFields to Write values
        string pdfTemplate = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\Fidelity-SLOA.pdf";

        string newFile = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "-SLOA1.pdf";

        //  string newFile = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FolderName + "\\" + FileName;



        PdfReader pdfReader = new PdfReader(pdfTemplate);
        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                    newFile, FileMode.Create));

        AcroFields pdfFormFields = pdfStamper.AcroFields;

        // set form pdfFormFields
        string LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));
        string FidelityAcc = Convert.ToString(table.Rows[0]["ssi_accountnumber"]);
        string FidelityAcc2 = null;
        if (FidelityAcc != "")
        {
            FidelityAcc2 = FidelityAcc.Replace("-", "");
            int accountNoChars = FidelityAcc2.Length;

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < accountNoChars; i++)
            {
                sb.Append(FidelityAcc2[i] + " ");
            }
            string FidelityAcc3 = sb.ToString();
            string[] strAccNo = FidelityAcc3.Split(' ');
            if (accountNoChars >= 1)
            {
                pdfFormFields.SetField("AccNum1", strAccNo[0]);
                if (accountNoChars >= 2)
                {
                    pdfFormFields.SetField("AccNum2", strAccNo[1]);
                    if (accountNoChars >= 3)
                    {
                        pdfFormFields.SetField("AccNum3", strAccNo[2]);
                        if (accountNoChars >= 4)
                        {
                            pdfFormFields.SetField("AccNum4", strAccNo[3]);
                            if (accountNoChars >= 5)
                            {
                                pdfFormFields.SetField("AccNum5", strAccNo[4]);
                                if (accountNoChars >= 6)
                                {
                                    pdfFormFields.SetField("AccNum6", strAccNo[5]);
                                    if (accountNoChars >= 7)
                                    {
                                        pdfFormFields.SetField("AccNum7", strAccNo[6]);
                                        if (accountNoChars >= 8)
                                        {
                                            pdfFormFields.SetField("AccNum8", strAccNo[7]);
                                            if (accountNoChars >= 9)
                                            {
                                                pdfFormFields.SetField("AccNum9", strAccNo[8]);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        //string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        //string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        //string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);

        // The first worksheet
        //    pdfFormFields.SetField("LEName", LEname);

        string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);
        string AccFirstName = Convert.ToString(table.Rows[0]["Signor1FirstName"]);
        string AccLastName = Convert.ToString(table.Rows[0]["Signor1LastName"]);
        string SLOA_LE_Name = Convert.ToString(table.Rows[0]["SLOALegalEntityName"]);

        // The first worksheet
        //pdfFormFields.SetField("LEName", LEname);
        //pdfFormFields.SetField("AccountNumber", FidelityAcc);
        pdfFormFields.SetField("Signer1", strSigner1);
        pdfFormFields.SetField("Signer2", strSigner2);
        pdfFormFields.SetField("Signer3", strSigner3);
        //  pdfFormFields.SetField("AccFirstName", AccFirstName);
        pdfFormFields.SetField("AccFirstName", LEname);
        pdfFormFields.SetField("AccLastName", AccLastName);
        pdfFormFields.SetField("SLOA_LE_Name", SLOA_LE_Name);

        //pdfFormFields.SetField("AccountNumber", FidelityAcc);
        //pdfFormFields.SetField("Signer1", strSigner1);
        //pdfFormFields.SetField("Signer2", strSigner2);
        //pdfFormFields.SetField("Signer3", strSigner3);

        // flatten the form to remove editting options, set it to false
        // to leave the form open to subsequent manual edits
        pdfStamper.FormFlattening = true;
        pdfFormFields.GenerateAppearances = true;
        // close the pdf
        pdfStamper.Close();
        return newFile;
    }

    public string Get_SLOA2()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.SLOA);


        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();
        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        //pdfTemplate -- With TextFields to Write values
        string pdfTemplate = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\Fidelity-SLOA2.pdf";

        string newFile = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + "-SLOA2.pdf";

        PdfReader pdfReader = new PdfReader(pdfTemplate);
        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(
                    newFile, FileMode.Create));

        AcroFields pdfFormFields = pdfStamper.AcroFields;

        // set form pdfFormFields
        //string LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));
        //string FidelityAcc = Convert.ToString(table.Rows[0]["ssi_accountnumber"]);
        string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);

        // The first worksheet
        //pdfFormFields.SetField("LEName", LEname);
        //pdfFormFields.SetField("AccountNumber", FidelityAcc);
        pdfFormFields.SetField("Signer1", strSigner1);
        pdfFormFields.SetField("Signer2", strSigner2);
        pdfFormFields.SetField("Signer3", strSigner3);

        // flatten the form to remove editting options, set it to false
        // to leave the form open to subsequent manual edits
        pdfStamper.FormFlattening = true;
        pdfFormFields.GenerateAppearances = true;
        // close the pdf
        pdfStamper.Close();
        return newFile;
    }


    public string Get_SLOA3(DataTable loDataSet)
    {
        DataTable newdataset;
        DB clsDB = new DB();
        newdataset = null;
        //String lsSQL = getFinalSp(ReportType.SLOA);


        newdataset = loDataSet.Copy();
        newdataset.AcceptChanges();
        int DSCount = newdataset.Rows.Count;
        DataTable table = newdataset.Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "SLOA.pdf";




        // iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        //AddHeader(document);
        //AddFooter();

        document.Open();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        string FidelityAcc = string.Empty;
        string LENameNEW = "";
        string LENamOLD = "";

        //Creating Table for Heading with two column 
        PdfPTable LoHeader = new PdfPTable(1);

        int[] width12 = { 100 };
        LoHeader.SetWidths(width12);

        //Defining Width 100%
        LoHeader.WidthPercentage = 100f;

        FidelityAcc = Convert.ToString(table.Rows[0]["ssi_accountnumber"]);

        PdfPCell LEName = new PdfPCell();

        Paragraph HeadingLEtName = new Paragraph(LEname, setFontsAllTimesNewRoman(14, 1, 0));
        Paragraph HeadingAccount = new Paragraph("Fidelity Account #" + FidelityAcc, setFontsAllTimesNewRoman(14, 1, 0));

        LEName.AddElement(HeadingLEtName);
        LEName.AddElement(HeadingAccount);
        LEName.Border = 0;


        string strSigner1 = Convert.ToString(table.Rows[0]["ssi_signor1idname"]);
        string strSigner2 = Convert.ToString(table.Rows[0]["ssi_signor2idname"]);
        string strSigner3 = Convert.ToString(table.Rows[0]["ssi_signor3idname"]);


        PdfPTable LoFooter = new PdfPTable(3);
        int[] width1 = { 35, 35, 35 };
        LoFooter.SetWidths(width1);

        LoFooter.WidthPercentage = 100f;

        PdfPCell CellFooter11 = new PdfPCell();
        PdfPCell CellFooter12 = new PdfPCell();
        PdfPCell CellFooter13 = new PdfPCell();

        PdfPCell CellFooter21 = new PdfPCell();
        PdfPCell CellFooter22 = new PdfPCell();
        PdfPCell CellFooter23 = new PdfPCell();

        PdfPCell CellFooter31 = new PdfPCell();
        PdfPCell CellFooter32 = new PdfPCell();
        PdfPCell CellFooter33 = new PdfPCell();

        PdfPCell CellFooter1 = new PdfPCell();
        PdfPCell CellFooter2 = new PdfPCell();
        PdfPCell CellFooter3 = new PdfPCell();


        Paragraph PBlank = new Paragraph(" ", setFontsAllTimesNewRoman(10, 1, 0));
        Paragraph PSigner1 = new Paragraph(strSigner1, setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PSigner2 = new Paragraph(strSigner2, setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PSigner3 = new Paragraph(strSigner3, setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PAccountHolderSignature = new Paragraph("Account Holder Signature", setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PPrintedName = new Paragraph("Printed Name", setFontsAllTimesNewRoman(11, 0, 0));
        Paragraph PDate = new Paragraph("Date", setFontsAllTimesNewRoman(11, 0, 0));

        PAccountHolderSignature.SetAlignment("center");
        PPrintedName.SetAlignment("center");
        PDate.SetAlignment("center");

        PAccountHolderSignature.Leading = 8f;
        PPrintedName.Leading = 8f;
        PDate.Leading = 8f;


        PSigner1.SetAlignment("center");
        PSigner2.SetAlignment("center");
        PSigner3.SetAlignment("center");

        CellFooter11.Border = 0;
        CellFooter12.Border = 0;
        CellFooter13.Border = 0;

        CellFooter21.Border = 0;
        CellFooter22.Border = 0;
        CellFooter23.Border = 0;

        CellFooter31.Border = 0;
        CellFooter32.Border = 0;
        CellFooter33.Border = 0;

        CellFooter1.Border = 0;
        CellFooter2.Border = 0;
        CellFooter3.Border = 0;

        CellFooter1.BorderWidthTop = 1f;
        CellFooter2.BorderWidthTop = 1f;
        CellFooter3.BorderWidthTop = 1f;


        CellFooter11.AddElement(PBlank);
        CellFooter12.AddElement(PSigner1);
        CellFooter13.AddElement(PBlank);

        CellFooter11.PaddingTop = 10f;
        CellFooter12.PaddingTop = 10f;
        CellFooter13.PaddingTop = 10f;

        LoFooter.AddCell(CellFooter11);
        LoFooter.AddCell(CellFooter12);
        LoFooter.AddCell(CellFooter13);

        CellFooter1.AddElement(PAccountHolderSignature);
        CellFooter2.AddElement(PPrintedName);
        CellFooter3.AddElement(PDate);

        LoFooter.AddCell(CellFooter1);
        LoFooter.AddCell(CellFooter2);
        LoFooter.AddCell(CellFooter3);

        CellFooter21.AddElement(PBlank);
        CellFooter22.AddElement(PSigner2);
        CellFooter23.AddElement(PBlank);

        CellFooter21.PaddingTop = 10f;
        CellFooter22.PaddingTop = 10f;
        CellFooter23.PaddingTop = 10f;

        LoFooter.AddCell(CellFooter21);
        LoFooter.AddCell(CellFooter22);
        LoFooter.AddCell(CellFooter23);

        LoFooter.AddCell(CellFooter1);
        LoFooter.AddCell(CellFooter2);
        LoFooter.AddCell(CellFooter3);

        CellFooter31.AddElement(PBlank);
        CellFooter32.AddElement(PSigner3);
        CellFooter33.AddElement(PBlank);

        CellFooter31.PaddingTop = 10f;
        CellFooter32.PaddingTop = 10f;
        CellFooter33.PaddingTop = 10f;

        LoFooter.AddCell(CellFooter31);
        LoFooter.AddCell(CellFooter32);
        LoFooter.AddCell(CellFooter33);

        LoFooter.AddCell(CellFooter1);
        LoFooter.AddCell(CellFooter2);
        LoFooter.AddCell(CellFooter3);


        LoFooter.TotalWidth = 510;

        LoHeader.AddCell(LEName);

        //   document.Add(LoFooter);

        //Page Size Capacity for Portfolio Perf report
        int PageSize1 = 6;
        //Calculate How much requires to display records
        int TotalPage = (table.Rows.Count / PageSize1);

        //Check If it requires more page
        if (table.Rows.Count % PageSize1 != 0)
        {
            TotalPage = TotalPage + 1;
        }

        //Calulates and display records with Header and footer
        for (int CurPage = 1; CurPage <= TotalPage; CurPage++)
        {
            //New Table 
            PdfPTable ptab1 = new PdfPTable(1);
            ptab1.WidthPercentage = 100f;
            //New Cell
            PdfPCell loCellSLOA = new PdfPCell();

            //New table - Temp 
            PdfPTable Ptab;

            //Add New Page 
            document.NewPage();

            //add Header starts from 2nd page
            document.Add(LoHeader);

            Ptab = BindDynamicSLOA_Data(CurPage, TotalPage, PageSize1, table);

            Ptab.WidthPercentage = 100f;

            loCellSLOA.AddElement(Ptab);

            //No border
            loCellSLOA.Border = 0;
            loCellSLOA.PaddingLeft = 0f;
            loCellSLOA.PaddingRight = 0f;

            ptab1.AddCell(loCellSLOA);

            document.Add(ptab1);
            LoFooter.WriteSelectedRows(0, -1, 35, 140, writer.DirectContent);

            // document.Add(LoFooter);
        }


        document.Close();
        FileInfo loFile = new FileInfo(ls);
        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

        return fsFinalLocation.Replace(".xls", ".pdf");

    }

    public PdfPTable BindDynamicSLOA_Data(int CurrentPage, int TotalPage, int PageSize, DataTable table)
    {
        PdfPTable LoData = new PdfPTable(2);


        int[] width2 = { 50, 50 };
        LoData.SetWidths(width2);


        LoData.WidthPercentage = 100f;

        int DSCount = table.Rows.Count;

        //Calculate End Value of Grid.
        //i.e End Value  = Current Page Number * Page Size Capacity
        int finalVal = (CurrentPage * PageSize);

        //Calculate Starting Value of Grid.
        //i.e Starting Value  = End Value - Page Size Capacity
        int intialVal = finalVal - PageSize;

        //As Values in Grid Starting from Zero, End Value will Decreased by 1
        finalVal = finalVal - 1;

        if (CurrentPage == TotalPage)
            finalVal = table.Rows.Count - 1;


        for (int x = intialVal; x <= finalVal; x++)
        {
            PdfPCell CellData = new PdfPCell();
            string strPartnership = Convert.ToString(table.Rows[x]["ssi_fundname"]);
            string strABA = Convert.ToString(table.Rows[x]["Ssi_ABARoutingin"]);
            // string strCustName = Convert.ToString(table.Rows[x]["CustodianName"]);
            string strCustName = "";
            string strPartAcc = Convert.ToString(table.Rows[x]["ssi_accountnumber"]);
            string strOtherWireInst = Convert.ToString(table.Rows[x]["OtherWireInstr"]);
            string DDANumber = Convert.ToString(table.Rows[x]["Ssi_DDAoutgoing"]);
            string LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));
            string strBasicWireInfo = Convert.ToString(table.Rows[x]["BasicWireInfo"]);
            string strbeneficiaryAccount = Convert.ToString(table.Rows[x]["ssi_beneficiaryAccount"]);

            Paragraph PPartnership = new Paragraph(strPartnership, setFontsAllTimesNewRoman(10, 1, 0));


            //ABA
            Paragraph PABA = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PABAText = new Chunk("ABA #: ", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PABAVal = new Chunk(strABA, setFontsAllTimesNewRoman(10, 1, 0));

            //Bank
            Paragraph PBank = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PBankText = new Chunk("Bank: ", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PBankCustName = new Chunk(strCustName, setFontsAllTimesNewRoman(10, 0, 0));

            //Other info
            //Paragraph POtherInfo = new Paragraph("Other Info*", setFontsAllTimesNewRoman(10, 0, 0));
            Paragraph POtherInfo = new Paragraph(strBasicWireInfo, setFontsAllTimesNewRoman(10, 0, 0));

            //Beneficiary            
            Paragraph PBeneficiaryText = new Paragraph("Beneficiary: ", setFontsAllTimesNewRoman(10, 1, 0));
            Paragraph PBeneficiaryVal = new Paragraph(strPartnership, setFontsAllTimesNewRoman(10, 0, 0));

            //Bene Acct
            Paragraph PBeneAcct = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PBeneAcctText = new Chunk("Bene Acct #: ", setFontsAllTimesNewRoman(10, 1, 0));
            //Chunk PBeneAcctVal = new Chunk("Partnership Acct #", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PBeneAcctVal = new Chunk(strbeneficiaryAccount, setFontsAllTimesNewRoman(10, 0, 0));
            //Chunk PBeneAcctVal = new Chunk(strPartAcc, setFontsAllTimesNewRoman(10, 0, 0));

            //Details
            Paragraph PDetails = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PDetailsText = new Chunk("Details: REF: ", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PDetailsTVal = new Chunk(LEname, setFontsAllTimesNewRoman(10, 0, 0));

            Paragraph POtherWireInstVal = new Paragraph(strOtherWireInst, setFontsAllTimesNewRoman(10, 0, 0));

            //DDA
            Paragraph PDDA = new Paragraph("", setFontsAllTimesNewRoman(10, 0, 0));
            Chunk PDDAText = new Chunk("DDA #:", setFontsAllTimesNewRoman(10, 1, 0));
            Chunk PDDAVal = new Chunk(DDANumber, setFontsAllTimesNewRoman(10, 0, 0));

            PABA.Add(PABAText);
            PABA.Add(PABAVal);

            PBank.Add(PBankText);
            PBank.Add(PBankCustName);

            PBeneAcct.Add(PBeneAcctText);
            PBeneAcct.Add(PBeneAcctVal);

            PDetails.Add(PDetailsText);
            PDetails.Add(PDetailsTVal);

            PDDA.Add(PDDAText);
            PDDA.Add(PDDAVal);



            CellData.AddElement(PPartnership);
            CellData.AddElement(PABA);
            CellData.AddElement(PBank);
            CellData.AddElement(POtherInfo);

            CellData.AddElement(PBeneficiaryText);
            CellData.AddElement(PBeneficiaryVal);
            CellData.AddElement(PBeneAcct);
            CellData.AddElement(PDetails);

            CellData.AddElement(POtherWireInstVal);
            if (!string.IsNullOrEmpty(DDANumber))
                CellData.AddElement(PDDA);

            CellData.Border = 0;
            CellData.PaddingTop = 12f;
            LoData.AddCell(CellData);
        }

        if (DSCount % 2 != 0)
        {
            PdfPCell Test1 = new PdfPCell();
            Paragraph Test2 = new Paragraph(" ", setFontsAllTimesNewRoman(10, 0, 0));
            Test1.Border = 0;
            Test1.AddElement(Test2);
            Test1.PaddingTop = 12f;
            LoData.AddCell(Test1);
        }
        LoData.SpacingBefore = 10f;
        return LoData;
    }


    public string createPdftemp(string str)
    {
        StringBuilder sb = new StringBuilder(str);

        ArrayList objects = null;
        using (MemoryStream ms = new MemoryStream())
        {

            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

            Document document = new Document(PageSize.A4, 25, 25, 30, 30);
            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".pdf";

            iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));



            document.Open();

            // HTMLWorker hw = new HTMLWorker(document);


            iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

            iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();


            hw.Style = styles;

            //MemoryStream output = new MemoryStream();
            //StreamWriter html = new StreamWriter(output, Encoding.UTF8);



            //html.Write(string.Concat(DOCUMENT_HTML_START, str, DOCUMENT_HTML_END));
            //html.Close();
            //html.Dispose();

            //MemoryStream generate = new MemoryStream(output.ToArray());
            //StreamReader stringReader = new StreamReader(generate);

            //objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(new StringReader(str), null);
            //hw.Parse(new StringReader(str.ToString()));

            //document.Add(new Paragraph(str));

            string STYLE_DEFAULT_TYPE = "style";
            string DOCUMENT_HTML_START = "<html><head></head><body>";
            string DOCUMENT_HTML_END = "</body></html>";
            string REGEX_GROUP_SELECTOR = "selector";
            string REGEX_GROUP_STYLE = "style";

            //amazing regular expression magic
            string REGEX_GET_STYLES = @"(?<selector>[^\{\s]+\w+(\s\[^\{\s]+)?)\s?\{(?<style>[^\}]*)\}";

            foreach (Match match in Regex.Matches(str, REGEX_GET_STYLES))
            {
                string selector = match.Groups[REGEX_GROUP_SELECTOR].Value;
                string style = match.Groups[REGEX_GROUP_STYLE].Value;
                this.AddStyle(selector, style);
            }

            string strhtml = "<h5 style='margin: 0in 0in 0pt'><em><u><font color='#e36c0a'>Distribution Letter <o:p></o:p></font></u></em></h5> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>This template would require the following user input:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>As of Date<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Letter Date<o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Specific Text<o:p></o:p></font></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>The dynamic fields encoded into this template are the following fields from the mail record:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Salutation (ssi_salutation_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Full Name (ssi_fullname_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 1 (ssi_addressline1_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 2 (ssi_addressline2_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 3 (ssi_addressline3_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>City (ssi_city_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>State (ssi_stateprovince_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>ZIP Code (ssi_zipcode_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Country/Region (ssi_countryregion_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Dear (ssi_dear_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Legal Entity Name (ssi_legalentitynameid)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Name (ssi_fundname) <o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Nickname <i style='mso-bidi-font-style: normal'>(Need new field on the Fund to store this and ability to make the join back from the mail record to the fund dynamic to be able to retrieve the info)</i><o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Percent Called (ssi_percentcalled_ccsf)<o:p></o:p></font></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><o:p><font size='2'>&nbsp;</font></o:p></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>This template has the following permutations dependent on the mail record data:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l1 level1 lfo3'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Multiple Fund Holdings vs. Single Fund Holding<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Dependent on the number of different funds for a Legal Entity and Recipient the beginning of letter would change<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Multiple fund holdings would include a grid and plural text for the beginning ofthe letter <o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>A single fund holding would just have singular references for the beginning of the letter<o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 1in; mso-add-space: auto'><o:p><font size='2'>&nbsp;</font></o:p></p> <p class='MsoNormal' style='margin: 0in 0in 0pt 0.5in'><span style='color: #31849b; mso-themecolor: accent5; mso-themeshade: 191'><font size='2'>BEGINNING PARAGRAPH &ndash; MULTIPLE FUNDS</font></span><o:p></o:p></p> <p>&nbsp;</p>";

            //string str = System.Text.RegularExpressions.Regex.Matches(FundSpecificDesc, REGEX_GET_STYLES);
            MemoryStream output = new MemoryStream();
            StreamWriter html = new StreamWriter(output, Encoding.UTF8);


            html.Write(string.Concat(DOCUMENT_HTML_START, str, DOCUMENT_HTML_END));
            html.Close();
            html.Dispose();

            MemoryStream generate = new MemoryStream(output.ToArray());
            StreamReader stringReader = new StreamReader(generate);
            foreach (object item in iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, styles))
            {
                document.Add((IElement)item);
            }

            //cleanup these streams
            html.Dispose();
            stringReader.Dispose();
            output.Dispose();
            generate.Dispose();





            //  writer.Close();
            html.Dispose();
            stringReader.Dispose();
            output.Dispose();
            generate.Dispose();







            document.Close();
            return fsFinalLocation;


        }
    }


    #endregion


    #region Distribution Statement

    public string GetDistributionStatement()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.DistributionStatement);
        //String lsSQL = "SP_S_DistributionStatements @MailID = 2, @legalentitynameid = '6055D1AA-6E15-DE11-8391-001D09665E8F',@ContactFullnameID = '4F073302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "DistributionStatement.pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        //AddFooter();

        document.Open();

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //png.SetAbsolutePosition(45, 800);//540
        ////png.ScaleToFit(288f, 42f);
        //png.ScalePercent(8);
        //document.Add(png);


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        try
        {
            if (DSCount > 0)
            {
                for (int i = 0; i < DSCount; i++)
                {
                    if (i != 0)
                    {

                        document.NewPage();
                    }
                    string fundname = Convert.ToString(table.Rows[i]["ssi_fundname"]);
                    Chunk lochunk1 = new Chunk("\n\n" + fundname.ToUpper(), setFontsAllFrutiger(12, 1, 0));

                    string AsOfDate = "Statement as of " + Convert.ToString(table.Rows[i]["Ssi_WireAsOfDate"]);
                    Chunk lochunk2 = new Chunk("\n" + AsOfDate, setFontsAllFrutiger(10, 1, 0));

                    string Legalentityname = Convert.ToString(table.Rows[i]["ssi_legalentitynameidname"]);
                    Chunk lochunk3 = new Chunk("\n\n" + Legalentityname, setFontsAllFrutiger(11, 1, 0));

                    Paragraph p1 = new Paragraph();
                    p1.Add(lochunk1);
                    p1.Add(lochunk2);
                    p1.Add(lochunk3);
                    p1.Alignment = 1;
                    document.Add(p1);

                    iTextSharp.text.Table loTable = new iTextSharp.text.Table(4, 11);   // 2 rows, 2 columns           
                    iTextSharp.text.Cell loCell = new Cell();
                    lsTotalNumberofColumns = "4";
                    setTableProperty(loTable, ReportType.DistributionStatement);
                    iTextSharp.text.Chunk lochunk = new Chunk();

                    int rowsize = 11;
                    int colsize = 4;
                    for (int j = 0; j < rowsize; j++)
                    {
                        for (int k = 0; k < colsize; k++)
                        {
                            if (j == 0 && k == 0)
                            {
                                lochunk = new Chunk("% of      \nCommitted\n      Capital   ", setFontsAll(9, 1, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);

                                loCell.Colspan = 4;
                                k = k + 3;
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);
                            }
                            if (j == 1 && k == 0)
                            {
                                lochunk = new Chunk("Status of  Commitment", setFontsAll(9, 1, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);

                                loCell.Colspan = 4;
                                k = k + 3;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                                loTable.AddCell(loCell);
                            }
                            if (j == 2 && k == 0)
                            {
                                lochunk = new Chunk("Capital Commitment", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                //loCell.Colspan = 2;
                                loCell.Border = 0;

                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("$", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["totalcommitment_db"])).Replace("$", ""), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("100.00%", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 3 && k == 0)
                            {
                                lochunk = new Chunk("Called to Date", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.BorderWidthBottom = 1F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 1F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["calledtodate_db"])).Replace("$", ""), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 1F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["percentcalled_db"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 1F;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 4 && k == 0)
                            {
                                lochunk = new Chunk("Remaining Commitment\n\n\n", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                //loCell.Colspan = 2;
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                //loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("$\n\n\n", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                //loCell.BorderWidthBottom = 2F;
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["remainingcommitment_db"])).Replace("$", "") + "\n\n\n", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                //loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["remainingcommitment_per"])) + "\n\n\n", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                //loCell.BorderWidthBottom = 2F;
                                loTable.AddCell(loCell);
                            }
                            if (j == 5 && k == 0)
                            {
                                lochunk = new Chunk("Distributions to Date", setFontsAll(9, 1, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Add(lochunk);

                                loCell.Colspan = 4;
                                k = k + 3;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                                loTable.AddCell(loCell);
                            }
                            if (j == 6 && k == 0)
                            {
                                lochunk = new Chunk("Prior Distributions", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                //loCell.Colspan = 2;
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("$", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["priordistributions_db"])).Replace("$", ""), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["priordistp_db"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 7 && k == 0)
                            {
                                lochunk = new Chunk("Distribution- " + Convert.ToString(table.Rows[i]["Ssi_WireAsOfDate"]) + "", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                // loCell.Colspan = 2;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["capitaldistribution_db"])).Replace("$", ""), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["curdistp_db"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.Border = 0;
                                loCell.Leading = 11F;
                                loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                            if (j == 8 && k == 0)
                            {
                                if (Convert.ToString(table.Rows[i]["feeadj_db"]) != "0.0000")
                                {
                                    lochunk = new Chunk("Less: Accrued Class B management fee", setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    //loCell.Colspan = 2;
                                    loCell.BorderWidthBottom = 1F;
                                    loCell.Leading = 11F;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    loTable.AddCell(loCell);

                                    lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.BorderWidthBottom = 1F;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                    loCell.Leading = 11F;
                                    loTable.AddCell(loCell);

                                    lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["feeadj_db"])).Replace("$", ""), setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.BorderWidthBottom = 1F;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                    loCell.Leading = 11F;
                                    loTable.AddCell(loCell);

                                    lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.BorderWidthBottom = 1F;
                                    loCell.Leading = 11F;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                    loTable.AddCell(loCell);
                                }
                                else
                                {
                                    lochunk = new Chunk("", setFontsAll(9, 1, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.Colspan = 4;
                                    k = k + 3;
                                    loCell.BorderWidthTop = 1F;
                                    loCell.Leading = 11F;
                                    loTable.AddCell(loCell);
                                }
                            }
                            if (j == 9 && k == 0)
                            {
                                if (Convert.ToString(table.Rows[i]["feeadj_db"]) != "0.0000")
                                {
                                    lochunk = new Chunk("Total Adjusted Distribution", setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.Border = 0;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    loCell.Leading = 11F;
                                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                    loTable.AddCell(loCell);

                                    lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.Border = 0;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                    loCell.Leading = 11F;
                                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                    loTable.AddCell(loCell);

                                    lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["actualcashdistributions_db"])).Replace("$", ""), setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.Border = 0;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                    loCell.Leading = 11F;
                                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                    loTable.AddCell(loCell);

                                    lochunk = new Chunk("", setFontsAll(9, 0, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.Border = 0;
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                    loCell.Leading = 11F;
                                    loCell.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#B7DDE8"));
                                    loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                    loTable.AddCell(loCell);
                                }
                                else
                                {
                                    lochunk = new Chunk("", setFontsAll(9, 1, 0));
                                    loCell = new iTextSharp.text.Cell();
                                    loCell.Add(lochunk);
                                    loCell.Colspan = 4;
                                    k = k + 3;
                                    loCell.Border = 0;
                                    loCell.Leading = 11F;
                                    loTable.AddCell(loCell);
                                }
                            }
                            if (j == 10 && k == 0)
                            {
                                lochunk = new Chunk("Distributions to Date", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk("$", setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(RoundUp(Convert.ToString(table.Rows[i]["disttodate_db"])).Replace("$", ""), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 2F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 11F;
                                loTable.AddCell(loCell);

                                lochunk = new Chunk(Percentage(Convert.ToString(table.Rows[i]["distdatep_db"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();
                                loCell.Add(lochunk);
                                loCell.BorderWidthBottom = 2F;
                                loCell.Leading = 11F;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loTable.AddCell(loCell);
                            }
                        }
                    }
                    document.Add(loTable);
                }
            }
        }
        catch
        {

        }
        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    #endregion

    #region Distribution Wire Instructions

    public string DistributionWireInstruction()
    {
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.DistributionWireInstruction);
        //String lsSQL = "SP_S_DistributionWireInstructions @MailID = 2, @legalentitynameid = '6055D1AA-6E15-DE11-8391-001D09665E8F',@ContactFullnameID = '4F073302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();
        rand.Next();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "DistributionWireInstruction.pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        //AddFooter();

        document.Open();

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //png.SetAbsolutePosition(45, 800);//540
        ////png.ScaleToFit(288f, 42f);
        //png.ScalePercent(8);
        //document.Add(png);


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        try
        {

            if (table.Rows.Count > 0)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (i != 0)
                    {
                        document.NewPage();

                        ////iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
                        //png.SetAbsolutePosition(45, 557);//540
                        ////png.ScaleToFit(288f, 42f);
                        //png.ScalePercent(10);
                        //document.Add(png);
                    }

                    string PartnerShip = "Partnership: " + Convert.ToString(table.Rows[i]["ssi_fundname"]) + "";
                    Chunk lochunk1 = new Chunk("\n\n" + PartnerShip, setFontsAll(9, 0, 0));

                    string Investor = "Investor: " + Convert.ToString(table.Rows[i]["ssi_legalentitynameidname"]) + "";
                    Chunk lochunk2 = new Chunk("\n\n" + Investor, setFontsAll(9, 0, 0));

                    string Amount = RoundUp(Convert.ToString(table.Rows[i]["ssi_capitaldistribution_db"])) + "";
                    Chunk lochunk33 = new Chunk("\n\n Amount: ", setFontsAll(9, 0, 0));
                    Chunk lochunk3 = new Chunk(Amount, setFontsAll(9, 1, 0));

                    string WireInstr = "Wire Instructions:";
                    Chunk lochunk4 = new Chunk("\n\n" + WireInstr, setFontsAll(9, 0, 0));

                    Paragraph p1 = new Paragraph();
                    p1.Add(lochunk1);
                    p1.Add(lochunk2);
                    p1.Add(lochunk33);
                    p1.Add(lochunk3);
                    p1.Add(lochunk4);
                    p1.Leading = 11.0f;
                    document.Add(p1);

                    string basicwireinfo_household1 = Convert.ToString(table.Rows[i]["ssi_ssi_basicwireinfo_household1"]);//.Replace("\r\n", "")
                    Chunk lochunk5 = new Chunk("\n" + basicwireinfo_household1, setFontsAll(9, 0, 0));

                    string abarouting_household = "ABA: " + Convert.ToString(table.Rows[i]["ssi_abarouting_household"]) + "";
                    Chunk lochunk6 = new Chunk("\n" + abarouting_household, setFontsAll(9, 0, 0));

                    string ffcname_household = Convert.ToString(table.Rows[i]["ssi_ffcname_household"]);
                    Chunk lochunk7 = new Chunk("\n" + ffcname_household, setFontsAll(9, 0, 0));

                    string ffcacct_household = "FFC Account #: " + Convert.ToString(table.Rows[i]["ssi_ffcacct_household"]) + "";
                    Chunk lochunk8 = new Chunk("\n" + ffcacct_household, setFontsAll(9, 0, 0));

                    string otherwireinstr_household = Convert.ToString(table.Rows[i]["ssi_otherwireinstr_household"]);
                    Chunk lochunk9 = new Chunk("\n" + otherwireinstr_household, setFontsAll(9, 0, 0));

                    string accountname1 = "For Benefit of: " + Convert.ToString(table.Rows[i]["ssi_accountname1"]) + "";
                    Chunk lochunk10 = new Chunk("\n" + accountname1, setFontsAll(9, 0, 0));

                    string accountnumber = "Account #: " + Convert.ToString(table.Rows[i]["ssi_accountnumber"]) + "";
                    Chunk lochunk11 = new Chunk("\n" + accountnumber, setFontsAll(9, 0, 0));

                    Paragraph p2 = new Paragraph();
                    //Paragraph p3 = new Paragraph();
                    //p3.Add(lochunk5);
                    //p3.IndentationLeft = 75.0f;
                    //document.Add(p3);

                    p2.Add(lochunk5);
                    p2.Add(lochunk6);
                    p2.Add(lochunk7);
                    if (Convert.ToString(table.Rows[i]["ssi_ffcacct_household"]) != "")
                        p2.Add(lochunk8);
                    //if (Convert.ToString(table.Rows[i]["ssi_otherwireinstr_household"]) != "")
                    p2.Add(lochunk9);
                    p2.Add(lochunk10);
                    p2.Add(lochunk11);
                    p2.IndentationLeft = 75.0f;
                    p2.Leading = 11.0f;
                    document.Add(p2);
                }
            }
        }
        catch { }


        if (table.Rows.Count > 0)
        {
            document.Close();
            try
            {
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            {

            }
        }
        else
        {
            fsFinalLocation = "";
        }


        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    #endregion

    #endregion

    #region Custom Report Templates

    #region Capital Call Letter Custom
    public string GetCapitalCallLetterStatementCustom()
    {
        string[] CheckString;
        int liPageSize = 39;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.CapitalCallStatementCustom); //"SP_S_CapitalCallLetterCustom 13, '18511645-77CA-E111-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE11-8391-001D09665E8F', 'F7063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);

        var dv = newdataset.Tables[0].DefaultView;
        dv.RowFilter = "RetirementFlg = 0";
        var newDS = new DataSet();
        var newDT = dv.ToTable();
        newDS.Tables.Add(newDT);


        int DSCount = newDS.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();
        rand.Next();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 50, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        PdfWriter pdfwriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        AddFooter(document);

        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\GreshamAdvisors_logo.tif");
        png.SetAbsolutePosition(48, 800);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(22);
        document.Add(png);

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        try
        {
            string effCapitalCallDate = string.Empty;
            if (DSCount > 0)
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "")
                {
                    DateTime dt = Convert.ToDateTime(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]);
                    effCapitalCallDate = dt.ToString("MMMM") + " 1, " + dt.ToString("yyy");
                }
            }

            if (DSCount > 1)
            {
                #region Multiple Standard and Non Standard

                #region Details

                Chunk lochunkAsOfDate = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_LetterDate"]) != "")
                {
                    lochunkAsOfDate = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_LetterDate"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAsOfDate != null)
                {
                    //Chunk lochunkAsOfDate = new Chunk(LetterDate, setFontsAll(11, 0, 0));
                    Paragraph pAsOfDate = new Paragraph();
                    pAsOfDate.Add(lochunkAsOfDate);
                    pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(pAsOfDate);
                }

                Chunk lochunkFullName = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]) != "")
                {
                    lochunkFullName = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkFullName != null)
                {
                    // Chunk lochunkFullName = new Chunk(FullName, setFontsAll(11, 0, 0));
                    Paragraph pFullName = new Paragraph();
                    pFullName.Add(lochunkFullName);
                    pFullName.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pFullName.Leading = 11f;
                    document.Add(pFullName);
                }

                Chunk lochunkAddressLine1 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]) != "")
                {
                    lochunkAddressLine1 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressLine1 != null)
                {
                    //   Chunk lochunkAddressLine1 = new Chunk(AddressLine1, setFontsAll(11, 0, 0));
                    Paragraph pAddressLine1 = new Paragraph();
                    pAddressLine1.Add(lochunkAddressLine1);
                    pAddressLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressLine1.Leading = 11f;
                    document.Add(pAddressLine1);
                }

                Chunk lochunkAddressLine2 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]) != "")
                {
                    lochunkAddressLine2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressLine2 != null)
                {
                    //Chunk lochunkAddressLine2 = new Chunk(AddressLine2, setFontsAll(11, 0, 0));
                    Paragraph pAddressLine2 = new Paragraph();
                    pAddressLine2.Add(lochunkAddressLine2);
                    pAddressLine2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressLine2.Leading = 11f;
                    document.Add(pAddressLine2);
                }

                Chunk lochunkAddressLine3 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]) != "")
                {
                    lochunkAddressLine3 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressLine3 != null)
                {
                    //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                    Paragraph pAddressLine3 = new Paragraph();
                    pAddressLine3.Add(lochunkAddressLine3);
                    pAddressLine3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressLine3.Leading = 11f;
                    document.Add(pAddressLine3);
                }

                Chunk lochunkAddressDetails = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]) != "")
                {
                    lochunkAddressDetails = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) + ", " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressDetails != null)
                {
                    //iTextSharp.text.Font contentFont = iTextSharp.text.FontFactory.GetFont("Verdana", 12, iTextSharp.text.Font.NORMAL);
                    //Chunk lochunkAddressDetails = new Chunk(AddressDetails, setFontsAll(11, 0, 0));
                    Paragraph pAddressDetails = new Paragraph();
                    pAddressDetails.Add(lochunkAddressDetails);
                    pAddressDetails.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressDetails.Leading = 11f;
                    document.Add(pAddressDetails);
                }

                Chunk lochunkCountry = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]) != "")
                {
                    lochunkCountry = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkCountry != null)
                {
                    //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                    Paragraph pCountry3 = new Paragraph();
                    pCountry3.Add(lochunkCountry);
                    pCountry3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pCountry3.Leading = 11f;
                    document.Add(pCountry3);
                }

                Chunk lochunkFullName1 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]) != "")
                {
                    lochunkFullName1 = new Chunk("RE: " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]), setFontsAll(9, 1, 0));
                }
                if (lochunkFullName1 != null)
                {
                    //Chunk lochunkFullName1 = new Chunk(FullNameBold, setFontsAll(11, 1, 0));
                    Paragraph pFullName1 = new Paragraph();
                    pFullName1.Add(lochunkFullName1);
                    pFullName1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pFullName1.SpacingBefore = 12f;
                    document.Add(pFullName1);
                }

                Chunk lochunkdear = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                {
                    lochunkdear = new Chunk("\nDear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + ":", setFontsAll(9, 0, 0));
                }
                if (lochunkdear != null)
                {
                    //Chunk lochunkdear = new Chunk(dear, setFontsAll(11, 0, 0));
                    Paragraph pdear = new Paragraph();
                    pdear.Add(lochunkdear);
                    pdear.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(pdear);
                }
                #endregion

                //string Instructions = "\nThis letter constitutes a call on following partnerships:";
                Chunk lochunk1 = new Chunk("This letter constitutes a call on following partnerships:", setFontsAll(9, 0, 0));

                Paragraph p1 = new Paragraph();
                p1.Add(lochunk1);
                p1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                p1.SpacingBefore = 5f;
                document.Add(p1);

                iTextSharp.text.Table loTable = new iTextSharp.text.Table(3, newdataset.Tables[0].Rows.Count);  // 2 rows, 3 columns          
                iTextSharp.text.Cell loCell = new Cell();
                lsTotalNumberofColumns = "3";
                setTableProperty(loTable, ReportType.CapitalCallStatementCustom);
                iTextSharp.text.Chunk lochunk = new Chunk();

                int rowsize = newdataset.Tables[0].Rows.Count;
                int colsize = 3;

                //int liTotalPage = (rowsize / liPageSize);
                //int liCurrentPage = 0;
                //if (rowsize % liPageSize != 0)
                //{
                //    liTotalPage = liTotalPage + 1;
                //}
                //else
                //{
                //    liPageSize = 2;
                //    liTotalPage = liTotalPage + 1;
                //}

                // Loop for Rows
                setHeaderCapitalCallStatementCustom(document);
                string strRetirementflg = "0";
                for (int j = 0; j < rowsize; j++)
                {
                    strRetirementflg = Convert.ToString(newdataset.Tables[0].Rows[j]["RetirementFlg"]);
                    if (strRetirementflg != "1")
                    {
                        // Loop for Columns
                        for (int k = 0; k < colsize; k++)
                        {
                            if (k == 0)
                            {
                                lochunk = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_fundname"]), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 9f;
                            }
                            if (k == 1)
                            {
                                lochunk = new Chunk(Percentage(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_percentcalled_ccsf"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell.Leading = 9f;
                            }
                            if (k == 2)
                            {
                                lochunk = new Chunk(RoundUp(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_currentcall_ccsf"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 9f;
                            }
                            loCell.Add(lochunk);
                            loTable.AddCell(loCell);
                        }


                        //document.Add(loTable);
                        //liCurrentPage = liCurrentPage + 1;

                        if (j == newdataset.Tables[0].Rows.Count - 1)
                        {
                            document.Add(loTable);
                            //liCurrentPage = liCurrentPage + 1;
                            //document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                        }
                    }
                }
                getTotalAmountCapitalCallStatementCustom(document, (Convert.ToString(newdataset.Tables[0].Rows[0]["TotalcapitalCall"])));

                Chunk Instructions1 = new Chunk("The enclosed statements provide you the call amounts, as well as detailed information on your commitments and funding percentages. These calls are due on ", setFontsAll(9, 0, 0));

                Chunk GGESInstructions = new Chunk("  Your investment in GP Diversified Growth Strategies has been deducted by the amount of these capital calls, effective " + effCapitalCallDate + ".  No action is required on your part.", setFontsAll(9, 1, 0));

                Chunk descretionarytext = new Chunk(" Your Fidelity account will be debited by the amount of these capital calls on that date. No action is required on your part.", setFontsAll(9, 1, 0));

                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "") //ssi_wireasofdate
                {
                    Chunk lochunk2;
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                        lochunk2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]), setFontsAll(9, 0, 0));
                    else
                        lochunk2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]), setFontsAll(9, 1, 0));

                    Chunk fullstop = new Chunk(".", setFontsAll(9, 0, 0));
                    Paragraph p2 = new Paragraph();
                    p2.Add(Instructions1);
                    p2.Add(lochunk2);
                    p2.Add(fullstop);
                    if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) == true)
                        p2.Add(descretionarytext);

                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                        p2.Add(GGESInstructions);
                    p2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    p2.SpacingAfter = 11f;
                    p2.SpacingBefore = 3f;
                    p2.Leading = 11f;
                    document.Add(p2);
                }


                iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);




                hw.Style = styles;

                //styles.LoadTagStyle("ol", "leading", "16,0");
                //styles.LoadTagStyle("li", "face", "garamond");
                //styles.LoadTagStyle("span", "size", "11pt");
                //styles.LoadTagStyle("body", "font-family", "Verdana");
                //styles.LoadTagStyle("body", "font-size", "11pt");
                //styles.LoadTagStyle(HtmlTags.STRONG, HtmlTags.FONT, "Verdana");
                //styles.LoadTagStyle(HtmlTags.STRONG, HtmlTags.SIZE, "9");

                ArrayList objects = null;

                string FundSpecificDesc = "";
                //List list = new List(List.UNORDERED, 10f);
                //list.SetListSymbol("\u2022");
                //list.IndentationLeft = 30f;
                for (int m = 0; m < rowsize; m++)
                {

                    if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                    {

                        FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");

                        //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                        // FundSpecificDesc = FundSpecificDesc.Replace("<p style=\"margin-left:", "<li style='margin: 0in 0in 0pt 0.5in; ' ");
                        //                        F/undSpecificDesc = FundSpecificDesc.Replace("pt\">  </p>", "</li>");
                        //FundSpecificDesc = FundSpecificDesc.Replace("<p", "<li");
                        //FundSpecificDesc = FundSpecificDesc.Replace("</p>", "</li>");
                        Chunk NextLine1 = new Chunk("\n");
                        Paragraph pNextLine1 = new Paragraph();
                        pNextLine1.Add(NextLine1);
                        pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        //document.Add(pNextLine1);


                        //FontFactory.RegisterDirectories();




                        //string str = System.Text.RegularExpressions.Regex.Matches(FundSpecificDesc, REGEX_GET_STYLES);
                        MemoryStream output = new MemoryStream();
                        StreamWriter html = new StreamWriter(output, Encoding.UTF8);



                        html.Write(string.Concat(DOCUMENT_HTML_START, FundSpecificDesc, DOCUMENT_HTML_END));
                        html.Close();
                        html.Dispose();

                        MemoryStream generate = new MemoryStream(output.ToArray());
                        StreamReader stringReader = new StreamReader(generate);

                        objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(new StringReader(FundSpecificDesc), null);

                        int rowcount = objects.Count;
                        //int colsize = 2;

                        int liTotalPage = (rowcount / liPageSize);
                        int liCurrentPage = 0;
                        if (rowcount % liPageSize != 0)
                        {
                            liTotalPage = liTotalPage + 1;
                        }
                        else
                        {
                            liPageSize = 39;
                            liTotalPage = liTotalPage + 1;
                        }
                        //add the collection to the document
                        for (int k = 0; k < objects.Count; k++)
                        {
                            if (k % liPageSize == 0)
                            {
                                document.Add((IElement)objects[k]);
                                if (k != 0)
                                {
                                    liCurrentPage = liCurrentPage + 1;
                                    document.NewPage();
                                }
                            }
                            else
                            {
                                document.Add((IElement)objects[k]);
                            }
                        }

                        //foreach (object item in iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, styles))
                        //{
                        //    document.Add((IElement)item);
                        //}

                        //cleanup these streams
                        html.Dispose();
                        stringReader.Dispose();
                        output.Dispose();
                        generate.Dispose();

                        //using (StreamReader stringReader = new StreamReader(generate))
                        //{
                        //    //List<IElement> parsedList = new List<IElement>();

                        //    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                        //    //hw.
                        //    //hw.Parse(stringAddReader);

                        //    //document.Open();\xB7
                        //    foreach (object item in objects)
                        //    {
                        //        //document.Add((ITextElementArray)item);
                        //        if (item is List)
                        //        {
                        //            List list = item as List;
                        //            list.Autoindent = false;
                        //            list.IndentationLeft = 25f;
                        //            list.SetListSymbol("\u2022                      ");
                        //            list.SymbolIndent = 25f;
                        //            document.Add((IElement)item);
                        //        }


                        //        if (item is Paragraph)
                        //        {
                        //            Paragraph para = item as Paragraph; //setFontsverdana

                        //            if (para.ToArray()[0].ToString() == "ע || para.ToArray()[0].ToString() == "Ǣ || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "آ || para.ToArray()[0].ToString() == "𢠼| para.ToArray()[0].ToString() == "v")
                        //            {
                        //                //((iTextSharp.text.Chunk)(para.ToArray()[0])).SetGenericTag("\u20AC");
                        //                //((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                        //                //para.IndentationLeft = 30f;
                        //                document.Add(para);
                        //            }
                        //            else
                        //            {
                        //                //for (int j = 0; j < para.ToArray().Length; j++)
                        //                //{
                        //                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                        //                //}
                        //                document.Add(para);
                        //            }
                        //        }
                        //    }
                        //    //document.Close();
                        //}

                        //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                        if (m != rowsize - 1)
                        {
                            Chunk NextLine = new Chunk("\n");
                            Paragraph pNextLine = new Paragraph();
                            pNextLine.Add(NextLine);
                            pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(pNextLine);
                        }
                    }
                }
                //document.Add(list);

                double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                if (remainingPageSpace < 217.00)
                    document.NewPage();

                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Verify".ToUpper())
                {
                    Chunk EndingParaMulStandard1 = new Chunk("To wire transfer necessary funds for capital calls, please sign the enclosed wire request form, and", setFontsAll(9, 0, 0));
                    Chunk ParaMulStandard1 = new Chunk(". We will ensure that your wire is processed by the due date.", setFontsAll(9, 0, 0));
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["TemplateMultiFundDate"]) != "")
                    {
                        //EndingParaMulStandard1 = EndingParaMulStandard1 + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]) + ParaMulStandard1;
                        Chunk lochunk3 = new Chunk(" fax it to us at (312)960-0204 by " + Convert.ToString(newdataset.Tables[0].Rows[0]["TemplateMultiFundDate"]), setFontsAll(9, 1, 0));

                        if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                        {
                            Paragraph EndMulStandard1 = new Paragraph();
                            // EndMulStandard1.Add(EndingParaMulStandard1);
                            //  EndMulStandard1.Add(lochunk3);
                            // EndMulStandard1.Add(ParaMulStandard1);

                            EndMulStandard1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            EndMulStandard1.SpacingBefore = 11f;
                            EndMulStandard1.Leading = 11f;
                            document.Add(EndMulStandard1);
                        }
                    }


                    //string EndingParaMulStandard2 = "\nIf you have questions, please call Ted Neild (312)960-0231 or Ben Beavers (312)960-0211";
                    Chunk lophonenumber1 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                    Chunk lodesc1 = new Chunk("or Ben Beavers", setFontsAll(9, 0, 0));
                    Chunk lophonenumber2 = new Chunk("(312)960-0211", setFontsAll(9, 1, 0));
                    Chunk lochunk4 = new Chunk("\n" + lophonenumber1, setFontsAll(9, 0, 0));
                    Paragraph EndMulStandard2 = new Paragraph();
                    EndMulStandard2.Add(lophonenumber1);
                    //EndMulStandard2.Add(lodesc1);
                    //EndMulStandard2.Add(lophonenumber2);
                    //EndMulStandard2.Add(lophonenumber1);
                    //EndMulStandard2.Add(lophonenumber1);
                    //EndMulStandard2.Add(lophonenumber2);
                    EndMulStandard2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndMulStandard2.SpacingBefore = 11f;
                    EndMulStandard2.Leading = 11f;
                    document.Add(EndMulStandard2);


                    // string EndingParaMulStandard3 = "Sincerely Yours,";
                    Chunk lochunk5 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph EndMulStandard3 = new Paragraph();
                    EndMulStandard3.Add(lochunk5);
                    EndMulStandard3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndMulStandard3.SpacingBefore = 8f;
                    document.Add(EndMulStandard3);

                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                    try
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                        else if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = HttpContext.Current.Server.MapPath("") + @"images\ImageNotAvailable.jpg";
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                    }
                    catch (Exception ex) { }

                    //string EndingParaMulStandard4 = "Ted Neild";
                    Chunk lochunk6 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                    Paragraph EndMulStandard4 = new Paragraph();
                    EndMulStandard4.Add(lochunk6);
                    EndMulStandard4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(EndMulStandard4);
                }
                else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Verify".ToUpper())
                {
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["TemplateMultiFundDate"]) != "")
                    {
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Pays from GGES".ToUpper())
                        {
                            //string EndingParaNonStandard1 = "You can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.";
                            //Chunk lochunkNonStd1 = new Chunk("\nYou can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            Chunk lochunkNonStd1;
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Check".ToUpper())
                            {
                                lochunkNonStd1 = new Chunk("You can fund your capital call by sending a check as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            }
                            else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Other Wire".ToUpper())
                            {
                                lochunkNonStd1 = new Chunk("You can fund your capital call by wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            }
                            else
                            {
                                lochunkNonStd1 = new Chunk("You can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            }

                            if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                            {
                                Paragraph EndNonStandard1 = new Paragraph();
                                EndNonStandard1.Add(lochunkNonStd1);
                                EndNonStandard1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                EndNonStandard1.Leading = 11f;
                                EndNonStandard1.SpacingBefore = 11f;
                                document.Add(EndNonStandard1);
                            }
                        }
                    }
                    //string EndingParaNonStandard2 = "payment instructions.";
                    //Chunk lochunkNonStd2 = new Chunk("\n" + EndingParaNonStandard2, setFontsAll(11, 0, 0));
                    //Paragraph EndNonStandard2 = new Paragraph();
                    //EndNonStandard2.Add(lochunkNonStd2);
                    //EndNonStandard2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //document.Add(EndNonStandard2);


                    //string EndingParaNonStandard3 = "\nIf you have questions, please call Ted Neild (312)960-0231 or Ben Beavers (312)960-0211";
                    Chunk lophone1 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " (312) 960-0200.", setFontsAll(9, 0, 0));
                    Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                    Chunk lophone2 = new Chunk("(312)960-0211.", setFontsAll(9, 1, 0));
                    Chunk lochunkNonStd3 = new Chunk("" + lophone1, setFontsAll(9, 0, 0));
                    Paragraph EndNonStandard3 = new Paragraph();
                    EndNonStandard3.Add(lochunkNonStd3);
                    // EndNonStandard3.Add(lophone1);
                    //EndNonStandard3.Add(lodesc1);
                    //EndNonStandard3.Add(lophone2);
                    EndNonStandard3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndNonStandard3.SpacingBefore = 11f;
                    EndNonStandard3.Leading = 11f;
                    document.Add(EndNonStandard3);



                    Chunk lochunkNonStd4 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph EndNonStandard4 = new Paragraph();
                    EndNonStandard4.Add(lochunkNonStd4);
                    EndNonStandard4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndNonStandard4.SpacingBefore = 8f;
                    document.Add(EndNonStandard4);


                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                    try
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                        else if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg";
                            //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                    }
                    catch (Exception ex) { }
                    Chunk lochunkNonStd5 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                    Paragraph EndNonStandard5 = new Paragraph();
                    EndNonStandard5.Add(lochunkNonStd5);
                    EndNonStandard5.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(EndNonStandard5);

                }

                #endregion
            }
            else if (DSCount > 0)
            {
                #region Single Standard and Non Standard
                string strretirementflg = "0";
                for (int i = 0; i < DSCount; i++)
                {
                    strretirementflg = Convert.ToString(newdataset.Tables[0].Rows[i]["RetirementFlg"]);
                    if (strretirementflg != "1")
                    {
                        #region Details
                        lsTotalNumberofColumns = "";
                        Chunk lochunkAsOfDate = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["Ssi_LetterDate"]) != "")
                        {
                            lochunkAsOfDate = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[i]["Ssi_LetterDate"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAsOfDate != null)
                        {
                            //Chunk lochunkAsOfDate = new Chunk(LetterDate, setFontsAll(11, 0, 0));
                            Paragraph pAsOfDate = new Paragraph();
                            pAsOfDate.Add(lochunkAsOfDate);
                            pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(pAsOfDate);
                        }


                        Chunk lochunkFullName = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_salutation_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_fullname_mail"]) != "")
                        {
                            lochunkFullName = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_fullname_mail"]), setFontsAll(9, 0, 0));
                        }

                        if (lochunkFullName != null)
                        {
                            //Chunk lochunkFullName = new Chunk(FullName, setFontsAll(11, 0, 0));
                            Paragraph pFullName = new Paragraph();
                            pFullName.Add(lochunkFullName);
                            pFullName.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pFullName.Leading = 11f;
                            document.Add(pFullName);
                        }

                        Chunk lochunkAddressLine1 = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline1_mail"]) != "")
                        {
                            lochunkAddressLine1 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline1_mail"]), setFontsAll(9, 0, 0));
                        }

                        if (lochunkAddressLine1 != null)
                        {
                            //Chunk lochunkAddressLine1 = new Chunk(AddressLine1, setFontsAll(11, 0, 0));
                            Paragraph pAddressLine1 = new Paragraph();
                            pAddressLine1.Add(lochunkAddressLine1);
                            pAddressLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressLine1.Leading = 11f;
                            document.Add(pAddressLine1);
                        }

                        Chunk lochunkAddressLine2 = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline2_mail"]) != "")
                        {
                            lochunkAddressLine2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline2_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAddressLine2 != null)
                        {
                            //Chunk lochunkAddressLine2 = new Chunk(AddressLine2, setFontsAll(11, 0, 0));
                            Paragraph pAddressLine2 = new Paragraph();
                            pAddressLine2.Add(lochunkAddressLine2);
                            pAddressLine2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressLine2.Leading = 11f;
                            document.Add(pAddressLine2);
                        }

                        Chunk lochunkAddressLine3 = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline3_mail"]) != "")
                        {
                            lochunkAddressLine3 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline3_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAddressLine3 != null)
                        {
                            //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                            Paragraph pAddressLine3 = new Paragraph();
                            pAddressLine3.Add(lochunkAddressLine3);
                            pAddressLine3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressLine3.Leading = 11f;
                            document.Add(pAddressLine3);
                        }

                        Chunk lochunkAddressDetails = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_city_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_stateprovince_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_zipcode_mail"]) != "")
                        {
                            lochunkAddressDetails = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_city_mail"]) + ", " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_stateprovince_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_zipcode_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAddressDetails != null)
                        {
                            //Chunk lochunkAddressDetails = new Chunk(AddressDetails, setFontsAll(11, 0, 0));
                            Paragraph pAddressDetails = new Paragraph();
                            pAddressDetails.Add(lochunkAddressDetails);
                            pAddressDetails.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressDetails.Leading = 11f;
                            document.Add(pAddressDetails);
                        }

                        Chunk lochunkCountry = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]) != "")
                        {
                            lochunkCountry = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkCountry != null)
                        {
                            //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                            Paragraph pCountry3 = new Paragraph();
                            pCountry3.Add(lochunkCountry);
                            pCountry3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pCountry3.Leading = 11f;
                            document.Add(pCountry3);
                        }

                        Chunk lochunkFullName1 = null;
                        //string FullNameBold = "\n\n RE: ";
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_legalentitynameidname"]) != "")
                        {
                            lochunkFullName1 = new Chunk("RE: " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_legalentitynameidname"]), setFontsAll(9, 1, 0));
                        }
                        if (lochunkFullName1 != null)
                        {
                            //Chunk lochunkFullName1 = new Chunk(FullNameBold, setFontsAll(11, 0, 0));
                            Paragraph pFullName1 = new Paragraph();
                            pFullName1.Add(lochunkFullName1);
                            pFullName1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pFullName1.SpacingBefore = 12f;
                            document.Add(pFullName1);
                        }

                        Chunk lochunkdear = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                        {
                            lochunkdear = new Chunk("\nDear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + ":", setFontsAll(9, 0, 0));
                        }
                        if (lochunkdear != null)
                        {
                            //Chunk lochunkdear = new Chunk(dear, setFontsAll(11, 0, 0));
                            Paragraph pdear = new Paragraph();
                            pdear.Add(lochunkdear);
                            pdear.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(pdear);
                        }


                        #endregion


                        iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, newdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                        iTextSharp.text.Cell loCell = new Cell();

                        iTextSharp.text.Chunk lochunk = new Chunk();

                        int rowsize = newdataset.Tables[0].Rows.Count;


                        //string InstructionsSTDandNonStd = "\nThis letter constitutes a ";
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) != "" && Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_percentcalled_ccsf"]) != "" && Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_currentcall_ccsf"]) != "")
                        {
                            //InstructionsSTDandNonStd = "\nThis letter constitutes a " + Percentage(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_percentcalled_ccsf"])) + " call on your commitment to " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) + ".";
                            Chunk lochunkSTDandNonStd = new Chunk("This letter constitutes a " + Percentage(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_percentcalled_ccsf"])) + " call on your commitment to " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                            Paragraph plochunkSTDandNonStd = new Paragraph();
                            plochunkSTDandNonStd.Add(lochunkSTDandNonStd);
                            plochunkSTDandNonStd.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSTDandNonStd.SpacingBefore = 12f;
                            plochunkSTDandNonStd.Leading = 11f;
                            document.Add(plochunkSTDandNonStd);
                        }


                        //string InstructionsSTDandNonStd1 = "\nThe enclosed statement provides you the call amount, as well as detailed information on";
                        Chunk lochunkSTDandNonStd1 = new Chunk("The enclosed statement provides you the call amount, as well as detailed information on commitment and", setFontsAll(9, 0, 0));

                        Chunk GGESInstructions = new Chunk("  Your investment in GP Diversified Growth Strategies has been deducted by the amount of this capital call, effective " + effCapitalCallDate + ".  No action is required on your part.", setFontsAll(9, 1, 0));

                        Chunk descretionarytextsingle = new Chunk(" Your Fidelity account will be debited by the amount of this capital call on that date. No action is required on your part.", setFontsAll(9, 1, 0));

                        Paragraph plochunkSTDandNonStd1 = new Paragraph();
                        plochunkSTDandNonStd1.Add(lochunkSTDandNonStd1);
                        plochunkSTDandNonStd1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        plochunkSTDandNonStd1.SpacingBefore = 12f;
                        plochunkSTDandNonStd1.Leading = 11f;
                        document.Add(plochunkSTDandNonStd1);


                        Chunk InstructionsSTDandNonStd2 = new Chunk("funding percentage. This call is due on ", setFontsAll(9, 0, 0));
                        Chunk InstructionsSTDandNonStd33 = new Chunk(".", setFontsAll(9, 0, 0));
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "")
                        {
                            //InstructionsSTDandNonStd2 = InstructionsSTDandNonStd2 + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]);
                            Chunk lochunkSTDandNonStd2;
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                                lochunkSTDandNonStd2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) + "", setFontsAll(9, 0, 0));
                            else
                                lochunkSTDandNonStd2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) + "", setFontsAll(9, 1, 0));

                            Chunk lochunkSTDandNonStd222 = new Chunk(".", setFontsAll(9, 0, 0));
                            Paragraph plochunkSTDandNonStd2 = new Paragraph();
                            plochunkSTDandNonStd2.Add(InstructionsSTDandNonStd2);
                            plochunkSTDandNonStd2.Add(lochunkSTDandNonStd2);
                            plochunkSTDandNonStd2.Add(lochunkSTDandNonStd222);
                            if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) == true)
                                plochunkSTDandNonStd2.Add(descretionarytextsingle);

                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                                plochunkSTDandNonStd2.Add(GGESInstructions);
                            //plochunkSTDandNonStd2.Add(InstructionsSTDandNonStd33);
                            plochunkSTDandNonStd2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSTDandNonStd2.Leading = 11f;
                            plochunkSTDandNonStd2.SpacingAfter = 12f;
                            document.Add(plochunkSTDandNonStd2);
                        }

                        iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                        iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                        hw.Style = styles;
                        ArrayList objects = null;

                        string FundSpecificDesc = "";

                        for (int m = 0; m < rowsize; m++)
                        {
                            if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                            {
                                FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);

                                FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("<p style=\"margin-left:", "<li style='margin: 0in 0in 0pt 0.5in; ' ");
                                //FundSpecificDesc = FundSpecificDesc.Replace("pt\">  </p>", "</li>");
                                //FundSpecificDesc = FundSpecificDesc.Replace("<p", "<li");
                                //FundSpecificDesc = FundSpecificDesc.Replace("</p>", "</li>");
                                Chunk NextLine1 = new Chunk("\n");
                                Paragraph pNextLine1 = new Paragraph();
                                pNextLine1.Add(NextLine1);
                                pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                //document.Add(pNextLine1);
                                //FundSpecificDesc = RemoveStyle(FundSpecificDesc);

                                FontFactory.RegisterDirectories();
                                Font fontNormal = new Font(FontFactory.GetFont("Verdana", 9, Font.NORMAL));


                                //objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(new StringReader(FundSpecificDesc), null);

                                ////add the collection to the document
                                //for (int k = 0; k < objects.Count; k++)
                                //{

                                //    document.Add((IElement)objects[k]);
                                //}

                                string STYLE_DEFAULT_TYPE = "style";
                                string DOCUMENT_HTML_START = "<html><head></head><body>";
                                string DOCUMENT_HTML_END = "</body></html>";
                                string REGEX_GROUP_SELECTOR = "selector";
                                string REGEX_GROUP_STYLE = "style";

                                //amazing regular expression magic
                                string REGEX_GET_STYLES = @"(?<selector>[^\{\s]+\w+(\s\[^\{\s]+)?)\s?\{(?<style>[^\}]*)\}";

                                foreach (Match match in Regex.Matches(FundSpecificDesc, REGEX_GET_STYLES))
                                {
                                    string selector = match.Groups[REGEX_GROUP_SELECTOR].Value;
                                    string style = match.Groups[REGEX_GROUP_STYLE].Value;
                                    this.AddStyle(selector, style);
                                }

                                string strhtml = "<h5 style='margin: 0in 0in 0pt'><em><u><font color='#e36c0a'>Distribution Letter <o:p></o:p></font></u></em></h5> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>This template would require the following user input:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>As of Date<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Letter Date<o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Specific Text<o:p></o:p></font></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>The dynamic fields encoded into this template are the following fields from the mail record:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Salutation (ssi_salutation_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Full Name (ssi_fullname_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 1 (ssi_addressline1_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 2 (ssi_addressline2_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 3 (ssi_addressline3_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>City (ssi_city_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>State (ssi_stateprovince_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>ZIP Code (ssi_zipcode_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Country/Region (ssi_countryregion_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Dear (ssi_dear_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Legal Entity Name (ssi_legalentitynameid)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Name (ssi_fundname) <o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Nickname <i style='mso-bidi-font-style: normal'>(Need new field on the Fund to store this and ability to make the join back from the mail record to the fund dynamic to be able to retrieve the info)</i><o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Percent Called (ssi_percentcalled_ccsf)<o:p></o:p></font></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><o:p><font size='2'>&nbsp;</font></o:p></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>This template has the following permutations dependent on the mail record data:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l1 level1 lfo3'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Multiple Fund Holdings vs. Single Fund Holding<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Dependent on the number of different funds for a Legal Entity and Recipient the beginning of letter would change<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Multiple fund holdings would include a grid and plural text for the beginning ofthe letter <o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>A single fund holding would just have singular references for the beginning of the letter<o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 1in; mso-add-space: auto'><o:p><font size='2'>&nbsp;</font></o:p></p> <p class='MsoNormal' style='margin: 0in 0in 0pt 0.5in'><span style='color: #31849b; mso-themecolor: accent5; mso-themeshade: 191'><font size='2'>BEGINNING PARAGRAPH &ndash; MULTIPLE FUNDS</font></span><o:p></o:p></p> <p>&nbsp;</p>";

                                //string str = System.Text.RegularExpressions.Regex.Matches(FundSpecificDesc, REGEX_GET_STYLES);
                                MemoryStream output = new MemoryStream();
                                StreamWriter html = new StreamWriter(output, Encoding.UTF8);


                                html.Write(string.Concat(DOCUMENT_HTML_START, FundSpecificDesc, DOCUMENT_HTML_END));
                                html.Close();
                                html.Dispose();

                                MemoryStream generate = new MemoryStream(output.ToArray());
                                StreamReader stringReader = new StreamReader(generate);
                                foreach (object item in iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, styles))
                                {
                                    document.Add((IElement)item);
                                }

                                //cleanup these streams
                                html.Dispose();
                                stringReader.Dispose();
                                output.Dispose();
                                generate.Dispose();

                                //using (StringReader stringReader = new StringReader(FundSpecificDesc))
                                //{
                                //    //List<IElement> parsedList = new List<IElement>();

                                //    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                                //    //document.Open();
                                //    foreach (object item in objects)
                                //    {
                                //        if (item is List)
                                //        {
                                //            List list = item as List;
                                //            list.Autoindent = false;
                                //            list.IndentationLeft = 25f;
                                //            list.SetListSymbol("\u2022                      ");
                                //            list.SymbolIndent = 25f;
                                //            document.Add((IElement)item);
                                //        }


                                //        if (item is Paragraph)
                                //        {
                                //            Paragraph para = item as Paragraph; //setFontsverdana

                                //            if (para.ToArray()[0].ToString() == "ע || para.ToArray()[0].ToString() == "Ǣ || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "آ || para.ToArray()[0].ToString() == "𢠼| para.ToArray()[0].ToString() == "v")
                                //            {
                                //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                //                para.IndentationLeft = 30f;
                                //                document.Add(para);
                                //            }
                                //            else
                                //            {
                                //                //for (int j = 0; j < para.ToArray().Length; j++)
                                //                //{
                                //                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                //                //}
                                //                document.Add(para);
                                //            }
                                //        }


                                //    }
                                //    //document.Close();
                                //}


                                //hw.Parse(new StringReader(FundSpecificDesc));

                                Chunk NextLine = new Chunk("\n");
                                Paragraph pNextLine = new Paragraph();
                                pNextLine.Add(NextLine);
                                pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                // document.Add(pNextLine);
                            }
                        }
                        //document.Add(list);

                        double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                        if (remainingPageSpace < 217.00)
                            document.NewPage();

                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Verify".ToUpper())
                        {
                            Chunk lochunkSingleFundSTD = new Chunk("To wire transfer the necessary funds for your capital call, please sign the enclosed wire request", setFontsAll(9, 0, 0));
                            //Paragraph plochunkSingleFundSTD = new Paragraph();
                            //plochunkSingleFundSTD.Add(lochunkSingleFundSTD);
                            //plochunkSingleFundSTD.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(plochunkSingleFundSTD);

                            Chunk SingleStd1 = new Chunk(" form, and fax it to us at", setFontsAll(9, 0, 0));
                            Chunk lophone1 = new Chunk(" (312)960-0204 ", setFontsAll(9, 0, 0));
                            Chunk SingleStd2 = new Chunk("by ", setFontsAll(9, 0, 0));
                            Chunk SingleStd222 = new Chunk(". We will ensure that your wire is processed by the due date.", setFontsAll(9, 0, 0));
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "")
                            {
                                if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                                {
                                    Chunk lochunkSingleFundSTD1111 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]), setFontsAll(9, 1, 0));
                                    Paragraph plochunkSingleFundSTD1 = new Paragraph();
                                    //  plochunkSingleFundSTD1.Add(lochunkSingleFundSTD);
                                    //  plochunkSingleFundSTD1.Add(SingleStd1);
                                    //   plochunkSingleFundSTD1.Add(lophone1);
                                    //  plochunkSingleFundSTD1.Add(SingleStd2);
                                    // plochunkSingleFundSTD1.Add(lochunkSingleFundSTD1111);
                                    //  plochunkSingleFundSTD1.Add(SingleStd222);
                                    plochunkSingleFundSTD1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    plochunkSingleFundSTD1.SpacingBefore = 11f;
                                    plochunkSingleFundSTD1.Leading = 11f;
                                    document.Add(plochunkSingleFundSTD1);
                                }
                            }


                            //Chunk lochunkSingleFundSTD2 = new Chunk("is processed by the due date.", setFontsAll(11, 0, 0));
                            //Paragraph plochunkSingleFundSTD2 = new Paragraph();
                            //plochunkSingleFundSTD2.Add(lochunkSingleFundSTD2);
                            //plochunkSingleFundSTD2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(plochunkSingleFundSTD2);

                            Chunk lochunkSingleFundSTD3 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                            Chunk lophone11 = new Chunk(" (312)960-0231 ", setFontsAll(9, 1, 0));
                            Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                            Chunk lophone2 = new Chunk("(312)960-0211.", setFontsAll(9, 1, 0));
                            Paragraph plochunkSingleFundSTD3 = new Paragraph();
                            plochunkSingleFundSTD3.Add(lochunkSingleFundSTD3);
                            //plochunkSingleFundSTD3.Add(lophone11);
                            //plochunkSingleFundSTD3.Add(lodesc1);
                            //plochunkSingleFundSTD3.Add(lophone2);
                            plochunkSingleFundSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundSTD3.SpacingBefore = 11f;
                            plochunkSingleFundSTD3.Leading = 11f;
                            document.Add(plochunkSingleFundSTD3);

                            Chunk lochunkSingleFundSTD4 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundSTD4 = new Paragraph();
                            plochunkSingleFundSTD4.Add(lochunkSingleFundSTD4);
                            plochunkSingleFundSTD4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundSTD4.SpacingBefore = 8f;
                            document.Add(plochunkSingleFundSTD4);

                            Paragraph pimage = new Paragraph();
                            Paragraph PSignature = new Paragraph();

                            string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                            try
                            {
                                if (File.Exists(Imagepath))
                                {
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                                else if (!File.Exists(Imagepath))
                                {
                                    Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                            }
                            catch (Exception ex) { }
                            Chunk lochunkSingleFundSTD5 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundSTD5 = new Paragraph();
                            plochunkSingleFundSTD5.Add(lochunkSingleFundSTD5);
                            plochunkSingleFundSTD5.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(plochunkSingleFundSTD5);
                        }
                        else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Verify".ToUpper())
                        {
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Pays from GGES".ToUpper())
                            {
                                //Chunk lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by sending a check or by wiring the funds as noted in the attached", setFontsAll(9, 0, 0));
                                //Chunk lochunkSingleFundNONSTD1 = new Chunk(" payment instructions.", setFontsAll(9, 0, 0));
                                Chunk lochunkSingleFundNONSTD;
                                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Check".ToUpper())
                                {
                                    lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by sending a check as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                                }
                                else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Other Wire".ToUpper())
                                {
                                    lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                                }
                                else
                                {
                                    lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                                }
                                if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                                {
                                    Paragraph plochunkSingleFundNONSTD = new Paragraph();
                                    plochunkSingleFundNONSTD.Add(lochunkSingleFundNONSTD);
                                    //plochunkSingleFundNONSTD.Add(lochunkSingleFundNONSTD1);
                                    plochunkSingleFundNONSTD.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    plochunkSingleFundNONSTD.Leading = 11f;
                                    plochunkSingleFundNONSTD.SpacingBefore = 11f;
                                    document.Add(plochunkSingleFundNONSTD);
                                }
                            }

                            //Chunk lochunkSingleFundNONSTD1 = new Chunk("payment instructions.", setFontsAll(11, 0, 0));
                            //Paragraph plochunkSingleFundNONSTD1 = new Paragraph();
                            //plochunkSingleFundNONSTD1.Add(lochunkSingleFundNONSTD1);
                            //plochunkSingleFundNONSTD1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(plochunkSingleFundNONSTD1);


                            Chunk lochunkSingleFundNONSTD2 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                            Chunk lophone1 = new Chunk(" (312)960-0231 ", setFontsAll(9, 1, 0));
                            Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                            Chunk lophone2 = new Chunk("(312)960-0211.", setFontsAll(9, 1, 0));
                            Paragraph plochunkSingleFundNONSTD2 = new Paragraph();
                            plochunkSingleFundNONSTD2.Add(lochunkSingleFundNONSTD2);
                            //plochunkSingleFundNONSTD2.Add(lophone1);
                            //plochunkSingleFundNONSTD2.Add(lodesc1);
                            // plochunkSingleFundNONSTD2.Add(lophone2);
                            plochunkSingleFundNONSTD2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundNONSTD2.SpacingBefore = 11f;
                            plochunkSingleFundNONSTD2.Leading = 11f;
                            document.Add(plochunkSingleFundNONSTD2);


                            Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
                            plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
                            plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundNONSTD3.SpacingBefore = 8f;
                            document.Add(plochunkSingleFundNONSTD3);


                            iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(2, 2);   // 2 rows, 2 columns 
                            loTable1.Border = 0;
                            loTable1.Cellpadding = 0;
                            loTable1.Cellspacing = 0;
                            iTextSharp.text.Cell loCell1 = new Cell();
                            setTableProperty(loTable1, ReportType.CapitalCallStatementCustom);


                            Paragraph pimage = new Paragraph();
                            Paragraph PSignature = new Paragraph();

                            string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                            try
                            {
                                if (File.Exists(Imagepath))
                                {
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                                else if (!File.Exists(Imagepath))
                                {
                                    Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                            }
                            catch (Exception ex) { }

                            Chunk lochunkSingleFundNONSTD4 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundNONSTD4 = new Paragraph();
                            plochunkSingleFundNONSTD4.Add(lochunkSingleFundNONSTD4);
                            plochunkSingleFundNONSTD4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(plochunkSingleFundNONSTD4);
                        }
                    }
                }
                #endregion
            }

        }
        catch { }


        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }
    #endregion

    #region Capital Call Letter Custom Retirement
    public string GetCapitalCallLetterStatementCustomRetirement()
    {
        string[] CheckString;
        int liPageSize = 39;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.CapitalCallStatementCustom); //"SP_S_CapitalCallLetterCustom 13, '18511645-77CA-E111-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE11-8391-001D09665E8F', 'F7063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);

        var dv = newdataset.Tables[0].DefaultView;
        dv.RowFilter = "RetirementFlg = 1";
        var newDS = new DataSet();
        var newDT = dv.ToTable();
        newDS.Tables.Add(newDT);

        int DSCount = newDS.Tables[0].Rows.Count;

        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();
        rand.Next();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 50, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        PdfWriter pdfwriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        AddFooter(document);

        document.Open();

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\GreshamAdvisors_logo.tif");
        png.SetAbsolutePosition(48, 800);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(22);
        document.Add(png);

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        try
        {
            string effCapitalCallDate = string.Empty;
            if (DSCount > 0)
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "")
                {
                    DateTime dt = Convert.ToDateTime(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]);
                    effCapitalCallDate = dt.ToString("MMMM") + " 1, " + dt.ToString("yyy");
                }
            }

            if (DSCount > 1)
            {
                #region Multiple Standard and Non Standard

                #region Details

                Chunk lochunkAsOfDate = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_LetterDate"]) != "")
                {
                    lochunkAsOfDate = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_LetterDate"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAsOfDate != null)
                {
                    //Chunk lochunkAsOfDate = new Chunk(LetterDate, setFontsAll(11, 0, 0));
                    Paragraph pAsOfDate = new Paragraph();
                    pAsOfDate.Add(lochunkAsOfDate);
                    pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(pAsOfDate);
                }

                Chunk lochunkFullName = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]) != "")
                {
                    lochunkFullName = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkFullName != null)
                {
                    // Chunk lochunkFullName = new Chunk(FullName, setFontsAll(11, 0, 0));
                    Paragraph pFullName = new Paragraph();
                    pFullName.Add(lochunkFullName);
                    pFullName.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pFullName.Leading = 11f;
                    document.Add(pFullName);
                }

                Chunk lochunkAddressLine1 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]) != "")
                {
                    lochunkAddressLine1 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressLine1 != null)
                {
                    //   Chunk lochunkAddressLine1 = new Chunk(AddressLine1, setFontsAll(11, 0, 0));
                    Paragraph pAddressLine1 = new Paragraph();
                    pAddressLine1.Add(lochunkAddressLine1);
                    pAddressLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressLine1.Leading = 11f;
                    document.Add(pAddressLine1);
                }

                Chunk lochunkAddressLine2 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]) != "")
                {
                    lochunkAddressLine2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressLine2 != null)
                {
                    //Chunk lochunkAddressLine2 = new Chunk(AddressLine2, setFontsAll(11, 0, 0));
                    Paragraph pAddressLine2 = new Paragraph();
                    pAddressLine2.Add(lochunkAddressLine2);
                    pAddressLine2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressLine2.Leading = 11f;
                    document.Add(pAddressLine2);
                }

                Chunk lochunkAddressLine3 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]) != "")
                {
                    lochunkAddressLine3 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressLine3 != null)
                {
                    //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                    Paragraph pAddressLine3 = new Paragraph();
                    pAddressLine3.Add(lochunkAddressLine3);
                    pAddressLine3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressLine3.Leading = 11f;
                    document.Add(pAddressLine3);
                }

                Chunk lochunkAddressDetails = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]) != "")
                {
                    lochunkAddressDetails = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) + ", " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkAddressDetails != null)
                {
                    //iTextSharp.text.Font contentFont = iTextSharp.text.FontFactory.GetFont("Verdana", 12, iTextSharp.text.Font.NORMAL);
                    //Chunk lochunkAddressDetails = new Chunk(AddressDetails, setFontsAll(11, 0, 0));
                    Paragraph pAddressDetails = new Paragraph();
                    pAddressDetails.Add(lochunkAddressDetails);
                    pAddressDetails.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pAddressDetails.Leading = 11f;
                    document.Add(pAddressDetails);
                }

                Chunk lochunkCountry = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]) != "")
                {
                    lochunkCountry = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]), setFontsAll(9, 0, 0));
                }
                if (lochunkCountry != null)
                {
                    //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                    Paragraph pCountry3 = new Paragraph();
                    pCountry3.Add(lochunkCountry);
                    pCountry3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pCountry3.Leading = 11f;
                    document.Add(pCountry3);
                }

                Chunk lochunkFullName1 = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]) != "")
                {
                    lochunkFullName1 = new Chunk("RE: " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]), setFontsAll(9, 1, 0));
                }
                if (lochunkFullName1 != null)
                {
                    //Chunk lochunkFullName1 = new Chunk(FullNameBold, setFontsAll(11, 1, 0));
                    Paragraph pFullName1 = new Paragraph();
                    pFullName1.Add(lochunkFullName1);
                    pFullName1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pFullName1.SpacingBefore = 12f;
                    document.Add(pFullName1);
                }

                Chunk lochunkdear = null;
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                {
                    lochunkdear = new Chunk("\nDear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + ":", setFontsAll(9, 0, 0));
                }
                if (lochunkdear != null)
                {
                    //Chunk lochunkdear = new Chunk(dear, setFontsAll(11, 0, 0));
                    Paragraph pdear = new Paragraph();
                    pdear.Add(lochunkdear);
                    pdear.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(pdear);
                }
                #endregion

                //string Instructions = "\nThis letter constitutes a call on following partnerships:";
                Chunk lochunk1 = new Chunk("This letter constitutes a call on following partnerships:", setFontsAll(9, 0, 0));

                Paragraph p1 = new Paragraph();
                p1.Add(lochunk1);
                p1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                p1.SpacingBefore = 5f;
                document.Add(p1);

                iTextSharp.text.Table loTable = new iTextSharp.text.Table(3, newdataset.Tables[0].Rows.Count);  // 2 rows, 3 columns          
                iTextSharp.text.Cell loCell = new Cell();
                lsTotalNumberofColumns = "3";
                setTableProperty(loTable, ReportType.CapitalCallStatementCustom);
                iTextSharp.text.Chunk lochunk = new Chunk();

                int rowsize = newdataset.Tables[0].Rows.Count;
                int colsize = 3;

                //int liTotalPage = (rowsize / liPageSize);
                //int liCurrentPage = 0;
                //if (rowsize % liPageSize != 0)
                //{
                //    liTotalPage = liTotalPage + 1;
                //}
                //else
                //{
                //    liPageSize = 2;
                //    liTotalPage = liTotalPage + 1;
                //}

                // Loop for Rows
                setHeaderCapitalCallStatementCustom(document);
                string strRetirementFlg = "0";
                for (int j = 0; j < rowsize; j++)
                {
                    strRetirementFlg = Convert.ToString(newdataset.Tables[0].Rows[j]["RetirementFlg"]);

                    if (strRetirementFlg == "1")
                    {
                        // Loop for Columns
                        for (int k = 0; k < colsize; k++)
                        {
                            if (k == 0)
                            {
                                lochunk = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_fundname"]), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 9f;
                            }
                            if (k == 1)
                            {
                                lochunk = new Chunk(Percentage(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_percentcalled_ccsf"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell.Leading = 9f;
                            }
                            if (k == 2)
                            {
                                lochunk = new Chunk(RoundUp(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_currentcall_ccsf"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 9f;
                            }
                            loCell.Add(lochunk);
                            loTable.AddCell(loCell);
                        }


                        //document.Add(loTable);
                        //liCurrentPage = liCurrentPage + 1;

                        if (j == newdataset.Tables[0].Rows.Count - 1)
                        {
                            document.Add(loTable);
                            //liCurrentPage = liCurrentPage + 1;
                            //document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                        }

                    }
                }
                getTotalAmountCapitalCallStatementCustom(document, (Convert.ToString(newdataset.Tables[0].Rows[0]["TotalcapitalCall"])));

                Chunk Instructions1 = new Chunk("The enclosed statements provide you the call amounts, as well as detailed information on your commitments and funding percentages. These calls are due on ", setFontsAll(9, 0, 0));

                Chunk GGESInstructions = new Chunk("  Your investment in GP Diversified Growth Strategies has been deducted by the amount of these capital calls, effective " + effCapitalCallDate + ".  No action is required on your part.", setFontsAll(9, 1, 0));

                Chunk descretionarytext = new Chunk(" Your Fidelity account will be debited by the amount of these capital calls on that date. No action is required on your part.", setFontsAll(9, 1, 0));

                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "") //ssi_wireasofdate
                {
                    Chunk lochunk2;
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                        lochunk2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]), setFontsAll(9, 0, 0));
                    else
                        lochunk2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]), setFontsAll(9, 1, 0));

                    Chunk fullstop = new Chunk(".", setFontsAll(9, 0, 0));
                    Paragraph p2 = new Paragraph();
                    p2.Add(Instructions1);
                    p2.Add(lochunk2);
                    p2.Add(fullstop);
                    //  if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) == true)
                    p2.Add(descretionarytext);

                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                        p2.Add(GGESInstructions);
                    p2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    p2.SpacingAfter = 11f;
                    p2.SpacingBefore = 3f;
                    p2.Leading = 11f;
                    document.Add(p2);
                }


                iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                hw.Style = styles;

                //styles.LoadTagStyle("ol", "leading", "16,0");
                //styles.LoadTagStyle("li", "face", "garamond");
                //styles.LoadTagStyle("span", "size", "11pt");
                //styles.LoadTagStyle("body", "font-family", "Verdana");
                //styles.LoadTagStyle("body", "font-size", "11pt");
                //styles.LoadTagStyle(HtmlTags.STRONG, HtmlTags.FONT, "Verdana");
                //styles.LoadTagStyle(HtmlTags.STRONG, HtmlTags.SIZE, "9");

                ArrayList objects = null;

                string FundSpecificDesc = "";
                //List list = new List(List.UNORDERED, 10f);
                //list.SetListSymbol("\u2022");
                //list.IndentationLeft = 30f;
                for (int m = 0; m < rowsize; m++)
                {

                    if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                    {

                        FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");

                        //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                        // FundSpecificDesc = FundSpecificDesc.Replace("<p style=\"margin-left:", "<li style='margin: 0in 0in 0pt 0.5in; ' ");
                        //                        F/undSpecificDesc = FundSpecificDesc.Replace("pt\">  </p>", "</li>");
                        //FundSpecificDesc = FundSpecificDesc.Replace("<p", "<li");
                        //FundSpecificDesc = FundSpecificDesc.Replace("</p>", "</li>");
                        Chunk NextLine1 = new Chunk("\n");
                        Paragraph pNextLine1 = new Paragraph();
                        pNextLine1.Add(NextLine1);
                        pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        //document.Add(pNextLine1);


                        //FontFactory.RegisterDirectories();




                        //string str = System.Text.RegularExpressions.Regex.Matches(FundSpecificDesc, REGEX_GET_STYLES);
                        MemoryStream output = new MemoryStream();
                        StreamWriter html = new StreamWriter(output, Encoding.UTF8);



                        html.Write(string.Concat(DOCUMENT_HTML_START, FundSpecificDesc, DOCUMENT_HTML_END));
                        html.Close();
                        html.Dispose();

                        MemoryStream generate = new MemoryStream(output.ToArray());
                        StreamReader stringReader = new StreamReader(generate);

                        objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(new StringReader(FundSpecificDesc), null);

                        int rowcount = objects.Count;
                        //int colsize = 2;

                        int liTotalPage = (rowcount / liPageSize);
                        int liCurrentPage = 0;
                        if (rowcount % liPageSize != 0)
                        {
                            liTotalPage = liTotalPage + 1;
                        }
                        else
                        {
                            liPageSize = 39;
                            liTotalPage = liTotalPage + 1;
                        }
                        //add the collection to the document
                        for (int k = 0; k < objects.Count; k++)
                        {
                            if (k % liPageSize == 0)
                            {
                                document.Add((IElement)objects[k]);
                                if (k != 0)
                                {
                                    liCurrentPage = liCurrentPage + 1;
                                    document.NewPage();
                                }
                            }
                            else
                            {
                                document.Add((IElement)objects[k]);
                            }
                        }

                        //foreach (object item in iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, styles))
                        //{
                        //    document.Add((IElement)item);
                        //}

                        //cleanup these streams
                        html.Dispose();
                        stringReader.Dispose();
                        output.Dispose();
                        generate.Dispose();

                        //using (StreamReader stringReader = new StreamReader(generate))
                        //{
                        //    //List<IElement> parsedList = new List<IElement>();

                        //    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                        //    //hw.
                        //    //hw.Parse(stringAddReader);

                        //    //document.Open();\xB7
                        //    foreach (object item in objects)
                        //    {
                        //        //document.Add((ITextElementArray)item);
                        //        if (item is List)
                        //        {
                        //            List list = item as List;
                        //            list.Autoindent = false;
                        //            list.IndentationLeft = 25f;
                        //            list.SetListSymbol("\u2022                      ");
                        //            list.SymbolIndent = 25f;
                        //            document.Add((IElement)item);
                        //        }


                        //        if (item is Paragraph)
                        //        {
                        //            Paragraph para = item as Paragraph; //setFontsverdana

                        //            if (para.ToArray()[0].ToString() == "ע || para.ToArray()[0].ToString() == "Ǣ || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "آ || para.ToArray()[0].ToString() == "𢠼| para.ToArray()[0].ToString() == "v")
                        //            {
                        //                //((iTextSharp.text.Chunk)(para.ToArray()[0])).SetGenericTag("\u20AC");
                        //                //((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                        //                //para.IndentationLeft = 30f;
                        //                document.Add(para);
                        //            }
                        //            else
                        //            {
                        //                //for (int j = 0; j < para.ToArray().Length; j++)
                        //                //{
                        //                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                        //                //}
                        //                document.Add(para);
                        //            }
                        //        }
                        //    }
                        //    //document.Close();
                        //}

                        //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                        if (m != rowsize - 1)
                        {
                            Chunk NextLine = new Chunk("\n");
                            Paragraph pNextLine = new Paragraph();
                            pNextLine.Add(NextLine);
                            pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(pNextLine);
                        }
                    }
                }
                //document.Add(list);

                double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                if (remainingPageSpace < 217.00)
                    document.NewPage();

                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Verify".ToUpper())
                {
                    Chunk EndingParaMulStandard1 = new Chunk("To wire transfer necessary funds for capital calls, please sign the enclosed wire request form, and", setFontsAll(9, 0, 0));
                    Chunk ParaMulStandard1 = new Chunk(". We will ensure that your wire is processed by the due date.", setFontsAll(9, 0, 0));
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["TemplateMultiFundDate"]) != "")
                    {
                        //EndingParaMulStandard1 = EndingParaMulStandard1 + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]) + ParaMulStandard1;
                        Chunk lochunk3 = new Chunk(" fax it to us at (312)960-0204 by " + Convert.ToString(newdataset.Tables[0].Rows[0]["TemplateMultiFundDate"]), setFontsAll(9, 1, 0));

                        if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                        {
                            Paragraph EndMulStandard1 = new Paragraph();
                            EndMulStandard1.Add(EndingParaMulStandard1);
                            EndMulStandard1.Add(lochunk3);
                            EndMulStandard1.Add(ParaMulStandard1);

                            EndMulStandard1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            EndMulStandard1.SpacingBefore = 11f;
                            EndMulStandard1.Leading = 11f;
                            document.Add(EndMulStandard1);
                        }
                    }


                    //string EndingParaMulStandard2 = "\nIf you have questions, please call Ted Neild (312)960-0231 or Ben Beavers (312)960-0211";
                    Chunk lophonenumber1 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                    Chunk lodesc1 = new Chunk("or Ben Beavers", setFontsAll(9, 0, 0));
                    Chunk lophonenumber2 = new Chunk("(312)960-0211", setFontsAll(9, 1, 0));
                    Chunk lochunk4 = new Chunk("\n" + lophonenumber1, setFontsAll(9, 0, 0));
                    Paragraph EndMulStandard2 = new Paragraph();
                    EndMulStandard2.Add(lophonenumber1);
                    //EndMulStandard2.Add(lodesc1);
                    //EndMulStandard2.Add(lophonenumber2);
                    //EndMulStandard2.Add(lophonenumber1);
                    //EndMulStandard2.Add(lophonenumber1);
                    //EndMulStandard2.Add(lophonenumber2);
                    EndMulStandard2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndMulStandard2.SpacingBefore = 11f;
                    EndMulStandard2.Leading = 11f;
                    document.Add(EndMulStandard2);


                    // string EndingParaMulStandard3 = "Sincerely Yours,";
                    Chunk lochunk5 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph EndMulStandard3 = new Paragraph();
                    EndMulStandard3.Add(lochunk5);
                    EndMulStandard3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndMulStandard3.SpacingBefore = 8f;
                    document.Add(EndMulStandard3);

                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                    try
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                        else if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = HttpContext.Current.Server.MapPath("") + @"images\ImageNotAvailable.jpg";
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                    }
                    catch (Exception ex) { }

                    //string EndingParaMulStandard4 = "Ted Neild";
                    Chunk lochunk6 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                    Paragraph EndMulStandard4 = new Paragraph();
                    EndMulStandard4.Add(lochunk6);
                    EndMulStandard4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(EndMulStandard4);
                }
                else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Verify".ToUpper())
                {
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["TemplateMultiFundDate"]) != "")
                    {
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Pays from GGES".ToUpper())
                        {
                            //string EndingParaNonStandard1 = "You can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.";
                            //Chunk lochunkNonStd1 = new Chunk("\nYou can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            Chunk lochunkNonStd1;
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Check".ToUpper())
                            {
                                lochunkNonStd1 = new Chunk("You can fund your capital call by sending a check as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            }
                            else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Other Wire".ToUpper())
                            {
                                lochunkNonStd1 = new Chunk("You can fund your capital call by wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            }
                            else
                            {
                                lochunkNonStd1 = new Chunk("You can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                            }

                            if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                            {
                                Paragraph EndNonStandard1 = new Paragraph();
                                EndNonStandard1.Add(lochunkNonStd1);
                                EndNonStandard1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                EndNonStandard1.Leading = 11f;
                                EndNonStandard1.SpacingBefore = 11f;
                                document.Add(EndNonStandard1);
                            }
                        }
                    }
                    //string EndingParaNonStandard2 = "payment instructions.";
                    //Chunk lochunkNonStd2 = new Chunk("\n" + EndingParaNonStandard2, setFontsAll(11, 0, 0));
                    //Paragraph EndNonStandard2 = new Paragraph();
                    //EndNonStandard2.Add(lochunkNonStd2);
                    //EndNonStandard2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //document.Add(EndNonStandard2);


                    //string EndingParaNonStandard3 = "\nIf you have questions, please call Ted Neild (312)960-0231 or Ben Beavers (312)960-0211";
                    Chunk lophone1 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " (312) 960-0200.", setFontsAll(9, 0, 0));
                    Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                    Chunk lophone2 = new Chunk("(312)960-0211.", setFontsAll(9, 1, 0));
                    Chunk lochunkNonStd3 = new Chunk("" + lophone1, setFontsAll(9, 0, 0));
                    Paragraph EndNonStandard3 = new Paragraph();
                    EndNonStandard3.Add(lochunkNonStd3);
                    // EndNonStandard3.Add(lophone1);
                    //EndNonStandard3.Add(lodesc1);
                    //EndNonStandard3.Add(lophone2);
                    EndNonStandard3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndNonStandard3.SpacingBefore = 11f;
                    EndNonStandard3.Leading = 11f;
                    document.Add(EndNonStandard3);



                    Chunk lochunkNonStd4 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph EndNonStandard4 = new Paragraph();
                    EndNonStandard4.Add(lochunkNonStd4);
                    EndNonStandard4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    EndNonStandard4.SpacingBefore = 8f;
                    document.Add(EndNonStandard4);


                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                    try
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                        else if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg";
                            //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);
                        }
                    }
                    catch (Exception ex) { }
                    Chunk lochunkNonStd5 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                    Paragraph EndNonStandard5 = new Paragraph();
                    EndNonStandard5.Add(lochunkNonStd5);
                    EndNonStandard5.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(EndNonStandard5);

                }

                #endregion
            }
            else if (DSCount > 0)
            {
                #region Single Standard and Non Standard
                string strretirementflg = "0";
                for (int i = 0; i < DSCount; i++)
                {
                    strretirementflg = Convert.ToString(newdataset.Tables[0].Rows[i]["RetirementFlg"]);
                    if (strretirementflg == "1")
                    {
                        #region Details
                        lsTotalNumberofColumns = "";
                        Chunk lochunkAsOfDate = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["Ssi_LetterDate"]) != "")
                        {
                            lochunkAsOfDate = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[i]["Ssi_LetterDate"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAsOfDate != null)
                        {
                            //Chunk lochunkAsOfDate = new Chunk(LetterDate, setFontsAll(11, 0, 0));
                            Paragraph pAsOfDate = new Paragraph();
                            pAsOfDate.Add(lochunkAsOfDate);
                            pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(pAsOfDate);
                        }


                        Chunk lochunkFullName = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_salutation_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_fullname_mail"]) != "")
                        {
                            lochunkFullName = new Chunk("\n\n" + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_fullname_mail"]), setFontsAll(9, 0, 0));
                        }

                        if (lochunkFullName != null)
                        {
                            //Chunk lochunkFullName = new Chunk(FullName, setFontsAll(11, 0, 0));
                            Paragraph pFullName = new Paragraph();
                            pFullName.Add(lochunkFullName);
                            pFullName.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pFullName.Leading = 11f;
                            document.Add(pFullName);
                        }

                        Chunk lochunkAddressLine1 = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline1_mail"]) != "")
                        {
                            lochunkAddressLine1 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline1_mail"]), setFontsAll(9, 0, 0));
                        }

                        if (lochunkAddressLine1 != null)
                        {
                            //Chunk lochunkAddressLine1 = new Chunk(AddressLine1, setFontsAll(11, 0, 0));
                            Paragraph pAddressLine1 = new Paragraph();
                            pAddressLine1.Add(lochunkAddressLine1);
                            pAddressLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressLine1.Leading = 11f;
                            document.Add(pAddressLine1);
                        }

                        Chunk lochunkAddressLine2 = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline2_mail"]) != "")
                        {
                            lochunkAddressLine2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline2_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAddressLine2 != null)
                        {
                            //Chunk lochunkAddressLine2 = new Chunk(AddressLine2, setFontsAll(11, 0, 0));
                            Paragraph pAddressLine2 = new Paragraph();
                            pAddressLine2.Add(lochunkAddressLine2);
                            pAddressLine2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressLine2.Leading = 11f;
                            document.Add(pAddressLine2);
                        }

                        Chunk lochunkAddressLine3 = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline3_mail"]) != "")
                        {
                            lochunkAddressLine3 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_addressline3_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAddressLine3 != null)
                        {
                            //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                            Paragraph pAddressLine3 = new Paragraph();
                            pAddressLine3.Add(lochunkAddressLine3);
                            pAddressLine3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressLine3.Leading = 11f;
                            document.Add(pAddressLine3);
                        }

                        Chunk lochunkAddressDetails = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_city_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_stateprovince_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_zipcode_mail"]) != "")
                        {
                            lochunkAddressDetails = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_city_mail"]) + ", " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_stateprovince_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_zipcode_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkAddressDetails != null)
                        {
                            //Chunk lochunkAddressDetails = new Chunk(AddressDetails, setFontsAll(11, 0, 0));
                            Paragraph pAddressDetails = new Paragraph();
                            pAddressDetails.Add(lochunkAddressDetails);
                            pAddressDetails.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pAddressDetails.Leading = 11f;
                            document.Add(pAddressDetails);
                        }

                        Chunk lochunkCountry = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]) != "")
                        {
                            lochunkCountry = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]), setFontsAll(9, 0, 0));
                        }
                        if (lochunkCountry != null)
                        {
                            //Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(11, 0, 0));
                            Paragraph pCountry3 = new Paragraph();
                            pCountry3.Add(lochunkCountry);
                            pCountry3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pCountry3.Leading = 11f;
                            document.Add(pCountry3);
                        }

                        Chunk lochunkFullName1 = null;
                        //string FullNameBold = "\n\n RE: ";
                        if (Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_legalentitynameidname"]) != "")
                        {
                            lochunkFullName1 = new Chunk("RE: " + Convert.ToString(newdataset.Tables[0].Rows[i]["ssi_legalentitynameidname"]), setFontsAll(9, 1, 0));
                        }
                        if (lochunkFullName1 != null)
                        {
                            //Chunk lochunkFullName1 = new Chunk(FullNameBold, setFontsAll(11, 0, 0));
                            Paragraph pFullName1 = new Paragraph();
                            pFullName1.Add(lochunkFullName1);
                            pFullName1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            pFullName1.SpacingBefore = 12f;
                            document.Add(pFullName1);
                        }

                        Chunk lochunkdear = null;
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                        {
                            lochunkdear = new Chunk("\nDear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + ":", setFontsAll(9, 0, 0));
                        }
                        if (lochunkdear != null)
                        {
                            //Chunk lochunkdear = new Chunk(dear, setFontsAll(11, 0, 0));
                            Paragraph pdear = new Paragraph();
                            pdear.Add(lochunkdear);
                            pdear.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(pdear);
                        }


                        #endregion


                        iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, newdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                        iTextSharp.text.Cell loCell = new Cell();

                        iTextSharp.text.Chunk lochunk = new Chunk();

                        int rowsize = newdataset.Tables[0].Rows.Count;


                        //string InstructionsSTDandNonStd = "\nThis letter constitutes a ";
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) != "" && Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_percentcalled_ccsf"]) != "" && Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_currentcall_ccsf"]) != "")
                        {
                            //InstructionsSTDandNonStd = "\nThis letter constitutes a " + Percentage(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_percentcalled_ccsf"])) + " call on your commitment to " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) + ".";
                            Chunk lochunkSTDandNonStd = new Chunk("This letter constitutes a " + Percentage(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_percentcalled_ccsf"])) + " call on your commitment to " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) + ".", setFontsAll(9, 0, 0));
                            Paragraph plochunkSTDandNonStd = new Paragraph();
                            plochunkSTDandNonStd.Add(lochunkSTDandNonStd);
                            plochunkSTDandNonStd.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSTDandNonStd.SpacingBefore = 12f;
                            plochunkSTDandNonStd.Leading = 11f;
                            document.Add(plochunkSTDandNonStd);
                        }


                        //string InstructionsSTDandNonStd1 = "\nThe enclosed statement provides you the call amount, as well as detailed information on";
                        Chunk lochunkSTDandNonStd1 = new Chunk("The enclosed statement provides you the call amount, as well as detailed information on commitment and", setFontsAll(9, 0, 0));

                        Chunk GGESInstructions = new Chunk("  Your investment in GP Diversified Growth Strategies has been deducted by the amount of this capital call, effective " + effCapitalCallDate + ".  No action is required on your part.", setFontsAll(9, 1, 0));

                        Chunk descretionarytextsingle = new Chunk(" Your Fidelity account will be debited by the amount of this capital call on that date. No action is required on your part.", setFontsAll(9, 1, 0));

                        Paragraph plochunkSTDandNonStd1 = new Paragraph();
                        plochunkSTDandNonStd1.Add(lochunkSTDandNonStd1);
                        plochunkSTDandNonStd1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        plochunkSTDandNonStd1.SpacingBefore = 12f;
                        plochunkSTDandNonStd1.Leading = 11f;
                        document.Add(plochunkSTDandNonStd1);


                        Chunk InstructionsSTDandNonStd2 = new Chunk("funding percentage. This call is due on ", setFontsAll(9, 0, 0));
                        Chunk InstructionsSTDandNonStd33 = new Chunk(".", setFontsAll(9, 0, 0));
                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "")
                        {
                            //InstructionsSTDandNonStd2 = InstructionsSTDandNonStd2 + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]);
                            Chunk lochunkSTDandNonStd2;
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                                lochunkSTDandNonStd2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) + "", setFontsAll(9, 0, 0));
                            else
                                lochunkSTDandNonStd2 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) + "", setFontsAll(9, 1, 0));

                            Chunk lochunkSTDandNonStd222 = new Chunk(".", setFontsAll(9, 0, 0));
                            Paragraph plochunkSTDandNonStd2 = new Paragraph();
                            plochunkSTDandNonStd2.Add(InstructionsSTDandNonStd2);
                            plochunkSTDandNonStd2.Add(lochunkSTDandNonStd2);
                            plochunkSTDandNonStd2.Add(lochunkSTDandNonStd222);
                            //  if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) == true)
                            plochunkSTDandNonStd2.Add(descretionarytextsingle);

                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Pays from GGES".ToUpper())
                                plochunkSTDandNonStd2.Add(GGESInstructions);
                            //plochunkSTDandNonStd2.Add(InstructionsSTDandNonStd33);
                            plochunkSTDandNonStd2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSTDandNonStd2.Leading = 11f;
                            plochunkSTDandNonStd2.SpacingAfter = 12f;
                            document.Add(plochunkSTDandNonStd2);
                        }

                        iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                        iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                        hw.Style = styles;
                        ArrayList objects = null;

                        string FundSpecificDesc = "";

                        for (int m = 0; m < rowsize; m++)
                        {
                            if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                            {
                                FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);

                                FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("<p style=\"margin-left:", "<li style='margin: 0in 0in 0pt 0.5in; ' ");
                                //FundSpecificDesc = FundSpecificDesc.Replace("pt\">  </p>", "</li>");
                                //FundSpecificDesc = FundSpecificDesc.Replace("<p", "<li");
                                //FundSpecificDesc = FundSpecificDesc.Replace("</p>", "</li>");
                                Chunk NextLine1 = new Chunk("\n");
                                Paragraph pNextLine1 = new Paragraph();
                                pNextLine1.Add(NextLine1);
                                pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                //document.Add(pNextLine1);
                                //FundSpecificDesc = RemoveStyle(FundSpecificDesc);

                                FontFactory.RegisterDirectories();
                                Font fontNormal = new Font(FontFactory.GetFont("Verdana", 9, Font.NORMAL));


                                //objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(new StringReader(FundSpecificDesc), null);

                                ////add the collection to the document
                                //for (int k = 0; k < objects.Count; k++)
                                //{

                                //    document.Add((IElement)objects[k]);
                                //}

                                string STYLE_DEFAULT_TYPE = "style";
                                string DOCUMENT_HTML_START = "<html><head></head><body>";
                                string DOCUMENT_HTML_END = "</body></html>";
                                string REGEX_GROUP_SELECTOR = "selector";
                                string REGEX_GROUP_STYLE = "style";

                                //amazing regular expression magic
                                string REGEX_GET_STYLES = @"(?<selector>[^\{\s]+\w+(\s\[^\{\s]+)?)\s?\{(?<style>[^\}]*)\}";

                                foreach (Match match in Regex.Matches(FundSpecificDesc, REGEX_GET_STYLES))
                                {
                                    string selector = match.Groups[REGEX_GROUP_SELECTOR].Value;
                                    string style = match.Groups[REGEX_GROUP_STYLE].Value;
                                    this.AddStyle(selector, style);
                                }

                                string strhtml = "<h5 style='margin: 0in 0in 0pt'><em><u><font color='#e36c0a'>Distribution Letter <o:p></o:p></font></u></em></h5> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>This template would require the following user input:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>As of Date<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Letter Date<o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l2 level1 lfo1'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Specific Text<o:p></o:p></font></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>The dynamic fields encoded into this template are the following fields from the mail record:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Salutation (ssi_salutation_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Full Name (ssi_fullname_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 1 (ssi_addressline1_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 2 (ssi_addressline2_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Address Line 3 (ssi_addressline3_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>City (ssi_city_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>State (ssi_stateprovince_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>ZIP Code (ssi_zipcode_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Country/Region (ssi_countryregion_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Dear (ssi_dear_mail)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Legal Entity Name (ssi_legalentitynameid)<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Name (ssi_fundname) <o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Fund Nickname <i style='mso-bidi-font-style: normal'>(Need new field on the Fund to store this and ability to make the join back from the mail record to the fund dynamic to be able to retrieve the info)</i><o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l0 level1 lfo2'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Percent Called (ssi_percentcalled_ccsf)<o:p></o:p></font></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><o:p><font size='2'>&nbsp;</font></o:p></p> <p class='MsoNormal' style='margin: 0in 0in 0pt'><font size='2'>This template has the following permutations dependent on the mail record data:<o:p></o:p></font></p> <p class='MsoListParagraphCxSpFirst' style='margin: 0in 0in 0pt 0.5in; text-indent: -0.25in; mso-list: l1 level1 lfo3'><span style='font-family: Symbol; mso-fareast-font-family: Symbol; mso-bidi-font-family: Symbol'><span style='mso-list: Ignore'><font size='2'>&middot;</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Multiple Fund Holdings vs. Single Fund Holding<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Dependent on the number of different funds for a Legal Entity and Recipient the beginning of letter would change<o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>Multiple fund holdings would include a grid and plural text for the beginning ofthe letter <o:p></o:p></font></p> <p class='MsoListParagraphCxSpMiddle' style='margin: 0in 0in 0pt 1in; text-indent: -0.25in; mso-add-space: auto; mso-list: l1 level2 lfo3'><span style='font-family: &quot;Courier New&quot;; mso-fareast-font-family: 'Courier New''><span style='mso-list: Ignore'><font size='2'>o</font><span style='font: 7pt &quot;Times New Roman&quot;'>&nbsp;&nbsp;&nbsp; </span></span></span><font size='2'>A single fund holding would just have singular references for the beginning of the letter<o:p></o:p></font></p> <p class='MsoListParagraphCxSpLast' style='margin: 0in 0in 0pt 1in; mso-add-space: auto'><o:p><font size='2'>&nbsp;</font></o:p></p> <p class='MsoNormal' style='margin: 0in 0in 0pt 0.5in'><span style='color: #31849b; mso-themecolor: accent5; mso-themeshade: 191'><font size='2'>BEGINNING PARAGRAPH &ndash; MULTIPLE FUNDS</font></span><o:p></o:p></p> <p>&nbsp;</p>";

                                //string str = System.Text.RegularExpressions.Regex.Matches(FundSpecificDesc, REGEX_GET_STYLES);
                                MemoryStream output = new MemoryStream();
                                StreamWriter html = new StreamWriter(output, Encoding.UTF8);


                                html.Write(string.Concat(DOCUMENT_HTML_START, FundSpecificDesc, DOCUMENT_HTML_END));
                                html.Close();
                                html.Dispose();

                                MemoryStream generate = new MemoryStream(output.ToArray());
                                StreamReader stringReader = new StreamReader(generate);
                                foreach (object item in iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, styles))
                                {
                                    document.Add((IElement)item);
                                }

                                //cleanup these streams
                                html.Dispose();
                                stringReader.Dispose();
                                output.Dispose();
                                generate.Dispose();

                                //using (StringReader stringReader = new StringReader(FundSpecificDesc))
                                //{
                                //    //List<IElement> parsedList = new List<IElement>();

                                //    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                                //    //document.Open();
                                //    foreach (object item in objects)
                                //    {
                                //        if (item is List)
                                //        {
                                //            List list = item as List;
                                //            list.Autoindent = false;
                                //            list.IndentationLeft = 25f;
                                //            list.SetListSymbol("\u2022                      ");
                                //            list.SymbolIndent = 25f;
                                //            document.Add((IElement)item);
                                //        }


                                //        if (item is Paragraph)
                                //        {
                                //            Paragraph para = item as Paragraph; //setFontsverdana

                                //            if (para.ToArray()[0].ToString() == "ע || para.ToArray()[0].ToString() == "Ǣ || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "آ || para.ToArray()[0].ToString() == "𢠼| para.ToArray()[0].ToString() == "v")
                                //            {
                                //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                //                para.IndentationLeft = 30f;
                                //                document.Add(para);
                                //            }
                                //            else
                                //            {
                                //                //for (int j = 0; j < para.ToArray().Length; j++)
                                //                //{
                                //                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                //                //}
                                //                document.Add(para);
                                //            }
                                //        }


                                //    }
                                //    //document.Close();
                                //}


                                //hw.Parse(new StringReader(FundSpecificDesc));

                                Chunk NextLine = new Chunk("\n");
                                Paragraph pNextLine = new Paragraph();
                                pNextLine.Add(NextLine);
                                pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                // document.Add(pNextLine);
                            }
                        }
                        //document.Add(list);

                        double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                        if (remainingPageSpace < 217.00)
                            document.NewPage();

                        if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Verify".ToUpper())
                        {
                            Chunk lochunkSingleFundSTD = new Chunk("To wire transfer the necessary funds for your capital call, please sign the enclosed wire request", setFontsAll(9, 0, 0));
                            //Paragraph plochunkSingleFundSTD = new Paragraph();
                            //plochunkSingleFundSTD.Add(lochunkSingleFundSTD);
                            //plochunkSingleFundSTD.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(plochunkSingleFundSTD);

                            Chunk SingleStd1 = new Chunk(" form, and fax it to us at", setFontsAll(9, 0, 0));
                            Chunk lophone1 = new Chunk(" (312)960-0204 ", setFontsAll(9, 0, 0));
                            Chunk SingleStd2 = new Chunk("by ", setFontsAll(9, 0, 0));
                            Chunk SingleStd222 = new Chunk(". We will ensure that your wire is processed by the due date.", setFontsAll(9, 0, 0));
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]) != "")
                            {
                                if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                                {
                                    Chunk lochunkSingleFundSTD1111 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_CapitalCallDate"]), setFontsAll(9, 1, 0));
                                    Paragraph plochunkSingleFundSTD1 = new Paragraph();
                                    plochunkSingleFundSTD1.Add(lochunkSingleFundSTD);
                                    plochunkSingleFundSTD1.Add(SingleStd1);
                                    plochunkSingleFundSTD1.Add(lophone1);
                                    plochunkSingleFundSTD1.Add(SingleStd2);
                                    plochunkSingleFundSTD1.Add(lochunkSingleFundSTD1111);
                                    plochunkSingleFundSTD1.Add(SingleStd222);
                                    plochunkSingleFundSTD1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    plochunkSingleFundSTD1.SpacingBefore = 11f;
                                    plochunkSingleFundSTD1.Leading = 11f;
                                    document.Add(plochunkSingleFundSTD1);
                                }
                            }


                            //Chunk lochunkSingleFundSTD2 = new Chunk("is processed by the due date.", setFontsAll(11, 0, 0));
                            //Paragraph plochunkSingleFundSTD2 = new Paragraph();
                            //plochunkSingleFundSTD2.Add(lochunkSingleFundSTD2);
                            //plochunkSingleFundSTD2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(plochunkSingleFundSTD2);

                            Chunk lochunkSingleFundSTD3 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                            Chunk lophone11 = new Chunk(" (312)960-0231 ", setFontsAll(9, 1, 0));
                            Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                            Chunk lophone2 = new Chunk("(312)960-0211.", setFontsAll(9, 1, 0));
                            Paragraph plochunkSingleFundSTD3 = new Paragraph();
                            plochunkSingleFundSTD3.Add(lochunkSingleFundSTD3);
                            //plochunkSingleFundSTD3.Add(lophone11);
                            //plochunkSingleFundSTD3.Add(lodesc1);
                            //plochunkSingleFundSTD3.Add(lophone2);
                            plochunkSingleFundSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundSTD3.SpacingBefore = 11f;
                            plochunkSingleFundSTD3.Leading = 11f;
                            document.Add(plochunkSingleFundSTD3);

                            Chunk lochunkSingleFundSTD4 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundSTD4 = new Paragraph();
                            plochunkSingleFundSTD4.Add(lochunkSingleFundSTD4);
                            plochunkSingleFundSTD4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundSTD4.SpacingBefore = 8f;
                            document.Add(plochunkSingleFundSTD4);

                            Paragraph pimage = new Paragraph();
                            Paragraph PSignature = new Paragraph();

                            string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                            try
                            {
                                if (File.Exists(Imagepath))
                                {
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                                else if (!File.Exists(Imagepath))
                                {
                                    Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                            }
                            catch (Exception ex) { }
                            Chunk lochunkSingleFundSTD5 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundSTD5 = new Paragraph();
                            plochunkSingleFundSTD5.Add(lochunkSingleFundSTD5);
                            plochunkSingleFundSTD5.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(plochunkSingleFundSTD5);
                        }
                        else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Account".ToUpper() || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Verify".ToUpper())
                        {
                            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() != "Pays from GGES".ToUpper())
                            {
                                //Chunk lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by sending a check or by wiring the funds as noted in the attached", setFontsAll(9, 0, 0));
                                //Chunk lochunkSingleFundNONSTD1 = new Chunk(" payment instructions.", setFontsAll(9, 0, 0));
                                Chunk lochunkSingleFundNONSTD;
                                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Check".ToUpper())
                                {
                                    lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by sending a check as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                                }
                                else if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_capitalcallpaymentmethod"]).ToUpper() == "Other Wire".ToUpper())
                                {
                                    lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                                }
                                else
                                {
                                    lochunkSingleFundNONSTD = new Chunk("You can fund your capital call by sending a check or wiring the funds as noted in the attached payment instructions.", setFontsAll(9, 0, 0));
                                }
                                if (Convert.ToBoolean(newdataset.Tables[0].Rows[0]["ssi_discretionaryflg"]) != true)
                                {
                                    Paragraph plochunkSingleFundNONSTD = new Paragraph();
                                    plochunkSingleFundNONSTD.Add(lochunkSingleFundNONSTD);
                                    //plochunkSingleFundNONSTD.Add(lochunkSingleFundNONSTD1);
                                    plochunkSingleFundNONSTD.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    plochunkSingleFundNONSTD.Leading = 11f;
                                    plochunkSingleFundNONSTD.SpacingBefore = 11f;
                                    document.Add(plochunkSingleFundNONSTD);
                                }
                            }

                            //Chunk lochunkSingleFundNONSTD1 = new Chunk("payment instructions.", setFontsAll(11, 0, 0));
                            //Paragraph plochunkSingleFundNONSTD1 = new Paragraph();
                            //plochunkSingleFundNONSTD1.Add(lochunkSingleFundNONSTD1);
                            //plochunkSingleFundNONSTD1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(plochunkSingleFundNONSTD1);


                            Chunk lochunkSingleFundNONSTD2 = new Chunk("If you have questions, please call " + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                            Chunk lophone1 = new Chunk(" (312)960-0231 ", setFontsAll(9, 1, 0));
                            Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                            Chunk lophone2 = new Chunk("(312)960-0211.", setFontsAll(9, 1, 0));
                            Paragraph plochunkSingleFundNONSTD2 = new Paragraph();
                            plochunkSingleFundNONSTD2.Add(lochunkSingleFundNONSTD2);
                            //plochunkSingleFundNONSTD2.Add(lophone1);
                            //plochunkSingleFundNONSTD2.Add(lodesc1);
                            // plochunkSingleFundNONSTD2.Add(lophone2);
                            plochunkSingleFundNONSTD2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundNONSTD2.SpacingBefore = 11f;
                            plochunkSingleFundNONSTD2.Leading = 11f;
                            document.Add(plochunkSingleFundNONSTD2);


                            Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
                            plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
                            plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            plochunkSingleFundNONSTD3.SpacingBefore = 8f;
                            document.Add(plochunkSingleFundNONSTD3);


                            iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(2, 2);   // 2 rows, 2 columns 
                            loTable1.Border = 0;
                            loTable1.Cellpadding = 0;
                            loTable1.Cellspacing = 0;
                            iTextSharp.text.Cell loCell1 = new Cell();
                            setTableProperty(loTable1, ReportType.CapitalCallStatementCustom);


                            Paragraph pimage = new Paragraph();
                            Paragraph PSignature = new Paragraph();

                            string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                            try
                            {
                                if (File.Exists(Imagepath))
                                {
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                                else if (!File.Exists(Imagepath))
                                {
                                    Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                    SignatureJpg.ScaleToFit(250, 42);
                                    pimage.Add(SignatureJpg);
                                    document.Add(pimage);
                                }
                            }
                            catch (Exception ex) { }

                            Chunk lochunkSingleFundNONSTD4 = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                            Paragraph plochunkSingleFundNONSTD4 = new Paragraph();
                            plochunkSingleFundNONSTD4.Add(lochunkSingleFundNONSTD4);
                            plochunkSingleFundNONSTD4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            document.Add(plochunkSingleFundNONSTD4);
                        }
                    }
                }
                #endregion
            }

        }
        catch { }


        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }
    #endregion

    #region Distribution Letter Custom
    public string GetDistributionLetterStatementCustom()
    {
        string[] CheckString;
        int liPageSize = 40;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.DistributionLetterCustom); //"SP_S_DistributionLetterCustom 13, '1859645-77CA-E91-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE9-8391-001D09665E8F', 'F7063302-DD15-DE9-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 50, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "_.pdf";
        PdfWriter pdfwriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        // AddHeader(document);
        AddFooter(document);

        document.Open();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        try
        {

            string effDistributionDate = string.Empty;
            string WireDt = string.Empty;
            string DateSuperScript = string.Empty;
            if (DSCount > 0)
            {
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_WireAsOfDate"]) != "")
                {
                    DateTime dt = Convert.ToDateTime(newdataset.Tables[0].Rows[0]["Ssi_WireAsOfDate"]);

                    string lastdate = DateTime.DaysInMonth(dt.Year, dt.Month).ToString();

                    effDistributionDate = dt.ToString("MMMM") + " " + lastdate + ", " + dt.ToString("yyy");
                    WireDt = dt.ToString("MMMM") + " " + dt.ToString("dd");
                    DateSuperScript = GetDateSuperScript(dt.ToString("dd"));
                }
            }

            if (DSCount > 0)
            {

                #region Details

                string AsOfDate = "\n\n";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_LetterDate"]) != "")
                {
                    AsOfDate = AsOfDate + Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_LetterDate"]);
                }
                Chunk lochunkAsOfDate = new Chunk(AsOfDate, setFontsAll(9, 0, 0));
                Paragraph pAsOfDate = new Paragraph();
                pAsOfDate.Add(lochunkAsOfDate);
                pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                document.Add(pAsOfDate);


                string FullName = "\n\n";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]) != "")
                {
                    FullName = FullName + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]);
                }
                Chunk lochunkFullName = new Chunk(FullName, setFontsAll(9, 0, 0));
                Paragraph pFullName = new Paragraph();
                pFullName.Add(lochunkFullName);
                pFullName.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pFullName.Leading = 11f;
                document.Add(pFullName);

                string AddressLine1 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]) != "")
                {
                    AddressLine1 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]);
                }
                Chunk lochunkAddressLine1 = new Chunk(AddressLine1, setFontsAll(9, 0, 0));
                Paragraph pAddressLine1 = new Paragraph();
                pAddressLine1.Add(lochunkAddressLine1);
                pAddressLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine1.Leading = 11f;
                document.Add(pAddressLine1);

                string AddressLine2 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]) != "")
                {
                    AddressLine2 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]);
                }
                Chunk lochunkAddressLine2 = new Chunk(AddressLine2, setFontsAll(9, 0, 0));
                Paragraph pAddressLine2 = new Paragraph();
                pAddressLine2.Add(lochunkAddressLine2);
                pAddressLine2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine2.Leading = 11f;
                document.Add(pAddressLine2);

                string AddressLine3 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]) != "")
                {
                    AddressLine3 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]);
                }
                Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(9, 0, 0));
                Paragraph pAddressLine3 = new Paragraph();
                pAddressLine3.Add(lochunkAddressLine3);
                pAddressLine3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine3.Leading = 11f;
                document.Add(pAddressLine3);

                string AddressDetails = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]) != "")
                {
                    AddressDetails = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) + ", " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]);
                }
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]) != "")
                {
                    AddressDetails = AddressDetails + "\n" + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Countryregion_mail"]);
                }
                Chunk lochunkAddressDetails = new Chunk(AddressDetails, setFontsAll(9, 0, 0));
                Paragraph pAddressDetails = new Paragraph();
                pAddressDetails.Add(lochunkAddressDetails);
                pAddressDetails.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressDetails.Leading = 11f;
                document.Add(pAddressDetails);


                string FullNameBold = "RE: ";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]) != "")
                {
                    FullNameBold = FullNameBold + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]);
                }
                Chunk lochunkFullName1 = new Chunk(FullNameBold, setFontsAll(9, 1, 0));
                Paragraph pFullName1 = new Paragraph();
                pFullName1.Add(lochunkFullName1);
                pFullName1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pFullName1.SpacingBefore = 12f;
                document.Add(pFullName1);



                #endregion

                iTextSharp.text.Table loTable = new iTextSharp.text.Table(3, newdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                iTextSharp.text.Cell loCell = new Cell();

                iTextSharp.text.Chunk lochunk = new Chunk();

                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\GreshamAdvisors_logo.tif");
                png.SetAbsolutePosition(48, 800);//540
                //png.ScaleToFit(288f, 42f);
                png.ScalePercent(22);
                document.Add(png);

                int rowsize = newdataset.Tables[0].Rows.Count;
                int colsize = 3;

                if (rowsize > 1)
                {
                    lsTotalNumberofColumns = "3";
                    setTableProperty(loTable, ReportType.DistributionLetterCustom);
                }
                else
                {
                    lsTotalNumberofColumns = "";
                    setTableProperty(loTable, ReportType.DistributionLetterCustom);
                }


                //int liTotalPage = (rowsize / liPageSize);
                //int liCurrentPage = 0;
                //if (rowsize % liPageSize != 0)
                //{
                //    liTotalPage = liTotalPage + 1;
                //}
                //else
                //{
                //    liPageSize = 2;
                //    liTotalPage = liTotalPage + 1;
                //}

                if (rowsize > 1)
                {
                    #region Multiple Distribution Letter Funds


                    string Instructions = "";
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                    {
                        Instructions = "Dear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + ":";
                    }

                    Chunk lochunk1 = new Chunk(Instructions, setFontsAll(9, 0, 0));
                    Paragraph p1 = new Paragraph();
                    p1.Add(lochunk1);
                    p1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    p1.SpacingBefore = 5f;
                    document.Add(p1);

                    Paragraph ParaStart = new Paragraph();
                    Chunk lochunkStart;
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_distributionpaymentmethod"]).ToUpper() == "Pays to GGES".ToUpper())
                        lochunkStart = new Chunk("We will make an additional investment in GP Diversified Growth Strategies on " + effDistributionDate + ", representing your share of distributions from the following partnerships: ", setFontsAll(9, 0, 0));
                    else
                    {
                        //lochunkStart = new Chunk("We have wired cash distributions to your account today, representing your share of distributions from the following partnerships: ", setFontsAll(9, 0, 0));

                        lochunkStart = new Chunk("Cash distributions will be wired to your account the week of " + WireDt + "", setFontsAll(9, 0, 0));
                        ParaStart.Add(lochunkStart);
                        Chunk superScript = new Chunk(DateSuperScript, setFontsAll(6, 0, 0));
                        superScript.SetTextRise(5f);
                        ParaStart.Add(superScript);
                        lochunkStart = new Chunk(", representing your share of distributions from the following partnerships: ", setFontsAll(9, 0, 0));
                    }

                    ParaStart.Add(lochunkStart);
                    ParaStart.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    ParaStart.SpacingBefore = 9f;
                    ParaStart.Leading = 11f;
                    document.Add(ParaStart);

                    //Chunk lochunkStart1 = new Chunk("the following partnerships: ", setFontsAll(9, 0, 0));
                    //Paragraph ParaStart1 = new Paragraph();
                    //ParaStart1.Add(lochunkStart1);
                    //ParaStart1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //ParaStart1.Leading = 11f;
                    //document.Add(ParaStart1);

                    setHeaderDistributionStatementCustom(document);//Header
                    for (int j = 0; j < rowsize; j++)
                    {

                        // Loop for Columns
                        for (int k = 0; k < colsize; k++)
                        {
                            if (k == 0)
                            {
                                lochunk = new Chunk(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_fundname"]), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell.Leading = 9f;
                            }
                            if (k == 1)
                            {
                                lochunk = new Chunk(Percentage(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_curdistp_db"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell.Leading = 9f;
                            }

                            if (k == 2)
                            {
                                lochunk = new Chunk(RoundUp(Convert.ToString(newdataset.Tables[0].Rows[j]["ssi_capitaldistribution_db"])), setFontsAll(9, 0, 0));
                                loCell = new iTextSharp.text.Cell();

                                loCell.Border = 0;
                                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                                loCell.Leading = 9f;
                            }

                            loCell.Add(lochunk);
                            loTable.AddCell(loCell);
                        }
                        //document.Add(loTable);
                        //liCurrentPage = liCurrentPage + 1;

                        if (j == newdataset.Tables[0].Rows.Count - 1)
                        {
                            // Below code to show the total of dollar amount in last row.
                            lochunk = new Chunk("", setFontsAll(9, 0, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Border = 0;
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            loCell.Leading = 9f;
                            loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                            loCell.Add(lochunk);
                            loTable.AddCell(loCell);

                            lochunk = new Chunk("Total", setFontsAll(9, 1, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Border = 0;
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell.Leading = 9f;
                            loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                            loCell.Add(lochunk);
                            loTable.AddCell(loCell);

                            lochunk = new Chunk(RoundUp(Convert.ToString(newdataset.Tables[0].Rows[j]["Totalcapitaldistribution"])), setFontsAll(9, 1, 0));
                            loCell = new iTextSharp.text.Cell();
                            loCell.Border = 0;
                            loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
                            loCell.Leading = 9f;
                            loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
                            loCell.Add(lochunk);
                            loTable.AddCell(loCell);

                            document.Add(loTable);
                            //liCurrentPage = liCurrentPage + 1;
                            //document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                        }
                    }
                    //document.Add(loTable);

                    Chunk lochunkEnd1 = new Chunk("The enclosed statements provide you with detailed information regarding your distribution", setFontsAll(9, 0, 0));
                    Paragraph ParaEnd1 = new Paragraph();
                    ParaEnd1.Add(lochunkEnd1);
                    ParaEnd1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    ParaEnd1.Leading = 11f;
                    ParaEnd1.SpacingBefore = 7f;
                    document.Add(ParaEnd1);


                    Chunk lochunkEnd2 = new Chunk("amounts, commitments, and funding status.", setFontsAll(9, 0, 0));
                    Paragraph ParaEnd2 = new Paragraph();
                    ParaEnd2.Add(lochunkEnd2);
                    ParaEnd2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    ParaEnd2.Leading = 11f;
                    ParaEnd2.SpacingAfter = 11f;
                    document.Add(ParaEnd2);

                    iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                    iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                    hw.Style = styles;
                    //styles.LoadTagStyle("ol", "leading", "16,0");
                    //styles.LoadTagStyle("li", "face", "garamond");
                    //styles.LoadTagStyle("span", "size", "11pt");
                    //styles.LoadTagStyle("body", "font-family", "verdana");
                    //styles.LoadTagStyle("body", "font-size", "11pt");

                    ArrayList objects = null;

                    string FundSpecificDesc = "";
                    //List list = new List(List.UNORDERED, 10f);
                    //list.SetListSymbol("\u2022");
                    //list.IndentationLeft = 30f;
                    for (int m = 0; m < rowsize; m++)
                    {
                        if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                        {

                            FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);
                            FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                            //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                            //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                            FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                            FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                            Chunk NextLine1 = new Chunk("\n");
                            Paragraph pNextLine1 = new Paragraph();
                            pNextLine1.Add(NextLine1);
                            pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(pNextLine1);


                            MemoryStream output = new MemoryStream();
                            StreamWriter html = new StreamWriter(output, Encoding.UTF8);


                            html.Write(string.Concat(DOCUMENT_HTML_START, FundSpecificDesc, DOCUMENT_HTML_END));
                            html.Close();
                            html.Dispose();

                            MemoryStream generate = new MemoryStream(output.ToArray());
                            StreamReader stringReader = new StreamReader(generate);


                            objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(new StringReader(FundSpecificDesc), null);

                            int rowcount = objects.Count;
                            //int colsize = 2;

                            int liTotalPage = (rowcount / liPageSize);
                            int liCurrentPage = 0;
                            if (rowcount % liPageSize != 0)
                            {
                                liTotalPage = liTotalPage + 1;
                            }
                            else
                            {
                                liPageSize = 40;
                                liTotalPage = liTotalPage + 1;
                            }
                            //add the collection to the document
                            for (int k = 0; k < objects.Count; k++)
                            {
                                if (k % liPageSize == 0)
                                {
                                    document.Add((IElement)objects[k]);
                                    if (k != 0)
                                    {
                                        liCurrentPage = liCurrentPage + 1;
                                        document.NewPage();
                                    }
                                }
                                else
                                {
                                    document.Add((IElement)objects[k]);
                                }
                            }

                            //foreach (object item in iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, styles))
                            //{
                            //    document.Add((IElement)item);
                            //}

                            //cleanup these streams
                            html.Dispose();
                            stringReader.Dispose();
                            output.Dispose();
                            generate.Dispose();

                            //using (StringReader stringReader = new StringReader(FundSpecificDesc))
                            //{
                            //    //List<IElement> parsedList = new List<IElement>();

                            //    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                            //    //document.Open();
                            //    foreach (object item in objects)
                            //    {
                            //        if (item is List)
                            //        {
                            //            List list = item as List;
                            //            list.Autoindent = false;
                            //            list.IndentationLeft = 25f;
                            //            list.SetListSymbol("\u2022                      ");
                            //            list.SymbolIndent = 25f;
                            //            document.Add((IElement)item);
                            //        }


                            //        if (item is Paragraph)
                            //        {
                            //            Paragraph para = item as Paragraph; //setFontsverdana

                            //            if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                            //            {
                            //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                            //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                            //                para.IndentationLeft = 30f;
                            //                document.Add(para);
                            //            }
                            //            else
                            //            {
                            //                //for (int j = 0; j < para.ToArray().Length; j++)
                            //                //{
                            //                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                            //                //}
                            //                document.Add(para);
                            //            }
                            //        }


                            //    }
                            //    //document.Close();
                            //}

                            //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));
                            if (m != rowsize - 1)
                            {
                                Chunk NextLine = new Chunk("\n");
                                Paragraph pNextLine = new Paragraph();
                                pNextLine.Add(NextLine);
                                pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                document.Add(pNextLine);
                            }
                        }
                    }
                    //document.Add(list);

                    double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                    if (remainingPageSpace < 197.00)
                        document.NewPage();

                    Chunk lochunkSingleFundNONSTD2 = new Chunk("If you have questions, please call " + Convert.ToString(table.Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                    Chunk lophone1 = new Chunk(" (312)960-0231 ", setFontsAll(9, 1, 0));
                    Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                    Chunk lophone2 = new Chunk("(312)960-0211.", setFontsAll(9, 1, 0));
                    Paragraph plochunkSingleFundNONSTD2 = new Paragraph();
                    plochunkSingleFundNONSTD2.Add(lochunkSingleFundNONSTD2);
                    //plochunkSingleFundNONSTD2.Add(lophone1);
                    // plochunkSingleFundNONSTD2.Add(lodesc1);
                    //plochunkSingleFundNONSTD2.Add(lophone2);
                    plochunkSingleFundNONSTD2.Leading = 11f;
                    plochunkSingleFundNONSTD2.SpacingBefore = 8f;
                    plochunkSingleFundNONSTD2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(plochunkSingleFundNONSTD2);


                    Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
                    plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
                    plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    plochunkSingleFundNONSTD3.SpacingBefore = 5f;
                    document.Add(plochunkSingleFundNONSTD3);


                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(table.Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                    if (File.Exists(Imagepath))
                    {
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);
                    }
                    else if (!File.Exists(Imagepath))
                    {
                        Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);
                    }

                    Chunk lochunkSingleFundNONSTD4 = new Chunk(Convert.ToString(table.Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                    Paragraph plochunkSingleFundNONSTD4 = new Paragraph();
                    plochunkSingleFundNONSTD4.Add(lochunkSingleFundNONSTD4);
                    plochunkSingleFundNONSTD4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(plochunkSingleFundNONSTD4);





                    #endregion
                }
                else if (rowsize == 1)
                {
                    #region Single Distribution Letter Fund

                    string InstructionsStart = "";
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                    {
                        InstructionsStart = "Dear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + ":";
                    }
                    Chunk lochunkSingleStart = new Chunk(InstructionsStart, setFontsAll(9, 0, 0));
                    Paragraph pSingleStart = new Paragraph();
                    pSingleStart.Add(lochunkSingleStart);
                    pSingleStart.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pSingleStart.SpacingBefore = 5f;
                    document.Add(pSingleStart);

                    Chunk lochunkSingleStart1;
                    Paragraph pSingleStart1 = new Paragraph();
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_distributionpaymentmethod"]).ToUpper() == "Pays to GGES".ToUpper())
                        lochunkSingleStart1 = new Chunk("We will make an additional investment in GP Diversified Growth Strategies on " + effDistributionDate + ", representing your share of distribution ", setFontsAll(9, 0, 0));
                    else
                    {
                        //lochunkSingleStart1 = new Chunk("We have wired a cash distribution to your account today, representing your share of distribution ", setFontsAll(9, 0, 0));
                        lochunkSingleStart1 = new Chunk("A cash distribution will be wired to your account the week of " + WireDt + "", setFontsAll(9, 0, 0));
                        pSingleStart1.Add(lochunkSingleStart1);
                        Chunk superScript = new Chunk(DateSuperScript, setFontsAll(6, 0, 0));
                        superScript.SetTextRise(5f);
                        pSingleStart1.Add(superScript);
                        lochunkSingleStart1 = new Chunk(", representing your share of distribution ", setFontsAll(9, 0, 0));
                    }


                    //pSingleStart1.Add(lochunkSingleStart1);
                    //pSingleStart1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //pSingleStart1.SpacingBefore = 9f;
                    //document.Add(pSingleStart1);

                    string InstructionsStart1 = "";
                    string InstructionsStart2 = "";
                    string InstructionsStart3 = "";
                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) != "")
                    {
                        InstructionsStart1 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]);
                        InstructionsStart1 = InstructionsStart1.Replace("(", "(''").Replace(")", "'')");
                    }

                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_curdistp_db"]) != "")
                    {
                        InstructionsStart2 = Percentage(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_curdistp_db"]));
                    }

                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]) != "")
                    {
                        InstructionsStart3 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundname"]);
                        string[] str1 = InstructionsStart3.Split('(');
                        InstructionsStart3 = str1[1].Replace(")", "");
                    }

                    Chunk lochunkSingleStart2 = new Chunk("from " + InstructionsStart1 + ". This distribution equals " + InstructionsStart2 + " of your ", setFontsAll(9, 0, 0));
                    //Paragraph pSingleStart2 = new Paragraph();
                    //pSingleStart2.Add(lochunkSingleStart2);
                    //pSingleStart2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //document.Add(pSingleStart2);

                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundnickname"]) != "")
                    {
                        InstructionsStart1 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fundnickname"]);
                    }


                    Chunk lochunkSingleStart3 = new Chunk("commitment to " + InstructionsStart3 + ". The enclosed statement provides you with the detailed information regarding your distribution amount, commitment, and funding status.", setFontsAll(9, 0, 0));
                    //Paragraph pSingleStart3 = new Paragraph();
                    //pSingleStart3.Add(lochunkSingleStart3);
                    //pSingleStart3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //document.Add(pSingleStart3);

                    pSingleStart1.Add(lochunkSingleStart1);
                    pSingleStart1.Add(lochunkSingleStart2);
                    pSingleStart1.Add(lochunkSingleStart3);
                    pSingleStart1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    pSingleStart1.SpacingBefore = 9f;
                    pSingleStart1.Leading = 11f;
                    pSingleStart1.SpacingAfter = 12f;
                    document.Add(pSingleStart1);

                    //Chunk lochunkSingleStart4 = new Chunk("your distribution amount, commitment, and funding status.", setFontsAll(9, 0, 0));
                    //Paragraph pSingleStart4 = new Paragraph();
                    //pSingleStart4.Add(lochunkSingleStart4);
                    //pSingleStart4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //document.Add(pSingleStart4);


                    iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                    iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                    hw.Style = styles;
                    //styles.LoadTagStyle("ol", "leading", "16,0");
                    //styles.LoadTagStyle("li", "face", "garamond");
                    //styles.LoadTagStyle("span", "size", "9pt");
                    //styles.LoadTagStyle("body", "font-family", "verdana");
                    //styles.LoadTagStyle("body", "font-size", "12pt");

                    ArrayList objects = null;

                    string FundSpecificDesc = "";
                    //List list = new List(List.UNORDERED, 10f);
                    //list.SetListSymbol("\u2022");
                    //list.IndentationLeft = 30f;
                    for (int m = 0; m < rowsize; m++)
                    {
                        if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                        {

                            FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);
                            FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                            //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                            //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                            FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                            FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                            //Chunk NextLine1 = new Chunk("\n");
                            //Paragraph pNextLine1 = new Paragraph();
                            //pNextLine1.Add(NextLine1);
                            //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(pNextLine1);


                            MemoryStream output = new MemoryStream();
                            StreamWriter html = new StreamWriter(output, Encoding.UTF8);


                            html.Write(string.Concat(DOCUMENT_HTML_START, FundSpecificDesc, DOCUMENT_HTML_END));
                            html.Close();
                            html.Dispose();

                            MemoryStream generate = new MemoryStream(output.ToArray());
                            StreamReader stringReader = new StreamReader(generate);
                            foreach (object item in iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, styles))
                            {
                                document.Add((IElement)item);
                            }

                            //cleanup these streams
                            html.Dispose();
                            stringReader.Dispose();
                            output.Dispose();
                            generate.Dispose();

                            //using (StringReader stringReader = new StringReader(FundSpecificDesc))
                            //{
                            //    //List<IElement> parsedList = new List<IElement>();

                            //    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                            //    //document.Open();
                            //    foreach (object item in objects)
                            //    {
                            //        if (item is List)
                            //        {
                            //            List list = item as List;
                            //            list.Autoindent = false;
                            //            list.IndentationLeft = 25f;
                            //            list.SetListSymbol("\u2022                      ");
                            //            list.SymbolIndent = 25f;
                            //            document.Add((IElement)item);
                            //        }


                            //        if (item is Paragraph)
                            //        {
                            //            Paragraph para = item as Paragraph; //setFontsverdana

                            //            if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                            //            {
                            //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                            //                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                            //                para.IndentationLeft = 30f;
                            //                document.Add(para);
                            //            }
                            //            else
                            //            {
                            //                //for (int j = 0; j < para.ToArray().Length; j++)
                            //                //{
                            //                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                            //                //}
                            //                document.Add(para);
                            //            }
                            //        }


                            //    }
                            //    //document.Close();
                            //}

                            //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                            //Chunk NextLine = new Chunk("\n");
                            //Paragraph pNextLine = new Paragraph();
                            //pNextLine.Add(NextLine);
                            //pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            // document.Add(pNextLine);
                        }
                    }
                    //document.Add(list);

                    double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                    if (remainingPageSpace < 197.00)
                        document.NewPage();

                    Chunk lochunkSingleFundNONSTD2 = new Chunk("If you have questions, please call " + Convert.ToString(table.Rows[0]["AdvisorName"]) + " at (312) 960-0200.", setFontsAll(9, 0, 0));
                    Chunk lophone1 = new Chunk(" (312)960-0231 ", setFontsAll(9, 1, 0));
                    Chunk lodesc1 = new Chunk("or Ben Beavers ", setFontsAll(9, 0, 0));
                    Chunk lophone2 = new Chunk("(312)960-029.", setFontsAll(9, 1, 0));
                    Paragraph plochunkSingleFundNONSTD2 = new Paragraph();
                    plochunkSingleFundNONSTD2.Add(lochunkSingleFundNONSTD2);
                    // plochunkSingleFundNONSTD2.Add(lophone1);
                    // plochunkSingleFundNONSTD2.Add(lodesc1);
                    //  plochunkSingleFundNONSTD2.Add(lophone2);
                    plochunkSingleFundNONSTD2.SpacingBefore = 8f;
                    plochunkSingleFundNONSTD2.Leading = 11f;
                    plochunkSingleFundNONSTD2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(plochunkSingleFundNONSTD2);


                    Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
                    plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
                    plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    plochunkSingleFundNONSTD3.SpacingBefore = 5f;
                    document.Add(plochunkSingleFundNONSTD3);


                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + Convert.ToString(table.Rows[0]["AdvisorName"]) + ".jpg";// "Ted Neild.jpg";
                    if (File.Exists(Imagepath))
                    {
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);
                    }
                    else if (!File.Exists(Imagepath))
                    {
                        Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);
                    }


                    Chunk lochunkSingleFundNONSTD4 = new Chunk(Convert.ToString(table.Rows[0]["AdvisorName"]), setFontsAll(9, 0, 0));
                    Paragraph plochunkSingleFundNONSTD4 = new Paragraph();
                    plochunkSingleFundNONSTD4.Add(lochunkSingleFundNONSTD4);
                    plochunkSingleFundNONSTD4.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    document.Add(plochunkSingleFundNONSTD4);


                    #endregion
                }

            }
        }
        catch
        { }





        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");


    }
    #endregion

    #region Gresham Advisors General Letter Regarding LegalEntity Custom

    public string GetGreshamAdvisorsGLRLegalEntity()
    {
        string[] SignatureText = new string[10000];
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.GreshamAdvisorsGLRLegalEntity); //"SP_S_GreshamAdvisorsGLRLegalEntity 13, '18511645-77CA-E111-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE11-8391-001D09665E8F', 'F7063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "GreshamAdvisorsGLRLegalEntity.pdf";
        PdfWriter pdfwriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        AddFooter(document);


        document.Open();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + ".xls";


        try
        {
            if (DSCount > 0)
            {

                //for (int i = 0; i < DSCount; i++)
                //{

                //if (i != 0)
                //{
                //    document.NewPage();
                //}

                iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, newdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                iTextSharp.text.Cell loCell = new Cell();
                setTableProperty(loTable, ReportType.GreshamAdvisorsGLRLegalEntity);
                iTextSharp.text.Chunk lochunk = new Chunk();

                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\GreshamAdvisors_logo.tif");
                png.SetAbsolutePosition(48, 800);//540
                //png.ScaleToFit(288f, 42f);
                png.ScalePercent(22);
                document.Add(png);

                string AsOfDate = "\n\n";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]) != "")
                {
                    AsOfDate = AsOfDate + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]);
                }
                Chunk lochunkAsOfDate = new Chunk(AsOfDate, setFontsAll(8, 0, 0));
                Paragraph pAsOfDate = new Paragraph();
                pAsOfDate.Add(lochunkAsOfDate);
                pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                document.Add(pAsOfDate);


                string FullName = "\n\n";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]) != "")
                {
                    FullName = FullName + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]);
                }
                Chunk lochunkFullName = new Chunk(FullName, setFontsAll(9, 0, 0));
                Paragraph pFullName = new Paragraph();
                pFullName.Add(lochunkFullName);
                pFullName.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pFullName.Leading = 11f;
                document.Add(pFullName);

                string AddressLine1 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]) != "")
                {
                    AddressLine1 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]);
                }
                Chunk lochunkAddressLine1 = new Chunk(AddressLine1, setFontsAll(9, 0, 0));
                Paragraph pAddressLine1 = new Paragraph();
                pAddressLine1.Add(lochunkAddressLine1);
                pAddressLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine1.Leading = 11f;
                document.Add(pAddressLine1);

                string AddressLine2 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]) != "")
                {
                    AddressLine2 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]);
                }
                Chunk lochunkAddressLine2 = new Chunk(AddressLine2, setFontsAll(9, 0, 0));
                Paragraph pAddressLine2 = new Paragraph();
                pAddressLine2.Add(lochunkAddressLine2);
                pAddressLine2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine2.Leading = 11f;
                document.Add(pAddressLine2);

                string AddressLine3 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]) != "")
                {
                    AddressLine3 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]);
                }
                Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(9, 0, 0));
                Paragraph pAddressLine3 = new Paragraph();
                pAddressLine3.Add(lochunkAddressLine3);
                pAddressLine3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine3.Leading = 11f;
                document.Add(pAddressLine3);

                string AddressDetails = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]) != "")
                {
                    AddressDetails = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]);
                }
                Chunk lochunkAddressDetails = new Chunk(AddressDetails, setFontsAll(9, 0, 0));
                Paragraph pAddressDetails = new Paragraph();
                pAddressDetails.Add(lochunkAddressDetails);
                pAddressDetails.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressDetails.Leading = 11f;
                document.Add(pAddressDetails);


                string FullNameBold = "RE: ";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]) != "")
                {
                    FullNameBold = FullNameBold + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]);
                }
                Chunk lochunkFullName1 = new Chunk(FullNameBold, setFontsAll(9, 1, 0));
                Paragraph pFullName1 = new Paragraph();
                pFullName1.Add(lochunkFullName1);
                pFullName1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pFullName1.SpacingBefore = 12f;
                document.Add(pFullName1);


                string InstructionsStart = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                {
                    InstructionsStart = "\nDear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + " :";
                }
                Chunk lochunkSingleStart = new Chunk(InstructionsStart, setFontsAll(9, 0, 0));
                Paragraph pSingleStart = new Paragraph();
                pSingleStart.Add(lochunkSingleStart);
                pSingleStart.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pSingleStart.SpacingAfter = 8f;
                document.Add(pSingleStart);

                Chunk NextLine1 = new Chunk("\n");
                Paragraph pNextLine1 = new Paragraph();
                pNextLine1.Add(NextLine1);
                pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pNextLine1.Leading = 7f;
                document.Add(pNextLine1);

                int rowsize = newdataset.Tables[0].Rows.Count;
                int colsize = 2;

                if (rowsize > 0)
                {

                    iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                    iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                    hw.Style = styles;
                    //styles.LoadTagStyle("ol", "leading", "16,0");
                    //styles.LoadTagStyle("li", "face", "garamond");
                    //styles.LoadTagStyle("span", "size", "11pt");
                    //styles.LoadTagStyle("body", "font-family", "verdana");
                    //styles.LoadTagStyle("body", "font-size", "11pt");

                    ArrayList objects = null;

                    string FundSpecificDesc = "";
                    //List list = new List(List.UNORDERED, 10f);
                    //list.SetListSymbol("\u2022");
                    //list.IndentationLeft = 30f;
                    //for (int m = 0; m < rowsize; m++)
                    //{
                    if (Convert.ToString(table.Rows[0]["ssi_lettertext"]) != "")
                    {

                        FundSpecificDesc = Convert.ToString(table.Rows[0]["ssi_lettertext"]);
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                        //Chunk NextLine1 = new Chunk("\n");
                        //Paragraph pNextLine1 = new Paragraph();
                        //pNextLine1.Add(NextLine1);
                        //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        //document.Add(pNextLine1);
                        using (StringReader stringReader = new StringReader(FundSpecificDesc))
                        {
                            //List<IElement> parsedList = new List<IElement>();

                            objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                            //document.Open();
                            foreach (object item in objects)
                            {
                                if (item is List)
                                {
                                    List list = item as List;
                                    list.Autoindent = false;
                                    list.IndentationLeft = 25f;
                                    list.SetListSymbol("\u2022                      ");
                                    list.SymbolIndent = 25f;
                                    document.Add((IElement)item);
                                }


                                if (item is Paragraph)
                                {
                                    Paragraph para = item as Paragraph; //setFontsverdana

                                    if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                    {
                                        ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                        ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                        para.IndentationLeft = 30f;
                                        document.Add(para);
                                    }
                                    else
                                    {
                                        //for (int j = 0; j < para.ToArray().Length; j++)
                                        //{
                                        //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                        //}
                                        document.Add(para);
                                    }
                                }


                            }
                            //document.Close();
                        }

                        //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                        Chunk NextLine = new Chunk("\n");
                        Paragraph pNextLine = new Paragraph();
                        pNextLine.Add(NextLine);
                        pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        pNextLine.Leading = 11f;
                        //document.Add(pNextLine);
                    }
                    //}

                    for (int m = 0; m < rowsize; m++)
                    {
                        if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                        {
                            if (Convert.ToBoolean(table.Rows[m]["ssi_fundspecificFlg"]) == true)
                            {
                                FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);
                                FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                                //Chunk NextLine1 = new Chunk("\n");
                                //Paragraph pNextLine1 = new Paragraph();
                                //pNextLine1.Add(NextLine1);
                                //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                //document.Add(pNextLine1);
                                using (StringReader stringReader = new StringReader(FundSpecificDesc))
                                {
                                    //List<IElement> parsedList = new List<IElement>();

                                    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                                    //document.Open();
                                    foreach (object item in objects)
                                    {
                                        if (item is List)
                                        {
                                            List list = item as List;
                                            list.Autoindent = false;
                                            list.IndentationLeft = 25f;
                                            list.SetListSymbol("\u2022                      ");
                                            list.SymbolIndent = 25f;
                                            document.Add((IElement)item);
                                        }


                                        if (item is Paragraph)
                                        {
                                            Paragraph para = item as Paragraph; //setFontsverdana

                                            if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                            {
                                                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                                para.IndentationLeft = 30f;
                                                document.Add(para);
                                            }
                                            else
                                            {
                                                //for (int j = 0; j < para.ToArray().Length; j++)
                                                //{
                                                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                                //}
                                                document.Add(para);
                                            }
                                        }


                                    }
                                    //document.Close();
                                }

                                //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                                if (m != rowsize - 1)
                                {
                                    Chunk NextLine = new Chunk("\n");
                                    Paragraph pNextLine = new Paragraph();
                                    pNextLine.Add(NextLine);
                                    pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    document.Add(pNextLine);
                                }
                            }
                        }
                    }

                    double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                    if (remainingPageSpace < 132.00)
                        document.NewPage();

                    Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
                    plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
                    plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    plochunkSingleFundNONSTD3.SpacingBefore = 8f;
                    document.Add(plochunkSingleFundNONSTD3);


                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]) != "")
                    {
                        SignatureText = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]).Split(',');
                    }


                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]) != "")
                    {
                        DynamicFlg = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]);
                    }


                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]) != "")
                    {
                        Orientation = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]);
                    }

                }
                //}

                #region Signature

                if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "1")
                {
                    iTextSharp.text.Cell loEmptyChunky = new Cell();
                    #region Row1

                    #region Signature Image Row1

                    iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable1.Border = 0;
                    loTable1.Width = 100f;

                    loTable1.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable1.Cellpadding = 0;
                    loTable1.Cellspacing = 0;
                    iTextSharp.text.Cell loCell1 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell1 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k < 4)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell1.Width = 25;
                                loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell1.AddElement(SignatureJpg);
                                loCell1.Border = 0;
                                loTable1.AddCell(loCell1);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loCell1.Border = 0;
                                loTable1.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n", setFontsAll(11, 0, 0));
                                //loCell1.Width = 25;
                                //loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell1.AddElement(loEmptyChunky);
                                //loCell1.Border = 0;
                                //loTable1.AddCell(loCell1);
                            }
                        }
                    }
                    document.Add(loTable1);
                    #endregion

                    #region Signature Text Row1

                    iTextSharp.text.Table loTable11 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable11.Border = 0;
                    loTable11.Width = 100f;
                    loTable11.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable11.Cellpadding = 0;
                    loTable11.Cellspacing = 0;
                    iTextSharp.text.Cell loCell11 = new Cell();
                    iTextSharp.text.Chunk lochunk11 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell11 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k < 4)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk11 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell11.Width = 25;
                                loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell11.AddElement(lochunk11);
                                loCell11.Border = 0;
                                loTable11.AddCell(loCell11);
                            }
                            else
                            {
                                lochunk11 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell11.Width = 25;
                                loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                loCell11.AddElement(lochunk11);
                                loCell11.Border = 0;
                                loTable11.AddCell(loCell11);
                            }
                        }

                    }
                    document.Add(loTable11);

                    #endregion

                    #endregion

                    #region Row2

                    #region Signature Image Row2

                    iTextSharp.text.Table loTable2 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable2.Border = 0;
                    loTable2.Width = 100f;
                    loTable2.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable2.Cellpadding = 0;
                    loTable2.Cellspacing = 0;
                    iTextSharp.text.Cell loCell2 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell2 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 3 && k < 8)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell2.Width = 25;
                                loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell2.AddElement(SignatureJpg);
                                loCell2.Border = 0;
                                loTable2.AddCell(loCell2);
                            }
                            else //if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable2.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell2.Width = 25;
                                //loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell2.AddElement(loEmptyChunky);
                                //loCell2.Border = 0;
                                //loTable2.AddCell(loCell2);
                            }
                        }
                    }
                    document.Add(loTable2);
                    #endregion

                    #region Signature Text Row2

                    iTextSharp.text.Table loTable22 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable22.Border = 0;
                    loTable22.Width = 100f;
                    loTable22.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable22.Cellpadding = 0;
                    loTable22.Cellspacing = 0;
                    iTextSharp.text.Cell loCell22 = new Cell();
                    iTextSharp.text.Chunk lochunk22 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell22 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 3 && k < 8)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk22 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell22.Width = 25;
                                loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell22.AddElement(lochunk22);
                                loCell22.Border = 0;
                                loTable22.AddCell(loCell22);
                            }
                            else
                            {
                                lochunk22 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell22.Width = 25;
                                loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell22.AddElement(lochunk22);
                                loCell22.Border = 0;
                                loTable22.AddCell(loCell22);
                            }
                        }

                    }
                    document.Add(loTable22);

                    #endregion

                    #endregion

                    #region Row3

                    #region Signature Image Row3

                    iTextSharp.text.Table loTable3 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable3.Border = 0;
                    loTable3.Width = 100f;
                    loTable3.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable3.Cellpadding = 0;
                    loTable3.Cellspacing = 0;
                    iTextSharp.text.Cell loCell3 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell3 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 7 && k < 12)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell3.Width = 25;
                                loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell3.AddElement(SignatureJpg);
                                loCell3.Border = 0;
                                loTable3.AddCell(loCell3);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable3.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell3.Width = 25;
                                //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell3.AddElement(loEmptyChunky);
                                //loCell3.Border = 0;
                                //loTable3.AddCell(loCell3);
                            }
                        }
                    }
                    document.Add(loTable3);
                    #endregion

                    #region Signature Text Row3

                    iTextSharp.text.Table loTable33 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable33.Border = 0;
                    loTable33.Width = 100f;
                    loTable33.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable33.Cellpadding = 0;
                    loTable33.Cellspacing = 0;
                    iTextSharp.text.Cell loCell33 = new Cell();
                    iTextSharp.text.Chunk lochunk33 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell33 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 7 && k < 12)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk33 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell33.Width = 25;
                                loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell33.AddElement(lochunk33);
                                loCell33.Border = 0;
                                loTable33.AddCell(loCell33);
                            }
                            else
                            {
                                lochunk33 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell33.Width = 25;
                                loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell33.AddElement(lochunk33);
                                loCell33.Border = 0;
                                loTable33.AddCell(loCell33);
                            }
                        }

                    }
                    document.Add(loTable33);

                    #endregion

                    #endregion

                    #region Row4

                    #region Signature Image Row4

                    iTextSharp.text.Table loTable4 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable4.Border = 0;
                    loTable4.Width = 100f;
                    loTable4.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable4.Cellpadding = 0;
                    loTable4.Cellspacing = 0;
                    iTextSharp.text.Cell loCell4 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell4 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 11 && k < 15)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell4.Width = 25;
                                loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell4.AddElement(SignatureJpg);
                                loCell4.Border = 0;
                                loTable4.AddCell(loCell4);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable4.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell4.Width = 25;
                                //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell4.AddElement(loEmptyChunky);
                                //loCell4.Border = 0;
                                //loTable4.AddCell(loCell4);
                            }
                        }
                    }
                    document.Add(loTable4);
                    #endregion

                    #region Signature Text Row4

                    iTextSharp.text.Table loTable44 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable44.Border = 0;
                    loTable44.Width = 100f;
                    loTable44.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable44.Cellpadding = 0;
                    loTable44.Cellspacing = 0;
                    iTextSharp.text.Cell loCell44 = new Cell();
                    iTextSharp.text.Chunk lochunk44 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell44 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 11 && k < 15)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk44 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell44.Width = 25;
                                loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell44.AddElement(lochunk44);
                                loCell44.Border = 0;
                                loTable44.AddCell(loCell44);
                            }
                            else
                            {
                                lochunk44 = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                                loCell44.Width = 25;
                                loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell44.AddElement(lochunk44);
                                loCell44.Border = 0;
                                loTable44.AddCell(loCell44);
                            }
                        }

                    }
                    document.Add(loTable44);

                    #endregion

                    #endregion

                }
                else if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "2")
                {
                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();
                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        pimage = new Paragraph();
                        PSignature = new Paragraph();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);

                            Chunk lochunkSign = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                            PSignature.Add(lochunkSign);
                            PSignature.SetAlignment("left");
                            document.Add(PSignature);
                        }
                        else
                        {
                            Chunk lochunkSign = new Chunk("\n\n\n\n\n" + SignatureText[k], setFontsAll(9, 0, 0));
                            PSignature.Add(lochunkSign);
                            PSignature.SetAlignment("left");
                            document.Add(PSignature);
                        }
                    }
                }
                else
                {
                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();
                    string SignImage = "Ted Neild";

                    if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_ownerfirstname_hh_mail"]) != "" && Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_ownerlname_hh_mail"]) != "")
                    {
                        SignImage = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_ownerfirstname_hh_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_ownerlname_hh_mail"]);
                    }

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignImage + ".jpg";
                    if (File.Exists(Imagepath))
                    {
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);

                        Chunk lochunkSign = new Chunk(SignImage, setFontsAll(9, 0, 0));
                        PSignature.Add(lochunkSign);
                        PSignature.SetAlignment("left");
                        document.Add(PSignature);

                    }
                }
                #endregion

            }
        }
        catch
        { }





        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");

    }

    #endregion

    #region Gresham Advisors General Letter Non-Specific Recepient

    public string GetGreshamAdvisorsGLNSRecipient()
    {
        string[] SignatureText = new string[10000];
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.GreshamAdvisorsNonSpecific);// "SP_S_GreshamAdvisorsGLNSRecipient 13, '18511645-77CA-E111-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE11-8391-001D09665E8F', 'F7063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();


        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "GreshamAdvisorsGLNSRecipient.pdf";
        PdfWriter pdfwriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        AddFooter(document);

        document.Open();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        if (DSCount > 0)
        {

            //for (int i = 0; i < DSCount; i++)
            //{

            //if (i != 0)
            //{
            //    document.NewPage();
            //}



            iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, newdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
            iTextSharp.text.Cell loCell = new Cell();
            setTableProperty(loTable, ReportType.GreshamAdvisorsNonSpecific);
            iTextSharp.text.Chunk lochunk = new Chunk();

            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\GreshamAdvisors_logo.tif");
            png.SetAbsolutePosition(48, 800);//540
            //png.ScaleToFit(288f, 42f);
            png.ScalePercent(22);
            document.Add(png);

            string AsOfDate = "\n\n";
            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]) != "")
            {
                AsOfDate = AsOfDate + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]);
            }
            Chunk lochunkAsOfDate = new Chunk(AsOfDate + "\n\n", setFontsAll(8, 1, 0));
            Paragraph pAsOfDate = new Paragraph();
            pAsOfDate.Add(lochunkAsOfDate);
            pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
            document.Add(pAsOfDate);


            int rowsize = newdataset.Tables[0].Rows.Count;
            int colsize = 2;

            if (rowsize > 0)
            {

                iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                hw.Style = styles;
                //styles.LoadTagStyle("ol", "leading", "16,0");
                //styles.LoadTagStyle("li", "face", "garamond");
                //styles.LoadTagStyle("span", "size", "11pt");
                //styles.LoadTagStyle("body", "font-family", "verdana");
                //styles.LoadTagStyle("body", "font-size", "11pt");

                ArrayList objects = null;

                string FundSpecificDesc = "";
                //List list = new List(List.UNORDERED, 10f);
                //list.SetListSymbol("\u2022");
                //list.IndentationLeft = 30f;
                //for (int m = 0; m < rowsize; m++)
                //{
                if (Convert.ToString(table.Rows[0]["ssi_lettertext"]) != "")
                {

                    FundSpecificDesc = Convert.ToString(table.Rows[0]["ssi_lettertext"]);
                    FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                    //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                    //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                    FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                    FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                    //Chunk NextLine1 = new Chunk("\n");
                    //Paragraph pNextLine1 = new Paragraph();
                    //pNextLine1.Add(NextLine1);
                    //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //document.Add(pNextLine1);

                    using (StringReader stringReader = new StringReader(FundSpecificDesc))
                    {
                        //List<IElement> parsedList = new List<IElement>();

                        objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                        //document.Open();
                        foreach (object item in objects)
                        {
                            if (item is List)
                            {
                                List list = item as List;
                                list.Autoindent = false;
                                list.IndentationLeft = 25f;
                                list.SetListSymbol("\u2022                      ");
                                list.SymbolIndent = 25f;
                                document.Add((IElement)item);
                            }


                            if (item is Paragraph)
                            {
                                Paragraph para = item as Paragraph; //setFontsverdana

                                if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                {
                                    ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                    ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                    para.IndentationLeft = 30f;
                                    document.Add(para);
                                }
                                else
                                {
                                    //for (int j = 0; j < para.ToArray().Length; j++)
                                    //{
                                    //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                    //}
                                    document.Add(para);
                                }
                            }


                        }
                        //document.Close();
                    }
                    //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                    //Chunk NextLine = new Chunk("\n");
                    //Paragraph pNextLine = new Paragraph();
                    //pNextLine.Add(NextLine);
                    //pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    //document.Add(pNextLine);
                }
                //}
                //document.Add(list);

                for (int m = 0; m < rowsize; m++)
                {
                    if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                    {
                        if (Convert.ToBoolean(table.Rows[m]["ssi_fundspecificFlg"]) == true)
                        {

                            FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);
                            FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                            //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                            //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                            FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                            FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                            //Chunk NextLine1 = new Chunk("\n");
                            //Paragraph pNextLine1 = new Paragraph();
                            //pNextLine1.Add(NextLine1);
                            //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                            //document.Add(pNextLine1);

                            using (StringReader stringReader = new StringReader(FundSpecificDesc))
                            {
                                //List<IElement> parsedList = new List<IElement>();

                                objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                                //document.Open();
                                foreach (object item in objects)
                                {
                                    if (item is List)
                                    {
                                        List list = item as List;
                                        list.Autoindent = false;
                                        list.IndentationLeft = 25f;
                                        list.SetListSymbol("\u2022                      ");
                                        list.SymbolIndent = 25f;
                                        document.Add((IElement)item);
                                    }


                                    if (item is Paragraph)
                                    {
                                        Paragraph para = item as Paragraph; //setFontsverdana

                                        if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                        {
                                            ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                            ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                            para.IndentationLeft = 30f;
                                            document.Add(para);
                                        }
                                        else
                                        {
                                            //for (int j = 0; j < para.ToArray().Length; j++)
                                            //{
                                            //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                            //}
                                            document.Add(para);
                                        }
                                    }


                                }
                                //document.Close();
                            }
                            //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                            if (m != rowsize - 1)
                            {
                                Chunk NextLine = new Chunk("\n");
                                Paragraph pNextLine = new Paragraph();
                                pNextLine.Add(NextLine);
                                pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                document.Add(pNextLine);
                            }
                        }
                    }
                }

            }

            double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

            if (remainingPageSpace < 150.00)
                document.NewPage();

            Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
            Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
            plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
            plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
            plochunkSingleFundNONSTD3.SpacingBefore = 8f;
            document.Add(plochunkSingleFundNONSTD3);


            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]) != "")
            {
                SignatureText = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]).Split(',');
            }


            if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]) != "")
            {
                DynamicFlg = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]);
            }


            if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]) != "")
            {
                Orientation = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]);
            }
            //}

            #region Signature

            if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "1")
            {
                iTextSharp.text.Cell loEmptyChunky = new Cell();
                #region Row1

                #region Signature Image Row1

                iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable1.Border = 0;
                loTable1.Width = 100f;

                loTable1.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable1.Cellpadding = 0;
                loTable1.Cellspacing = 0;
                iTextSharp.text.Cell loCell1 = new Cell();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell1 = new iTextSharp.text.Cell();
                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k < 4)
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loCell1.Width = 25;
                            loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell1.AddElement(SignatureJpg);
                            loCell1.Border = 0;
                            loTable1.AddCell(loCell1);
                        }
                        else if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loEmptyChunky = new Cell();
                            loEmptyChunky.Width = 25;
                            loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loEmptyChunky.AddElement(SignatureJpg);
                            loCell1.Border = 0;
                            loTable1.AddCell(loEmptyChunky);
                            //Chunk loEmptyChunky = new Chunk("\n\n", setFontsAll(11, 0, 0));
                            //loCell1.Width = 25;
                            //loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            //loCell1.AddElement(loEmptyChunky);
                            //loCell1.Border = 0;
                            //loTable1.AddCell(loCell1);
                        }
                    }
                }
                document.Add(loTable1);
                #endregion

                #region Signature Text Row1

                iTextSharp.text.Table loTable11 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable11.Border = 0;
                loTable11.Width = 100f;
                loTable11.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable11.Cellpadding = 0;
                loTable11.Cellspacing = 0;
                iTextSharp.text.Cell loCell11 = new Cell();
                iTextSharp.text.Chunk lochunk11 = new Chunk();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell11 = new iTextSharp.text.Cell();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k < 4)
                    {
                        if (File.Exists(Imagepath))
                        {
                            lochunk11 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell11.Width = 25;
                            loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell11.AddElement(lochunk11);
                            loCell11.Border = 0;
                            loTable11.AddCell(loCell11);
                        }
                        else
                        {
                            lochunk11 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell11.Width = 25;
                            loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell11.AddElement(lochunk11);
                            loCell11.Border = 0;
                            loTable11.AddCell(loCell11);
                        }
                    }

                }
                document.Add(loTable11);

                #endregion

                #endregion

                #region Row2

                #region Signature Image Row2

                iTextSharp.text.Table loTable2 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable2.Border = 0;
                loTable2.Width = 100f;
                loTable2.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable2.Cellpadding = 0;
                loTable2.Cellspacing = 0;
                iTextSharp.text.Cell loCell2 = new Cell();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell2 = new iTextSharp.text.Cell();
                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k > 3 && k < 8)
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loCell2.Width = 25;
                            loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell2.AddElement(SignatureJpg);
                            loCell2.Border = 0;
                            loTable2.AddCell(loCell2);
                        }
                        else //if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loEmptyChunky = new Cell();
                            loEmptyChunky.Width = 25;
                            loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loEmptyChunky.AddElement(SignatureJpg);
                            loEmptyChunky.Border = 0;
                            loTable2.AddCell(loEmptyChunky);
                            //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                            //loCell2.Width = 25;
                            //loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            //loCell2.AddElement(loEmptyChunky);
                            //loCell2.Border = 0;
                            //loTable2.AddCell(loCell2);
                        }
                    }
                }
                document.Add(loTable2);
                #endregion

                #region Signature Text Row2

                iTextSharp.text.Table loTable22 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable22.Border = 0;
                loTable22.Width = 100f;
                loTable22.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable22.Cellpadding = 0;
                loTable22.Cellspacing = 0;
                iTextSharp.text.Cell loCell22 = new Cell();
                iTextSharp.text.Chunk lochunk22 = new Chunk();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell22 = new iTextSharp.text.Cell();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k > 3 && k < 8)
                    {
                        if (File.Exists(Imagepath))
                        {
                            lochunk22 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell22.Width = 25;
                            loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell22.AddElement(lochunk22);
                            loCell22.Border = 0;
                            loTable22.AddCell(loCell22);
                        }
                        else
                        {
                            lochunk22 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell22.Width = 25;
                            loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell22.AddElement(lochunk22);
                            loCell22.Border = 0;
                            loTable22.AddCell(loCell22);
                        }
                    }

                }
                document.Add(loTable22);

                #endregion

                #endregion

                #region Row3

                #region Signature Image Row3

                iTextSharp.text.Table loTable3 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable3.Border = 0;
                loTable3.Width = 100f;
                loTable3.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable3.Cellpadding = 0;
                loTable3.Cellspacing = 0;
                iTextSharp.text.Cell loCell3 = new Cell();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell3 = new iTextSharp.text.Cell();
                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k > 7 && k < 12)
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loCell3.Width = 25;
                            loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell3.AddElement(SignatureJpg);
                            loCell3.Border = 0;
                            loTable3.AddCell(loCell3);
                        }
                        else if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loEmptyChunky = new Cell();
                            loEmptyChunky.Width = 25;
                            loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loEmptyChunky.AddElement(SignatureJpg);
                            loEmptyChunky.Border = 0;
                            loTable3.AddCell(loEmptyChunky);
                            //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                            //loCell3.Width = 25;
                            //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            //loCell3.AddElement(loEmptyChunky);
                            //loCell3.Border = 0;
                            //loTable3.AddCell(loCell3);
                        }
                    }
                }
                document.Add(loTable3);
                #endregion

                #region Signature Text Row3

                iTextSharp.text.Table loTable33 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable33.Border = 0;
                loTable33.Width = 100f;
                loTable33.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable33.Cellpadding = 0;
                loTable33.Cellspacing = 0;
                iTextSharp.text.Cell loCell33 = new Cell();
                iTextSharp.text.Chunk lochunk33 = new Chunk();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell33 = new iTextSharp.text.Cell();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k > 7 && k < 12)
                    {
                        if (File.Exists(Imagepath))
                        {
                            lochunk33 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell33.Width = 25;
                            loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell33.AddElement(lochunk33);
                            loCell33.Border = 0;
                            loTable33.AddCell(loCell33);
                        }
                        else
                        {
                            lochunk33 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell33.Width = 25;
                            loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell33.AddElement(lochunk33);
                            loCell33.Border = 0;
                            loTable33.AddCell(loCell33);
                        }
                    }

                }
                document.Add(loTable33);

                #endregion

                #endregion

                #region Row4

                #region Signature Image Row4

                iTextSharp.text.Table loTable4 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable4.Border = 0;
                loTable4.Width = 100f;
                loTable4.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable4.Cellpadding = 0;
                loTable4.Cellspacing = 0;
                iTextSharp.text.Cell loCell4 = new Cell();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell4 = new iTextSharp.text.Cell();
                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k > 11 && k < 15)
                    {
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loCell4.Width = 25;
                            loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell4.AddElement(SignatureJpg);
                            loCell4.Border = 0;
                            loTable4.AddCell(loCell4);
                        }
                        else if (!File.Exists(Imagepath))
                        {
                            Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.SetAbsolutePosition(45, 555);
                            SignatureJpg.ScaleToFit(250, 42);
                            loEmptyChunky = new Cell();
                            loEmptyChunky.Width = 25;
                            loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loEmptyChunky.AddElement(SignatureJpg);
                            loEmptyChunky.Border = 0;
                            loTable4.AddCell(loEmptyChunky);
                            //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                            //loCell4.Width = 25;
                            //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            //loCell4.AddElement(loEmptyChunky);
                            //loCell4.Border = 0;
                            //loTable4.AddCell(loCell4);
                        }
                    }
                }
                document.Add(loTable4);
                #endregion

                #region Signature Text Row4

                iTextSharp.text.Table loTable44 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                loTable44.Border = 0;
                loTable44.Width = 100f;
                loTable44.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                loTable44.Cellpadding = 0;
                loTable44.Cellspacing = 0;
                iTextSharp.text.Cell loCell44 = new Cell();
                iTextSharp.text.Chunk lochunk44 = new Chunk();

                for (int k = 0; k < SignatureText.Length; k++)
                {
                    loCell44 = new iTextSharp.text.Cell();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                    if (k > 11 && k < 15)
                    {
                        if (File.Exists(Imagepath))
                        {
                            lochunk44 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell44.Width = 25;
                            loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell44.AddElement(lochunk44);
                            loCell44.Border = 0;
                            loTable44.AddCell(loCell44);
                        }
                        else
                        {
                            lochunk44 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            loCell44.Width = 25;
                            loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                            loCell44.AddElement(lochunk44);
                            loCell44.Border = 0;
                            loTable44.AddCell(loCell44);
                        }
                    }

                }
                document.Add(loTable44);

                #endregion

                #endregion
            }
            else if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "2")
            {
                Paragraph pimage = new Paragraph();
                Paragraph PSignature = new Paragraph();
                for (int k = 0; k < SignatureText.Length; k++)
                {
                    pimage = new Paragraph();
                    PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";
                    if (File.Exists(Imagepath))
                    {
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);

                        Chunk lochunkSign = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                        PSignature.Add(lochunkSign);
                        document.Add(PSignature);
                    }
                    else
                    {
                        Chunk lochunkSign = new Chunk("\n\n\n\n\n" + SignatureText[k], setFontsAll(11, 0, 0));
                        PSignature.Add(lochunkSign);
                        document.Add(PSignature);
                    }
                }
            }
            else
            {
                Paragraph pimage = new Paragraph();
                Paragraph PSignature = new Paragraph();

                string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "Ted Neild.jpg";
                if (File.Exists(Imagepath))
                {
                    iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                    SignatureJpg.ScaleToFit(250, 42);
                    pimage.Add(SignatureJpg);
                    document.Add(pimage);

                    Chunk lochunkSign = new Chunk("Ted Neild", setFontsAll(11, 0, 0));
                    PSignature.Add(lochunkSign);
                    document.Add(PSignature);

                }
            }
            #endregion
        }

        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    #endregion

    #region Gresham Advisors General Letter Regarding Fund Custom

    public string GetGreshamAdvisorsGLRFund()
    {
        string[] SignatureText = new string[10000];
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.GreshamAdvisorsGLRFund); //"SP_S_GreshamAdvisorsGLRFund 13, '18511645-77CA-E111-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE11-8391-001D09665E8F', 'F7063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();


        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 50, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "GreshamAdvisorsGLRFund.pdf";
        PdfWriter pdfwriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        AddFooter(document);

        document.Open();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + System.Guid.NewGuid().ToString() + ".xls";

        try
        {
            if (DSCount > 0)
            {

                //for (int i = 0; i < DSCount; i++)
                //{

                iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, newdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                iTextSharp.text.Cell loCell = new Cell();
                setTableProperty(loTable, ReportType.GreshamAdvisorsGLRFund);
                iTextSharp.text.Chunk lochunk = new Chunk();

                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\GreshamAdvisors_logo.tif");
                png.SetAbsolutePosition(48, 800);//540
                //png.ScaleToFit(288f, 42f);
                png.ScalePercent(22);
                document.Add(png);

                string AsOfDate = "\n\n";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]) != "")
                {
                    AsOfDate = AsOfDate + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]);
                }
                Chunk lochunkAsOfDate = new Chunk(AsOfDate, setFontsAll(8, 1, 0));
                Paragraph pAsOfDate = new Paragraph();
                pAsOfDate.Add(lochunkAsOfDate);
                pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                document.Add(pAsOfDate);


                string FullName = "\n\n";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]) != "")
                {
                    FullName = FullName + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_fullname_mail"]);
                }
                Chunk lochunkFullName = new Chunk(FullName, setFontsAll(9, 0, 0));
                Paragraph pFullName = new Paragraph();
                pFullName.Add(lochunkFullName);
                pFullName.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pFullName.Leading = 11f;
                document.Add(pFullName);

                string AddressLine1 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]) != "")
                {
                    AddressLine1 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]);
                }
                Chunk lochunkAddressLine1 = new Chunk(AddressLine1, setFontsAll(9, 0, 0));
                Paragraph pAddressLine1 = new Paragraph();
                pAddressLine1.Add(lochunkAddressLine1);
                pAddressLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine1.Leading = 11f;
                document.Add(pAddressLine1);

                string AddressLine2 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]) != "")
                {
                    AddressLine2 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]);
                }
                Chunk lochunkAddressLine2 = new Chunk(AddressLine2, setFontsAll(9, 0, 0));
                Paragraph pAddressLine2 = new Paragraph();
                pAddressLine2.Add(lochunkAddressLine2);
                pAddressLine2.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine2.Leading = 11f;
                document.Add(pAddressLine2);

                string AddressLine3 = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]) != "")
                {
                    AddressLine3 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]);
                }
                Chunk lochunkAddressLine3 = new Chunk(AddressLine3, setFontsAll(9, 0, 0));
                Paragraph pAddressLine3 = new Paragraph();
                pAddressLine3.Add(lochunkAddressLine3);
                pAddressLine3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressLine3.Leading = 11f;
                document.Add(pAddressLine3);

                string AddressDetails = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) != "" || Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]) != "")
                {
                    AddressDetails = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]);
                }
                Chunk lochunkAddressDetails = new Chunk(AddressDetails, setFontsAll(9, 0, 0));
                Paragraph pAddressDetails = new Paragraph();
                pAddressDetails.Add(lochunkAddressDetails);
                pAddressDetails.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pAddressDetails.Leading = 11f;
                document.Add(pAddressDetails);


                string FullNameBold = "RE: ";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_FundName"]) != "")
                {
                    FullNameBold = FullNameBold + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_FundName"]);
                }
                Chunk lochunkFullName1 = new Chunk(FullNameBold, setFontsAll(9, 1, 0));
                Paragraph pFullName1 = new Paragraph();
                pFullName1.Add(lochunkFullName1);
                pFullName1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pFullName1.SpacingBefore = 12f;
                document.Add(pFullName1);


                string InstructionsStart = "";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) != "")
                {
                    InstructionsStart = "\nDear " + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_dear_mail"]) + " :";
                }
                Chunk lochunkSingleStart = new Chunk(InstructionsStart, setFontsAll(9, 0, 0));
                Paragraph pSingleStart = new Paragraph();
                pSingleStart.Add(lochunkSingleStart);
                pSingleStart.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                pSingleStart.SpacingAfter = 5f;
                document.Add(pSingleStart);

                int rowsize = newdataset.Tables[0].Rows.Count;
                int colsize = 2;

                if (rowsize > 0)
                {
                    iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                    iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                    hw.Style = styles;
                    //styles.LoadTagStyle("ol", "leading", "16,0");
                    //styles.LoadTagStyle("li", "face", "garamond");
                    //styles.LoadTagStyle("span", "size", "11pt");
                    //styles.LoadTagStyle("body", "font-family", "verdana");
                    //styles.LoadTagStyle("body", "font-size", "11pt");

                    ArrayList objects = null;

                    string FundSpecificDesc = "";
                    //List list = new List(List.UNORDERED, 10f);
                    //list.SetListSymbol("\u2022");
                    //list.IndentationLeft = 30f;
                    //for (int m = 0; m < rowsize; m++)
                    //{
                    if (Convert.ToString(table.Rows[0]["ssi_lettertext"]) != "")
                    {

                        FundSpecificDesc = Convert.ToString(table.Rows[0]["ssi_lettertext"]);

                        FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                        //Chunk NextLine1 = new Chunk("\n");
                        //Paragraph pNextLine1 = new Paragraph();
                        //pNextLine1.Add(NextLine1);
                        //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        //document.Add(pNextLine1);

                        using (StringReader stringReader = new StringReader(FundSpecificDesc))
                        {
                            //List<IElement> parsedList = new List<IElement>();

                            objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                            //document.Open();
                            foreach (object item in objects)
                            {
                                if (item is List)
                                {
                                    List list = item as List;
                                    list.Autoindent = false;
                                    list.IndentationLeft = 25f;
                                    list.SetListSymbol("\u2022                      ");
                                    list.SymbolIndent = 25f;
                                    document.Add((IElement)item);
                                }


                                if (item is Paragraph)
                                {
                                    Paragraph para = item as Paragraph; //setFontsverdana

                                    if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                    {
                                        ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                        ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                        para.IndentationLeft = 30f;
                                        document.Add(para);
                                    }
                                    else
                                    {
                                        //for (int j = 0; j < para.ToArray().Length; j++)
                                        //{
                                        //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                        //}
                                        document.Add(para);
                                    }
                                }


                            }
                            //document.Close();
                        }
                        //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                        //Chunk NextLine = new Chunk("\n");
                        //Paragraph pNextLine = new Paragraph();
                        //pNextLine.Add(NextLine);
                        //pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        //document.Add(pNextLine);
                    }
                    //}
                    //document.Add(list);
                    for (int m = 0; m < rowsize; m++)
                    {
                        if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                        {
                            if (Convert.ToBoolean(table.Rows[m]["ssi_fundSpecificFlg"]) == true)
                            {
                                FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);

                                FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                                //Chunk NextLine1 = new Chunk("\n");
                                //Paragraph pNextLine1 = new Paragraph();
                                //pNextLine1.Add(NextLine1);
                                //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                //document.Add(pNextLine1);

                                using (StringReader stringReader = new StringReader(FundSpecificDesc))
                                {
                                    //List<IElement> parsedList = new List<IElement>();

                                    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                                    //document.Open();
                                    foreach (object item in objects)
                                    {
                                        if (item is List)
                                        {
                                            List list = item as List;
                                            list.Autoindent = false;
                                            list.IndentationLeft = 25f;
                                            list.SetListSymbol("\u2022                      ");
                                            list.SymbolIndent = 25f;
                                            document.Add((IElement)item);
                                        }


                                        if (item is Paragraph)
                                        {
                                            Paragraph para = item as Paragraph; //setFontsverdana

                                            if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                            {
                                                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                                para.IndentationLeft = 30f;
                                                document.Add(para);
                                            }
                                            else
                                            {
                                                //for (int j = 0; j < para.ToArray().Length; j++)
                                                //{
                                                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                                //}
                                                document.Add(para);
                                            }
                                        }


                                    }
                                    //document.Close();
                                }
                                //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));
                                if (m != rowsize - 1)
                                {
                                    Chunk NextLine = new Chunk("\n");
                                    Paragraph pNextLine = new Paragraph();
                                    pNextLine.Add(NextLine);
                                    pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    document.Add(pNextLine);
                                }
                            }
                        }
                    }

                    double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                    if (remainingPageSpace < 150.00)
                        document.NewPage();

                    Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                    Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
                    plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
                    plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                    plochunkSingleFundNONSTD3.SpacingBefore = 8f;
                    document.Add(plochunkSingleFundNONSTD3);

                }

                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]) != "")
                {
                    SignatureText = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]).Split(',');
                }


                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]) != "")
                {
                    DynamicFlg = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]);
                }


                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]) != "")
                {
                    Orientation = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]);
                }

                //}
                #region Signature

                if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "1")
                {
                    iTextSharp.text.Cell loEmptyChunky = new Cell();
                    #region Row1

                    #region Signature Image Row1

                    iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable1.Border = 0;
                    loTable1.Width = 100f;

                    loTable1.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable1.Cellpadding = 0;
                    loTable1.Cellspacing = 0;
                    iTextSharp.text.Cell loCell1 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell1 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k < 4)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell1.Width = 25;
                                loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell1.AddElement(SignatureJpg);
                                loCell1.Border = 0;
                                loTable1.AddCell(loCell1);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loCell1.Border = 0;
                                loTable1.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n", setFontsAll(11, 0, 0));
                                //loCell1.Width = 25;
                                //loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell1.AddElement(loEmptyChunky);
                                //loCell1.Border = 0;
                                //loTable1.AddCell(loCell1);
                            }
                        }
                    }
                    document.Add(loTable1);
                    #endregion

                    #region Signature Text Row1

                    iTextSharp.text.Table loTable11 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable11.Border = 0;
                    loTable11.Width = 100f;
                    loTable11.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable11.Cellpadding = 0;
                    loTable11.Cellspacing = 0;
                    iTextSharp.text.Cell loCell11 = new Cell();
                    iTextSharp.text.Chunk lochunk11 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell11 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k < 4)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk11 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell11.Width = 25;
                                loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell11.AddElement(lochunk11);
                                loCell11.Border = 0;
                                loTable11.AddCell(loCell11);
                            }
                            else
                            {
                                lochunk11 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell11.Width = 25;
                                loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell11.AddElement(lochunk11);
                                loCell11.Border = 0;
                                loTable11.AddCell(loCell11);
                            }
                        }

                    }
                    document.Add(loTable11);

                    #endregion

                    #endregion

                    #region Row2

                    #region Signature Image Row2

                    iTextSharp.text.Table loTable2 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable2.Border = 0;
                    loTable2.Width = 100f;
                    loTable2.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable2.Cellpadding = 0;
                    loTable2.Cellspacing = 0;
                    iTextSharp.text.Cell loCell2 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell2 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 3 && k < 8)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell2.Width = 25;
                                loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell2.AddElement(SignatureJpg);
                                loCell2.Border = 0;
                                loTable2.AddCell(loCell2);
                            }
                            else //if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable2.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell2.Width = 25;
                                //loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell2.AddElement(loEmptyChunky);
                                //loCell2.Border = 0;
                                //loTable2.AddCell(loCell2);
                            }
                        }
                    }
                    document.Add(loTable2);
                    #endregion

                    #region Signature Text Row2

                    iTextSharp.text.Table loTable22 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable22.Border = 0;
                    loTable22.Width = 100f;
                    loTable22.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable22.Cellpadding = 0;
                    loTable22.Cellspacing = 0;
                    iTextSharp.text.Cell loCell22 = new Cell();
                    iTextSharp.text.Chunk lochunk22 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell22 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 3 && k < 8)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk22 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell22.Width = 25;
                                loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell22.AddElement(lochunk22);
                                loCell22.Border = 0;
                                loTable22.AddCell(loCell22);
                            }
                            else
                            {
                                lochunk22 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell22.Width = 25;
                                loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell22.AddElement(lochunk22);
                                loCell22.Border = 0;
                                loTable22.AddCell(loCell22);
                            }
                        }

                    }
                    document.Add(loTable22);

                    #endregion

                    #endregion

                    #region Row3

                    #region Signature Image Row3

                    iTextSharp.text.Table loTable3 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable3.Border = 0;
                    loTable3.Width = 100f;
                    loTable3.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable3.Cellpadding = 0;
                    loTable3.Cellspacing = 0;
                    iTextSharp.text.Cell loCell3 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell3 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 7 && k < 12)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell3.Width = 25;
                                loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell3.AddElement(SignatureJpg);
                                loCell3.Border = 0;
                                loTable3.AddCell(loCell3);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable3.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell3.Width = 25;
                                //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell3.AddElement(loEmptyChunky);
                                //loCell3.Border = 0;
                                //loTable3.AddCell(loCell3);
                            }
                        }
                    }
                    document.Add(loTable3);
                    #endregion

                    #region Signature Text Row3

                    iTextSharp.text.Table loTable33 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable33.Border = 0;
                    loTable33.Width = 100f;
                    loTable33.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable33.Cellpadding = 0;
                    loTable33.Cellspacing = 0;
                    iTextSharp.text.Cell loCell33 = new Cell();
                    iTextSharp.text.Chunk lochunk33 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell33 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 7 && k < 12)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk33 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell33.Width = 25;
                                loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell33.AddElement(lochunk33);
                                loCell33.Border = 0;
                                loTable33.AddCell(loCell33);
                            }
                            else
                            {
                                lochunk33 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell33.Width = 25;
                                loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell33.AddElement(lochunk33);
                                loCell33.Border = 0;
                                loTable33.AddCell(loCell33);
                            }
                        }

                    }
                    document.Add(loTable33);

                    #endregion

                    #endregion

                    #region Row4

                    #region Signature Image Row4

                    iTextSharp.text.Table loTable4 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable4.Border = 0;
                    loTable4.Width = 100f;
                    loTable4.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable4.Cellpadding = 0;
                    loTable4.Cellspacing = 0;
                    iTextSharp.text.Cell loCell4 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell4 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 11 && k < 15)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCell4.Width = 25;
                                loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell4.AddElement(SignatureJpg);
                                loCell4.Border = 0;
                                loTable4.AddCell(loCell4);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable4.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell4.Width = 25;
                                //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell4.AddElement(loEmptyChunky);
                                //loCell4.Border = 0;
                                //loTable4.AddCell(loCell4);
                            }
                        }
                    }
                    document.Add(loTable4);
                    #endregion

                    #region Signature Text Row4

                    iTextSharp.text.Table loTable44 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable44.Border = 0;
                    loTable44.Width = 100f;
                    loTable44.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable44.Cellpadding = 0;
                    loTable44.Cellspacing = 0;
                    iTextSharp.text.Cell loCell44 = new Cell();
                    iTextSharp.text.Chunk lochunk44 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell44 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 11 && k < 15)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk44 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell44.Width = 25;
                                loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell44.AddElement(lochunk44);
                                loCell44.Border = 0;
                                loTable44.AddCell(loCell44);
                            }
                            else
                            {
                                lochunk44 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell44.Width = 25;
                                loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell44.AddElement(lochunk44);
                                loCell44.Border = 0;
                                loTable44.AddCell(loCell44);
                            }
                        }

                    }
                    document.Add(loTable44);

                    #endregion

                    #endregion

                }
                else if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "2")
                {
                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();
                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        pimage = new Paragraph();
                        PSignature = new Paragraph();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);

                            Chunk lochunkSign = new Chunk(SignatureText[k], setFontsAll(9, 0, 0));
                            PSignature.Add(lochunkSign);
                            document.Add(PSignature);
                        }
                        else
                        {
                            Chunk lochunkSign = new Chunk("\n\n\n\n\n" + SignatureText[k], setFontsAll(9, 0, 0));
                            PSignature.Add(lochunkSign);
                            document.Add(PSignature);
                        }
                    }
                }
                else
                {
                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "Ted Neild.jpg";
                    if (File.Exists(Imagepath))
                    {
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);

                        Chunk lochunkSign = new Chunk("Ted Neild", setFontsAll(9, 0, 0));
                        PSignature.Add(lochunkSign);
                        document.Add(PSignature);

                    }
                }
                #endregion

            }
        }
        catch
        {
        }




        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");

    }


    #endregion

    #region Fund Memorandum

    public string GetFundMemoradum()
    {
        string[] SignatureText = new string[10000];
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.FundMemorandum);// "SP_S_MemorandumRegardingFund 13, '18511645-77CA-E111-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE11-8391-001D09665E8F', 'F7063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        string LEname = string.Empty;
        if (DSCount > 0)
            LEname = GeneralMethods.RemoveSpecialCharacters(GetLEName(Convert.ToString(table.Rows[0]["ssi_legalentitynameidname"])));

        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
        Random rand = new Random();


        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 80, 30, 31, 8);//10,10
        String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + System.DateTime.Now.ToString("MMddyyhhmmss") + System.Guid.NewGuid().ToString() + "FundMemoradum.pdf";
        PdfWriter pdfwriter = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //AddHeader(document);
        AddFooter(document);

        document.Open();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + LEname + "_" + strGUID + System.Guid.NewGuid().ToString() + "FundMemoradum.xls";

        try
        {
            if (DSCount > 0)
            {

                //for (int i = 0; i < DSCount; i++)
                //{
                //if (i != 0)
                //{
                //    document.NewPage();
                //}
                Chunk BlankChunk = new Chunk("", setFontsAll(9, 0, 0));
                Paragraph BalnkPara = new Paragraph();
                BalnkPara.Add(BlankChunk);
                document.Add(BalnkPara);

                iTextSharp.text.Table loTable = new iTextSharp.text.Table(2, 5);   // 2 rows, 2 columns           
                iTextSharp.text.Cell loCell = new Cell();
                iTextSharp.text.Cell loCell1 = new Cell();
                iTextSharp.text.Cell loCell2 = new Cell();
                iTextSharp.text.Cell loCell3 = new Cell();
                iTextSharp.text.Cell loCell4 = new Cell();
                iTextSharp.text.Cell loCell5 = new Cell();
                iTextSharp.text.Cell loCell6 = new Cell();
                setTableProperty(loTable, ReportType.FundMemorandum);
                iTextSharp.text.Chunk lochunk = new Chunk();

                iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\GreshamAdvisors_logo.tif");
                png.SetAbsolutePosition(48, 800);//540
                //png.ScaleToFit(288f, 42f);
                png.ScalePercent(22);
                document.Add(png);


                //lochunk = new Chunk("\n\n ", setFontsAll(9, 0, 0));
                lochunk = new Chunk("", setFontsAll(9, 0, 0));
                loCell6 = new iTextSharp.text.Cell();
                loCell6.Add(lochunk);
                //loCell6.Colspan = DSCount + DSCount;
                loCell6.Colspan = 2;
                loCell6.Border = 0;
                loCell6.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loTable.AddCell(loCell6);

                lochunk = new Chunk(" ", setFontsAll(9, 0, 0));
                loCell = new iTextSharp.text.Cell();
                loCell.Add(lochunk);
                //loCell.Colspan = DSCount + DSCount;
                //loCell.BorderWidthTop = 2f;
                loCell.Colspan = 2;

                loCell.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loCell.Leading = 11F;
                //loTable.AddCell(loCell);


                Chunk lochunkHeading = new Chunk("\nMEMORANDUM", setFontsAll(9, 1, 0));
                loCell1.Add(lochunkHeading);
                loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                //loCell1.Colspan = DSCount + DSCount;
                loCell1.Colspan = 2;
                //loCell1.EnableBorderSide(2);
                loCell1.Border = 0;
                loTable.AddCell(loCell1);



                //Paragraph pHeading = new Paragraph();
                //pHeading.Add(lochunkHeading);
                //pHeading.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                //document.Add(pTO);


                string StrTo = "TO: ";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]) != "")
                {
                    StrTo = StrTo + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_legalentitynameidname"]);
                }
                Chunk lochunkTo = new Chunk(StrTo, setFontsAll(9, 1, 0));
                loCell2.Add(lochunkTo);
                loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loCell2.Border = 0;
                //loCell2.Colspan = DSCount + DSCount;
                loCell2.Colspan = 2;
                loTable.AddCell(loCell2);


                Chunk lochunkFROM = new Chunk("FROM: GRESHAM ADVISORS LLC", setFontsAll(11, 1, 0));
                loCell3.Add(lochunkFROM);
                loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loCell3.Border = 0;
                //loCell3.Colspan = DSCount + DSCount;
                loCell3.Colspan = 2;
                loTable.AddCell(loCell3);

                string StrAsofdate = "DATE: ";
                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]) != "")
                {
                    StrAsofdate = StrAsofdate + Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_asofdate"]);
                }
                Chunk lochunkStrAsofdate = new Chunk(StrAsofdate + "\n\n", setFontsAll(8, 1, 0));
                loCell4.Add(lochunkStrAsofdate);
                loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_LEFT;
                loCell4.Border = 0;
                //loCell4.Colspan = DSCount + DSCount;
                loCell4.Colspan = 2;
                loTable.AddCell(loCell4);



                document.Add(loTable);
                //Paragraph pFROM = new Paragraph();
                //pFROM.Add(lochunkFROM);
                //pFROM.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                //document.Add(pFROM);


                //Paragraph pStrAsofdate = new Paragraph();
                //pStrAsofdate.Add(lochunkStrAsofdate);
                //pStrAsofdate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                //document.Add(pStrAsofdate);

                int rowsize = newdataset.Tables[0].Rows.Count;
                int colsize = 2;

                if (rowsize > 0)
                {
                    iTextSharp.text.html.simpleparser.StyleSheet styles = new iTextSharp.text.html.simpleparser.StyleSheet();
                    iTextSharp.text.html.simpleparser.HTMLWorker hw = new iTextSharp.text.html.simpleparser.HTMLWorker(document);

                    hw.Style = styles;
                    //styles.LoadTagStyle("ol", "leading", "16,0");
                    //styles.LoadTagStyle("li", "face", "garamond");
                    //styles.LoadTagStyle("span", "size", "11pt");
                    //styles.LoadTagStyle("body", "font-family", "verdana");
                    //styles.LoadTagStyle("body", "font-size", "11pt");

                    ArrayList objects = null;

                    string FundSpecificDesc = "";
                    //List list = new List(List.UNORDERED, 10f);
                    //list.SetListSymbol("\u2022");
                    //list.IndentationLeft = 30f;
                    //for (int m = 0; m < rowsize; m++)
                    //{
                    if (Convert.ToString(table.Rows[0]["ssi_lettertext"]) != "")
                    {

                        FundSpecificDesc = Convert.ToString(table.Rows[0]["ssi_lettertext"]);
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                        //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                        FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                        //Chunk NextLine1 = new Chunk("\n");
                        //Paragraph pNextLine1 = new Paragraph();
                        //pNextLine1.Add(NextLine1);
                        //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        //document.Add(pNextLine1);

                        using (StringReader stringReader = new StringReader(FundSpecificDesc))
                        {
                            //List<IElement> parsedList = new List<IElement>();

                            objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                            //document.Open();
                            foreach (object item in objects)
                            {
                                if (item is List)
                                {
                                    List list = item as List;
                                    list.Autoindent = false;
                                    list.IndentationLeft = 25f;
                                    list.SetListSymbol("\u2022                      ");
                                    list.SymbolIndent = 25f;
                                    document.Add((IElement)item);
                                }


                                if (item is Paragraph)
                                {
                                    Paragraph para = item as Paragraph; //setFontsverdana

                                    if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                    {
                                        ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                        ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                        para.IndentationLeft = 30f;
                                        document.Add(para);
                                    }
                                    else
                                    {
                                        //for (int j = 0; j < para.ToArray().Length; j++)
                                        //{
                                        //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                        //}
                                        document.Add(para);
                                    }
                                }


                            }
                            //document.Close();
                        }

                        //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));

                        //Chunk NextLine = new Chunk("\n");
                        //Paragraph pNextLine = new Paragraph();
                        //pNextLine.Add(NextLine);
                        //pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                        //document.Add(pNextLine);
                    }
                    //}
                    //document.Add(list);

                    for (int m = 0; m < rowsize; m++)
                    {
                        if (Convert.ToString(table.Rows[m]["ssi_fundtxt"]) != "")
                        {
                            if (Convert.ToBoolean(table.Rows[m]["ssi_fundSpecificFlg"]) == true)
                            {
                                FundSpecificDesc = Convert.ToString(table.Rows[m]["ssi_fundtxt"]);
                                FundSpecificDesc = FundSpecificDesc.Replace("xx-small", "9pt").Replace("x-small", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("smaller", "9pt").Replace("larger", "9pt");
                                //FundSpecificDesc = FundSpecificDesc.Replace("small", "9pt").Replace("medium", "9pt").Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("x-large", "9pt");
                                FundSpecificDesc = FundSpecificDesc.Replace("xx-large", "9pt");
                                //Chunk NextLine1 = new Chunk("\n");
                                //Paragraph pNextLine1 = new Paragraph();
                                //pNextLine1.Add(NextLine1);
                                //pNextLine1.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                //document.Add(pNextLine1);

                                using (StringReader stringReader = new StringReader(FundSpecificDesc))
                                {
                                    //List<IElement> parsedList = new List<IElement>();

                                    objects = iTextSharp.text.html.simpleparser.HTMLWorker.ParseToList(stringReader, hw.Style);
                                    //document.Open();
                                    foreach (object item in objects)
                                    {
                                        if (item is List)
                                        {
                                            List list = item as List;
                                            list.Autoindent = false;
                                            list.IndentationLeft = 25f;
                                            list.SetListSymbol("\u2022                      ");
                                            list.SymbolIndent = 25f;
                                            document.Add((IElement)item);
                                        }


                                        if (item is Paragraph)
                                        {
                                            Paragraph para = item as Paragraph; //setFontsverdana

                                            if (para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "o" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "�" || para.ToArray()[0].ToString() == "v")
                                            {
                                                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font = setFontsverdana();
                                                ((iTextSharp.text.Chunk)(para.ToArray()[0])).Font.Size = 9f;
                                                para.IndentationLeft = 30f;
                                                document.Add(para);
                                            }
                                            else
                                            {
                                                //for (int j = 0; j < para.ToArray().Length; j++)
                                                //{
                                                //    ((iTextSharp.text.Chunk)(para.ToArray()[j])).Font = setFontsverdana();
                                                //}
                                                document.Add(para);
                                            }
                                        }


                                    }
                                    //document.Close();
                                }

                                //hw.Parse(new StringReader(Convert.ToString(table.Rows[m]["ssi_fundtxt"])));
                                if (m != rowsize - 1)
                                {
                                    Chunk NextLine = new Chunk("\n");
                                    Paragraph pNextLine = new Paragraph();
                                    pNextLine.Add(NextLine);
                                    pNextLine.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                                    document.Add(pNextLine);
                                }
                            }
                        }
                    }

                }

                double remainingPageSpace = pdfwriter.GetVerticalPosition(false) - document.BottomMargin;

                if (remainingPageSpace < 150.00)
                    document.NewPage();

                Chunk lochunkSingleFundNONSTD3 = new Chunk("" + "Sincerely Yours,", setFontsAll(9, 0, 0));
                Paragraph plochunkSingleFundNONSTD3 = new Paragraph();
                plochunkSingleFundNONSTD3.Add(lochunkSingleFundNONSTD3);
                plochunkSingleFundNONSTD3.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
                plochunkSingleFundNONSTD3.SpacingBefore = 8f;
                document.Add(plochunkSingleFundNONSTD3);

                if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]) != "")
                {
                    SignatureText = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_includesignatureline"]).Split(',');
                }


                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]) != "")
                {
                    DynamicFlg = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_dynamicflg"]);
                }


                if (Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]) != "")
                {
                    Orientation = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_orientation"]);
                }

                //}

                #region Signature

                if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "1")
                {
                    iTextSharp.text.Cell loEmptyChunky = new Cell();
                    #region Row1

                    #region Signature Image Row1

                    iTextSharp.text.Table loTable1 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable1.Border = 0;
                    loTable1.Width = 100f;

                    loTable1.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable1.Cellpadding = 0;
                    loTable1.Cellspacing = 0;
                    iTextSharp.text.Cell loCellS1 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCellS1 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k < 4)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCellS1.Width = 25;
                                loCellS1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCellS1.AddElement(SignatureJpg);
                                loCellS1.Border = 0;
                                loTable1.AddCell(loCellS1);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loCell1.Border = 0;
                                loTable1.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n", setFontsAll(11, 0, 0));
                                //loCell1.Width = 25;
                                //loCell1.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell1.AddElement(loEmptyChunky);
                                //loCell1.Border = 0;
                                //loTable1.AddCell(loCell1);
                            }
                        }
                    }
                    document.Add(loTable1);
                    #endregion

                    #region Signature Text Row1

                    iTextSharp.text.Table loTable11 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable11.Border = 0;
                    loTable11.Width = 100f;
                    loTable11.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable11.Cellpadding = 0;
                    loTable11.Cellspacing = 0;
                    iTextSharp.text.Cell loCell11 = new Cell();
                    iTextSharp.text.Chunk lochunk11 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell11 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k < 4)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk11 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell11.Width = 25;
                                loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell11.AddElement(lochunk11);
                                loCell11.Border = 0;
                                loTable11.AddCell(loCell11);
                            }
                            else
                            {
                                lochunk11 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell11.Width = 25;
                                loCell11.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell11.AddElement(lochunk11);
                                loCell11.Border = 0;
                                loTable11.AddCell(loCell11);
                            }
                        }

                    }
                    document.Add(loTable11);

                    #endregion

                    #endregion

                    #region Row2

                    #region Signature Image Row2

                    iTextSharp.text.Table loTable2 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable2.Border = 0;
                    loTable2.Width = 100f;
                    loTable2.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable2.Cellpadding = 0;
                    loTable2.Cellspacing = 0;
                    iTextSharp.text.Cell loCellS2 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCellS2 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 3 && k < 8)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCellS2.Width = 25;
                                loCellS2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCellS2.AddElement(SignatureJpg);
                                loCellS2.Border = 0;
                                loTable1.AddCell(loCellS2);
                            }
                            else //if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable2.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell2.Width = 25;
                                //loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell2.AddElement(loEmptyChunky);
                                //loCell2.Border = 0;
                                //loTable2.AddCell(loCell2);
                            }
                        }
                    }
                    document.Add(loTable2);
                    #endregion

                    #region Signature Text Row2

                    iTextSharp.text.Table loTable22 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable22.Border = 0;
                    loTable22.Width = 100f;
                    loTable22.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable22.Cellpadding = 0;
                    loTable22.Cellspacing = 0;
                    iTextSharp.text.Cell loCell22 = new Cell();
                    iTextSharp.text.Chunk lochunk22 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell22 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 3 && k < 8)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk22 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell22.Width = 25;
                                loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell22.AddElement(lochunk22);
                                loCell22.Border = 0;
                                loTable22.AddCell(loCell22);
                            }
                            else
                            {
                                lochunk22 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell22.Width = 25;
                                loCell22.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell22.AddElement(lochunk22);
                                loCell22.Border = 0;
                                loTable22.AddCell(loCell22);
                            }
                        }

                    }
                    document.Add(loTable22);

                    #endregion

                    #endregion

                    #region Row3

                    #region Signature Image Row3

                    iTextSharp.text.Table loTable3 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable3.Border = 0;
                    loTable3.Width = 100f;
                    loTable3.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable3.Cellpadding = 0;
                    loTable3.Cellspacing = 0;
                    iTextSharp.text.Cell loCellS3 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCellS3 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 7 && k < 12)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCellS3.Width = 25;
                                loCellS3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCellS3.AddElement(SignatureJpg);
                                loCellS3.Border = 0;
                                loTable1.AddCell(loCellS3);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable3.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell3.Width = 25;
                                //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell3.AddElement(loEmptyChunky);
                                //loCell3.Border = 0;
                                //loTable3.AddCell(loCell3);
                            }
                        }
                    }
                    document.Add(loTable3);
                    #endregion

                    #region Signature Text Row3

                    iTextSharp.text.Table loTable33 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable33.Border = 0;
                    loTable33.Width = 100f;
                    loTable33.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable33.Cellpadding = 0;
                    loTable33.Cellspacing = 0;
                    iTextSharp.text.Cell loCell33 = new Cell();
                    iTextSharp.text.Chunk lochunk33 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell33 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 7 && k < 12)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk33 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell33.Width = 25;
                                loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell33.AddElement(lochunk33);
                                loCell33.Border = 0;
                                loTable33.AddCell(loCell33);
                            }
                            else
                            {
                                lochunk33 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell33.Width = 25;
                                loCell33.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell33.AddElement(lochunk33);
                                loCell33.Border = 0;
                                loTable33.AddCell(loCell33);
                            }
                        }

                    }
                    document.Add(loTable33);

                    #endregion

                    #endregion

                    #region Row4

                    #region Signature Image Row4

                    iTextSharp.text.Table loTable4 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable4.Border = 0;
                    loTable4.Width = 100f;
                    loTable4.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable4.Cellpadding = 0;
                    loTable4.Cellspacing = 0;
                    iTextSharp.text.Cell loCellS4 = new Cell();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCellS4 = new iTextSharp.text.Cell();
                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 11 && k < 15)
                        {
                            if (File.Exists(Imagepath))
                            {
                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loCellS4.Width = 25;
                                loCellS4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCellS4.AddElement(SignatureJpg);
                                loCellS4.Border = 0;
                                loTable1.AddCell(loCellS4);
                            }
                            else if (!File.Exists(Imagepath))
                            {
                                Imagepath = HttpContext.Current.Server.MapPath("") + @"\images\ImageNotAvailable.jpg"; //Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "ImageNotAvailable.jpg";

                                iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                                SignatureJpg.SetAbsolutePosition(45, 555);
                                SignatureJpg.ScaleToFit(250, 42);
                                loEmptyChunky = new Cell();
                                loEmptyChunky.Width = 25;
                                loEmptyChunky.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loEmptyChunky.AddElement(SignatureJpg);
                                loEmptyChunky.Border = 0;
                                loTable4.AddCell(loEmptyChunky);
                                //Chunk loEmptyChunky = new Chunk("\n\n\n", setFontsAll(11, 0, 0));
                                //loCell4.Width = 25;
                                //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                //loCell4.AddElement(loEmptyChunky);
                                //loCell4.Border = 0;
                                //loTable4.AddCell(loCell4);
                            }
                        }
                    }
                    document.Add(loTable4);
                    #endregion

                    #region Signature Text Row4

                    iTextSharp.text.Table loTable44 = new iTextSharp.text.Table(4, 1);   // 2 rows, 2 columns 

                    loTable44.Border = 0;
                    loTable44.Width = 100f;
                    loTable44.Alignment = iTextSharp.text.Table.ALIGN_LEFT;
                    loTable44.Cellpadding = 0;
                    loTable44.Cellspacing = 0;
                    iTextSharp.text.Cell loCell44 = new Cell();
                    iTextSharp.text.Chunk lochunk44 = new Chunk();

                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        loCell44 = new iTextSharp.text.Cell();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";

                        if (k > 11 && k < 15)
                        {
                            if (File.Exists(Imagepath))
                            {
                                lochunk44 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell44.Width = 25;
                                loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell44.AddElement(lochunk44);
                                loCell44.Border = 0;
                                loTable44.AddCell(loCell44);
                            }
                            else
                            {
                                lochunk44 = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                                loCell44.Width = 25;
                                loCell44.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
                                loCell44.AddElement(lochunk44);
                                loCell44.Border = 0;
                                loTable44.AddCell(loCell44);
                            }
                        }

                    }
                    document.Add(loTable44);

                    #endregion

                    #endregion

                }
                else if (DynamicFlg.ToUpper() == "FALSE" && Orientation == "2")
                {
                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();
                    for (int k = 0; k < SignatureText.Length; k++)
                    {
                        pimage = new Paragraph();
                        PSignature = new Paragraph();

                        string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + SignatureText[k] + ".jpg";
                        if (File.Exists(Imagepath))
                        {
                            iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                            SignatureJpg.ScaleToFit(250, 42);
                            pimage.Add(SignatureJpg);
                            document.Add(pimage);

                            Chunk lochunkSign = new Chunk(SignatureText[k], setFontsAll(11, 0, 0));
                            PSignature.Add(lochunkSign);
                            document.Add(PSignature);
                        }
                        else
                        {
                            Chunk lochunkSign = new Chunk("\n\n\n\n\n" + SignatureText[k], setFontsAll(11, 0, 0));
                            PSignature.Add(lochunkSign);
                            document.Add(PSignature);
                        }
                    }
                }
                else
                {
                    Paragraph pimage = new Paragraph();
                    Paragraph PSignature = new Paragraph();

                    string Imagepath = AppLogic.GetParam(AppLogic.ConfigParam.ImagePath) + "Ted Neild.jpg";
                    if (File.Exists(Imagepath))
                    {
                        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(Imagepath);
                        SignatureJpg.ScaleToFit(250, 42);
                        pimage.Add(SignatureJpg);
                        document.Add(pimage);

                        Chunk lochunkSign = new Chunk("Ted Neild", setFontsAll(11, 0, 0));
                        PSignature.Add(lochunkSign);
                        document.Add(PSignature);

                    }

                }
                #endregion


            }
        }
        catch
        { }




        if (DSCount > 0)
        {
            try
            {
                document.Close();
                FileInfo loFile = new FileInfo(ls);
                loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
            }
            catch
            { }

        }
        else
        {
            fsFinalLocation = "";
        }
        return fsFinalLocation.Replace(".xls", ".pdf");
    }

    #endregion

    #region Upload PDF

    public string GetUploadPdf()
    {
        string[] SignatureText = new string[10000];
        string fsFinalLocation = string.Empty;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsSQL = getFinalSp(ReportType.UploadPdf);// "SP_S_MemorandumRegardingFund 13, '18511645-77CA-E111-AD83-0019B9E7EE05', 'F9DC3BE5-6D15-DE11-8391-001D09665E8F', 'F7063302-DD15-DE11-8391-001D09665E8F'";

        newdataset = clsDB.getDataSet(lsSQL);
        int DSCount = newdataset.Tables[0].Rows.Count;
        DataTable table = newdataset.Tables[0].Copy();

        for (int i = 0; i < table.Rows.Count; i++)
        {
            fsFinalLocation = AppLogic.GetParam(AppLogic.ConfigParam.FileUploadUrl) + Convert.ToString(table.Rows[i]["Ssi_FileName"]);
        }

        return fsFinalLocation;
    }


    #endregion

    #region Billing Invoice
    public string GetBillingInvoice()
    {

        try
        {
            DB clsDB = new DB();
            // String lsSQL = "SP_S_PDF_BILLING @MailID = 3004 , @AsOfDate = '2019-12-31'";//Store Procedure call

            String lsSQL = getFinalSp(ReportType.Invoice);//Store Procedure call
            DataSet newdataset = clsDB.getDataSet(lsSQL);

            DataTable dtFormat = newdataset.Tables[1].Copy();
            dtFormat.AcceptChanges();
            string strBillableAUM = string.Empty;

            string underline = string.Empty;
            string underline1 = string.Empty;

            int DSC = newdataset.Tables.Count;
            int DSCount = newdataset.Tables[0].Rows.Count;
            // string strGUID = System.DateTime.Now.ToString("MMddyyhhmmssfff");
            // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + "_Billing.pdf";

            var strGUID = Guid.NewGuid().ToString();

            String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strGUID + "_Billing.pdf";

            String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmssfff") + System.Guid.NewGuid().ToString() + "_Billing.pdf";

            iTextSharp.text.Document pdoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
            PdfWriter writer = PdfWriter.GetInstance(pdoc, new FileStream(ls, FileMode.Create));
            //AddFooter(pdoc);
            //Footer
            Phrase footPhraseImg = new Phrase("Gresham Partners, LLC | 333 W. Wacker Dr. Suite 700 | Chicago, IL 60606 | P 312.960.0200 | F 312.960.0204 | www.greshampartners.com", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
            footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
            footer.Alignment = Element.ALIGN_LEFT;
            pdoc.Footer = footer;

            pdoc.Open();

            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
            png.SetAbsolutePosition(48, 800);//540
            //png.ScaleToFit(288f, 42f);
            png.ScalePercent(10);
            pdoc.Add(png);



            int rowsize = 4;//table.Rows.Count;
            //  iTextSharp.text.Table loTable = new iTextSharp.text.Table(4);   // 2 rows, 2 columns         

            PdfPTable loTable = new PdfPTable(4);
            string lsTotalNumberofColumns = "4";

            int colsize = 4;

            //iTextSharp.text.Cell loCell = new Cell();

            PdfPCell loCell = new PdfPCell();

            // setTableProperty(loTable);

            #region Table Style
            int[] headerwidths9 = { 53, 18, 15, 14 };

            //  int[] headerwidths9 = { 32, 8, 8, 8, 8, 9, 9, 9, 9 };
            //   int[] headerwidths2 = { 100 };
            loTable.WidthPercentage = 100f;
            loTable.SetWidths(headerwidths9);
            loTable.HorizontalAlignment = 0;

            #endregion


            string strLetterDate = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_LetterDate"]);
            if (strLetterDate != "")
                strLetterDate = Convert.ToDateTime(strLetterDate).ToString("MMMM dd,yyyy");


            //strLetterDate = "November 12, 2015";
            //  string strInvYear = Convert.ToString(newdataset.Tables[0].Rows[0]["LetterDate"]);
            //  if (strInvYear != "")
            //  strInvYear = Convert.ToDateTime(strLetterDate).ToString("yyyy");





            string strInvNumber = Convert.ToString(newdataset.Tables[0].Rows[0]["InvoiceNumber"]); //"1111";
            string strContactFullname = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ContactFullName"]);
            string strAddress1 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]);
            string strAddress2 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]);
            string strAddress3 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]);
            string strCity = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]);
            string strState = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]);
            string strZip = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]);
            string strCountry = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_countryregion_mail"]);
            string strHouseHold = Convert.ToString(newdataset.Tables[0].Rows[0]["HouseHoldName"]); //"Adams Family";
            string strRefName = Convert.ToString(newdataset.Tables[0].Rows[0]["RefName"]); //"Adams Family";
            string strAsOfDate = Convert.ToString(newdataset.Tables[0].Rows[0]["DateRange"]); //"November 2015 � January 2016";

            string strAODmailmerge = Convert.ToString(newdataset.Tables[0].Rows[0]["AumAsOfDate"]);  //"September 30, 2015";
            if (strAODmailmerge != "")
                strAODmailmerge = Convert.ToDateTime(strAODmailmerge).ToString("MMMM dd,yyyy");

            strBillableAUM = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["BillableAumAmt"])); //"$ 34,262,588";


            string strBillableAUM1 = strBillableAUM.Replace("$", "").Replace(" ", "");


            string strFeeRate = Percentage(Convert.ToString(newdataset.Tables[0].Rows[0]["FeeRatePct"]));//"0.68%";
            string strAnnualFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["AnnualisedFeeAmt"]));//"$ 233,813";
            string strQuaterFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])); //"$58,453";

            //string strQuaterFee = RoundUp(Convert.ToString(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])); //"$58,453";


            //decimal strQuaterFee1 = Decimal.Round((Convert.ToDecimal(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])), 2, MidpointRounding.AwayFromZero);

            //string strQuaterFee2 = Convert.ToString(strQuaterFee1);

            // string strQuaterFee1 =((Convert.ToString(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])), 2, MidpointRounding.AwayFromZero);

            string strAdjustment = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Adjustment"]));
            string strAdjustedFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_AdjustedFee"]));
            string strAdjustmentReason = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_AdjustmentReason"]);


            //string strRelationshipFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_RelationshipFee"]));


            //string strFeeName = (Convert.ToString(newdataset.Tables[2].Rows[0]["ssi_name"]));


            //string strAmount = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[2].Rows[0]["ssi_Amount"]));
            //string strAmount1 = strAmount.Replace("$", "").Replace(" ", "");


            //string strSetupFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_SetUpFee"]));
            //string strSetupFee1 = strSetupFee.Replace("$", "").Replace(" ", "");


            string strAnnFeeBeforeAdditionalFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["AnnFeeBeforeAdditionalFee"]));

            string strAnnFeeBeforeAdditionalFee1 = strAnnFeeBeforeAdditionalFee.Replace("$", "").Replace(" ", "");

            string Accrued = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Accrued"]);


            //  string Accrued = "True";

            string Monthly1 = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_Month1Fee"]));

            string strDiscount = "0.0";

            // string strDiscount = "7.145";

            if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Discount"]) != "")
                strDiscount = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Discount"]);

            string strInvYear = Convert.ToString(newdataset.Tables[0].Rows[0]["AumAsOfDate"]);

            if (strInvYear != "")
                strInvYear = Convert.ToDateTime(strAODmailmerge).ToString("yyyy");

            strInvYear = DateTime.Now.Year.ToString();
            string strAutoDebitDt = Convert.ToString(newdataset.Tables[0].Rows[0]["AutoDebitDt"]); //"November 19, 2015";

            if (strAutoDebitDt != "")
                strAutoDebitDt = Convert.ToDateTime(strAutoDebitDt).ToString("MMMM d, yyyy");

            //   PdfPTable tblMain = new PdfPTable(2);

            PdfPTable tblMain = new PdfPTable(2);
            int[] widthHeaderMain = { 80, 20 };
            tblMain.SetWidths(widthHeaderMain);
            tblMain.TotalWidth = 100f;
            tblMain.WidthPercentage = 100f;


            tblMain.DefaultCell.Border = 1;

            PdfPCell cellHeaderleft = new PdfPCell();
            PdfPCell cellHeaderright = new PdfPCell();
            PdfPCell cellAddress = new PdfPCell();
            PdfPCell cellRef = new PdfPCell();
            PdfPCell cellline4 = new PdfPCell();
            PdfPCell cellAsOfDate = new PdfPCell();
            PdfPCell cellline6left_1 = new PdfPCell();
            PdfPCell cellline6right_1 = new PdfPCell();


            PdfPCell cellline6left_2 = new PdfPCell();
            PdfPCell cellline6right_2 = new PdfPCell();

            PdfPCell cellline6left1 = new PdfPCell();
            PdfPCell cellline6right1 = new PdfPCell();


            PdfPCell celllineT12 = new PdfPCell();



            PdfPCell cellline7left = new PdfPCell();
            PdfPCell cellline7right = new PdfPCell();
            PdfPCell cellline8left = new PdfPCell();
            PdfPCell cellline8right = new PdfPCell();
            PdfPCell cellline9 = new PdfPCell(); //Border above -- blank

            PdfPCell celllineBlank = new PdfPCell(); //Border above -- blank
            PdfPCell celllineBlankright = new PdfPCell();

            PdfPCell celllineBlank1 = new PdfPCell(); //Border above -- blank
            PdfPCell celllineBlankright1 = new PdfPCell();

            PdfPCell cellline10left = new PdfPCell();
            PdfPCell cellline10right = new PdfPCell();

            PdfPCell cellline10left1 = new PdfPCell();
            PdfPCell cellline10right1 = new PdfPCell();

            PdfPCell cellline10noteleft = new PdfPCell();
            PdfPCell cellline10noteright = new PdfPCell();

            //**
            //New Section Added for adjustment 
            //**
            PdfPCell cellline1AdjustmentLeft = new PdfPCell();
            PdfPCell cellline1AdjustmentRight = new PdfPCell();

            PdfPCell cellline2AdjustmentLeft = new PdfPCell();
            PdfPCell cellline2AdjustmentRight = new PdfPCell();
            //**

            PdfPCell cellline11 = new PdfPCell();
            PdfPCell cellline12 = new PdfPCell();
            PdfPCell cellline13 = new PdfPCell();
            PdfPCell cellline14 = new PdfPCell();


            cellHeaderleft.Border = 0;
            cellHeaderright.Border = 0;
            cellAddress.Border = 0;
            cellRef.Border = 0;
            cellline4.Border = 0;
            cellAsOfDate.Border = 0;
            //cellline6left.Border = 0;
            //cellline6right.Border = 0;

            cellline6left1.Border = 0;
            cellline6right1.Border = 0;

            celllineT12.Border = 0;



            cellline7left.Border = 0;
            cellline7right.Border = 0;
            cellline8left.Border = 0;
            cellline8right.Border = 0;
            cellline9.Border = 0;

            celllineBlank.Border = 0;
            celllineBlankright.Border = 0;

            celllineBlank1.Border = 0;
            celllineBlankright1.Border = 0;

            cellline10left.Border = 0;
            cellline10right.Border = 0;

            cellline10noteleft.Border = 0;
            cellline10noteright.Border = 0;
            cellline1AdjustmentLeft.Border = 0;
            cellline1AdjustmentRight.Border = 0;
            cellline2AdjustmentLeft.Border = 0;
            cellline2AdjustmentRight.Border = 0;
            cellline11.Border = 0;
            cellline12.Border = 0;
            cellline13.Border = 0;
            cellline14.Border = 0;

            cellHeaderleft.PaddingTop = 65f;
            cellHeaderright.PaddingTop = 35f;

            cellAddress.PaddingTop = 60f;

            cellRef.PaddingTop = 60f;

            cellline4.PaddingTop = 36f;

            cellAsOfDate.PaddingTop = 10f;

            //cellline6left.PaddingTop = 10f;
            //cellline6right.PaddingTop = 10f;


            cellline6left1.PaddingTop = 10f;
            cellline6right1.PaddingTop = 10f;

            cellline8left.PaddingBottom = 10f;
            cellline8right.PaddingBottom = 10f;

            cellline11.PaddingTop = 10f;
            cellline12.PaddingTop = 10f;

            celllineT12.PaddingTop = 10f;

            cellline13.PaddingTop = 10f;
            cellline14.PaddingTop = 10f;

            int[] widthHeader = { 60, 40 };
            tblMain.SetWidths(widthHeader);
            tblMain.TotalWidth = 100f;
            tblMain.WidthPercentage = 100f;

            Chunk asterisk = new Chunk("*", setFontsAll(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            asterisk.SetTextRise(4);


            Chunk dollar = new Chunk("$", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Chunk glue = new Chunk(new VerticalPositionMark());
            //  asterisk.SetTextRise(4);

            /***    First Line  ***/


            Paragraph pLetterDate = new Paragraph(strLetterDate.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph pInvoice = new Paragraph("Invoice " + strInvYear.ToString() + "-" + strInvNumber.ToString(), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            pInvoice.SetAlignment("right");
            pLetterDate.Leading = 12f;
            pInvoice.Leading = 12f;

            /*** Address ***/
            Paragraph pContactFullName = new Paragraph(strContactFullname.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph pAddress1 = new Paragraph(strAddress1.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph pAddress2 = new Paragraph(strAddress2.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph pAddress3 = new Paragraph(strAddress3.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph pCityState = new Paragraph();
            Chunk pCity = new Chunk(strCity + ", ".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Chunk pState = new Chunk(strState + " ".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Chunk pZipCode = new Chunk(strZip + " ".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph pCountry = new Paragraph(strCountry + "".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            if (!string.IsNullOrEmpty(strCity))
                pCityState.Add(pCity);

            if (!string.IsNullOrEmpty(strState))
                pCityState.Add(pState);

            if (!string.IsNullOrEmpty(strZip))
                pCityState.Add(pZipCode);

            //if (!string.IsNullOrEmpty(strCountry))
            //    pCityState.Add(pCountry);

            pContactFullName.Leading = 12f;
            pAddress1.Leading = 12f;
            pAddress2.Leading = 12f;
            pAddress3.Leading = 12f;
            pCityState.Leading = 12f;
            pCountry.Leading = 12f;
            /*** REF ***/
            Paragraph pREF = new Paragraph("REF: " + strRefName, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pREF.Leading = 12f;

            /*** 4th Line ***/
            Paragraph p4 = new Paragraph("Invoice for financial counseling and investment services for the period: ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            p4.Leading = 12f;

            /*** As of Date ***/
            Paragraph pAsofdate = new Paragraph(strAsOfDate, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pAsofdate.SetAlignment("center");
            pAsofdate.Leading = 12f;
            /*** 6th Line ***/
            // Paragraph p6 = new Paragraph();

            strBillableAUM = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["BillableAumAmt"]));

            decimal BillAmt = Convert.ToDecimal(strBillableAUM1);

            Paragraph PBlank = new Paragraph(" ", setFontsAllTimesNewRoman(10, 1, 0));

            string text = "Annualized \n";

            text = text + "fees";

            Paragraph pheadercol1 = null;
            PdfPCell celllheader = null;

            Chunk loheaderchunk = null;

            pheadercol1 = new Paragraph("Billable Asset(s) as of " + strAODmailmerge.ToString().Replace("''", "'"), setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //  pheadercol1.SetAlignment("right");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheader = new PdfPCell();
            celllheader.Border = 0;

            celllheader.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));


            celllheader.AddElement(pheadercol1);

            //celllheader.BackgroundColor = iTextSharp.text.Color(128, 148, 170);
            loTable.AddCell(celllheader);

            pheadercol1 = new Paragraph("Assets", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pheadercol1.SetAlignment("center");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheader = new PdfPCell();
            celllheader.Border = 0;
            celllheader.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));
            celllheader.PaddingLeft = 20f;
            celllheader.AddElement(pheadercol1);
            loTable.AddCell(celllheader);

            pheadercol1 = new Paragraph("Annualized Fee", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            //   pheadercol1 = new Paragraph();
            //   loheaderchunk = new Chunk("Annualized" , setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));


            //   pheadercol1.Add(loheaderchunk);
            //   loheaderchunk = new Chunk("    Fee", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            ////   pheadercol1 = new Paragraph();
            //   pheadercol1.Add(loheaderchunk);

            //  pheadercol1 = new Paragraph(loheaderchunk);

            pheadercol1.SetAlignment("center");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheader = new PdfPCell();
            celllheader.Border = 0;
            celllheader.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));
            celllheader.AddElement(pheadercol1);
            loTable.AddCell(celllheader);

            pheadercol1 = new Paragraph("Agreed upon Fee Rate(s)", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pheadercol1.SetAlignment("right");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheader = new PdfPCell();
            celllheader.Border = 0;
            celllheader.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));
            celllheader.AddElement(pheadercol1);
            loTable.AddCell(celllheader);

            iTextSharp.text.Paragraph lochunk = new Paragraph();
            iTextSharp.text.Chunk lochunknew = new Chunk();


            #region Accruedtable
            PdfPTable loAccuredTbl = new PdfPTable(4);

            int[] widthHeaderAcurdtbl = { 80, 0, 12, 0 };
            loAccuredTbl.SetWidths(widthHeaderAcurdtbl);
            loAccuredTbl.TotalWidth = 100f;
            loAccuredTbl.WidthPercentage = 100f;


            Paragraph pheadercolAcrdTbl = null;
            PdfPCell celllheaderAcrdTbl = null;


            pheadercolAcrdTbl = new Paragraph("Description", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pheadercolAcrdTbl.SetAlignment("center");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheaderAcrdTbl = new PdfPCell();
            celllheaderAcrdTbl.Border = 0;
            celllheaderAcrdTbl.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));
            celllheaderAcrdTbl.AddElement(pheadercolAcrdTbl);
            loAccuredTbl.AddCell(celllheaderAcrdTbl);


            pheadercolAcrdTbl = new Paragraph("", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pheadercolAcrdTbl.SetAlignment("right");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheaderAcrdTbl = new PdfPCell();
            celllheaderAcrdTbl.Border = 0;
            celllheaderAcrdTbl.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));
            celllheaderAcrdTbl.AddElement(pheadercolAcrdTbl);
            loAccuredTbl.AddCell(celllheaderAcrdTbl);


            pheadercolAcrdTbl = new Paragraph("Annualized Fee", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pheadercolAcrdTbl.SetAlignment("center");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheaderAcrdTbl = new PdfPCell();
            celllheaderAcrdTbl.Border = 0;
            celllheaderAcrdTbl.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));
            celllheaderAcrdTbl.AddElement(pheadercolAcrdTbl);
            loAccuredTbl.AddCell(celllheaderAcrdTbl);


            pheadercolAcrdTbl = new Paragraph("", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            pheadercolAcrdTbl.SetAlignment("right");
            //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            celllheaderAcrdTbl = new PdfPCell();
            celllheaderAcrdTbl.Border = 0;
            celllheaderAcrdTbl.BackgroundColor = new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#8094AA"));
            celllheaderAcrdTbl.AddElement(pheadercolAcrdTbl);
            loAccuredTbl.AddCell(celllheaderAcrdTbl);


            #endregion

            loTable.DefaultCell.Padding = 3;






            //  iTextSharp.text.Chunk asterisk = new Chunk();

            #region With Row and Column
            for (int k = 0; k < dtFormat.Rows.Count; k++)
            {

                for (int col = 1; col <= colsize; col++)
                {

                    string DescTxt = Convert.ToString(dtFormat.Rows[k][col]);

                    if (DescTxt == "" && col == 1)
                    {
                        DescTxt = string.Empty;
                        loCell = new PdfPCell();
                        lochunknew = new Chunk("\n", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                        // lochunknew.SetUnderline(1f, 2f);
                        lochunk = new Paragraph(lochunknew);
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        loCell.PaddingBottom = 6f;
                        //    loCell.Padding = 0;
                        lochunk.Leading = 10f;
                        //loCell.Colspan = 4;
                        //loTable.AddCell(loCell);
                        // break;
                    }

                    int BoldFlg = Convert.ToInt32(dtFormat.Rows[k]["_BoldFlg"]);
                    int tabflg = Convert.ToInt32(dtFormat.Rows[k]["_TabFlg"]);

                    int UnderlineFlg = Convert.ToInt32(dtFormat.Rows[k]["_UnderlineFlg"]);


                    //if (BoldFlg == 1 && col == 1)
                    //    lochunk = new Paragraph(DescTxt, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

                    //loCell = new PdfPCell();
                    //loCell.AddElement(lochunk);
                    //loCell.Border = 0;





                    if (BoldFlg == 1 && col == 1 && tabflg == 0 && col == 1)
                    {
                        lochunk = new Paragraph(DescTxt, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        loCell.Padding = 0;
                        //    loCell.Padding = 0;
                        lochunk.Leading = 10f;
                    }

                    if (tabflg == 1 && col == 1 && BoldFlg == 0 && col == 1)
                    {
                        lochunknew = new Chunk(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                        // lochunknew.SetUnderline(1f, 2f);
                        lochunk = new Paragraph(lochunknew);

                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        //  loCell.Padding = 0;
                        loCell.Padding = 0;
                        lochunk.Leading = 10f;
                        // lochunk.Leading = 10f;
                        lochunk.IndentationLeft = 30f;
                    }


                    if (BoldFlg == 1 && col == 1 && tabflg == 1 && col == 1)
                    {
                        lochunknew = new Chunk(DescTxt, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                        // lochunknew.SetUnderline(1f, 2f);
                        lochunk = new Paragraph(lochunknew);

                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        loCell.Padding = 0;

                        lochunk.Leading = 10f;
                        // lochunk.Leading = 10f;
                        lochunk.IndentationLeft = 30f;
                    }


                    if (UnderlineFlg == 1 && col == 1 && tabflg == 1 && col == 1)
                    {
                        lochunk = new Paragraph(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        loCell.BorderWidthBottom = 0.1f;
                        loCell.Padding = 0;
                        loCell.PaddingBottom = 3f;
                        //    loCell.PaddingBottom = 3f;                    //  loCell.PaddingBottom = 2f;
                        lochunk.Leading = 10f;
                        lochunk.IndentationLeft = 30f;
                        // loCell.bo
                    }

                    if (UnderlineFlg == 1 && col == 1 && tabflg == 0 && col == 1)
                    {
                        lochunk = new Paragraph(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                        loCell = new PdfPCell();
                        loCell.AddElement(lochunk);
                        loCell.Border = 0;
                        loCell.BorderWidthBottom = 0.1f;
                        loCell.Padding = 0;
                        loCell.PaddingBottom = 3f;
                        //    loCell.PaddingBottom = 3f;                    //  loCell.PaddingBottom = 2f;
                        lochunk.Leading = 10f;
                        // loCell.bo
                    }

                    if (UnderlineFlg == 0 && BoldFlg == 0 && tabflg == 0 || col != 1)
                    {
                        if (col == 2 || col == 3)
                        {
                            //  string value = DescTxt;

                            DescTxt = currencyFormat(DescTxt).Replace("$", "");



                            //string DescTxt1 = currencyFormat(value);

                        }
                        else if (col == 4)
                        {
                            DescTxt = Percentage(DescTxt);
                        }


                        if (col == 2 || col == 3 || col == 4)
                        {

                            Phrase ph1 = new Phrase();

                            if (BoldFlg == 1)
                            {

                                lochunknew = new Chunk(DescTxt, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

                                if (col == 3)
                                    dollar = new Chunk("    $", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                                else if (col == 2)
                                    dollar = new Chunk("      $", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                            }
                            else
                            {
                                //if (col == 2)
                                //    lochunknew = new Chunk(DescTxt + "        ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                                //else
                                lochunknew = new Chunk(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                                if (col == 3)
                                    dollar = new Chunk("    $", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                                else if (col == 2)
                                    dollar = new Chunk("      $", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                            }

                            if (col == 2 || col == 3)
                            {
                                lochunk = new Paragraph();
                                lochunk.Add(dollar);
                                lochunk.Add(glue);
                                lochunk.Add(lochunknew);
                            }
                            else
                            {

                                lochunk = new Paragraph(lochunknew);
                            }
                            loCell = new PdfPCell();
                            if (DescTxt != "")
                            {
                                if (col == 3 || col == 4)
                                {
                                    if (strDiscount != "" && strDiscount != "0.0")
                                    {
                                        lochunk.Add(asterisk);
                                    }
                                }
                            }

                            else
                            {

                                if (Accrued != "0")
                                {
                                    lochunknew = new Chunk("\n", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                                    lochunk = new Paragraph(lochunknew);
                                }

                                else
                                {
                                    lochunknew = new Chunk(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                                    lochunk = new Paragraph(lochunknew);
                                }

                            }


                            //if (col == 4)
                            //    lochunk.SetAlignment("right");
                            //else
                            lochunk.SetAlignment("right");

                            loCell.Border = 0;
                            loCell.AddElement(lochunk);

                            if (col == 2)
                            {
                                // loCell.PaddingRight = 10f;
                                loCell.Padding = 0;
                                lochunk.Leading = 10f;
                            }
                            else
                            {
                                loCell.Padding = 0;

                                lochunk.Leading = 10f;
                            }


                            if (UnderlineFlg == 1)
                            {
                                loCell.BorderWidthBottom = 0.1f;

                            }
                        }

                        else
                        {


                            // lochunk.Add(glue);

                            if (BoldFlg == 1)
                                lochunknew = new Chunk(DescTxt, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
                            else
                                lochunknew = new Chunk(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

                            lochunk = new Paragraph(lochunknew);
                            loCell = new PdfPCell();
                            loCell.AddElement(lochunk);
                            loCell.Border = 0;
                            loCell.Padding = 0;
                            lochunk.Leading = 10f;
                        }
                    }


                    if (Accrued != "True")
                        loTable.AddCell(loCell);
                    else
                        loAccuredTbl.AddCell(loCell);

                }
            }
            #endregion


            #region with Row only
            //for (int k = 0; k < dtFormat.Rows.Count; k++)
            //{

            //        string DescTxt = Convert.ToString(dtFormat.Rows[k]["DescTxt"]);

            //   // string 


            //        int BoldFlg = Convert.ToInt32(dtFormat.Rows[k]["_BoldFlg"]);
            //        int tabflg = Convert.ToInt32(dtFormat.Rows[k]["_TabFlg"]);

            //        int UnderlineFlg = Convert.ToInt32(dtFormat.Rows[k]["_UnderlineFlg"]);


            //        if (BoldFlg == 1)
            //        {
            //            lochunk = new Paragraph(DescTxt, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //            loCell = new PdfPCell();
            //            loCell.AddElement(lochunk);
            //            loCell.Border = 0;
            //        }

            //        else if (tabflg == 1)
            //        {

            //            lochunk = new Paragraph(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //            lochunk.Leading = 10f;
            //            lochunk.Leading = 10f;
            //            lochunk.IndentationLeft = 30f;
            //            loCell = new PdfPCell();
            //            loCell.AddElement(lochunk);
            //            loCell.Border = 0;
            //        }

            //        else if (UnderlineFlg == 1 )
            //        {
            //            lochunk = new Paragraph(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //            loCell = new PdfPCell();
            //            loCell.AddElement(lochunk);
            //            loCell.Border = 0;
            //        }

            //        else

            //        {
            //            lochunk = new Paragraph(DescTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //            loCell = new PdfPCell();
            //            loCell.AddElement(lochunk);
            //            loCell.Border = 0;
            //        }




            //        //  loCell.Leading=
            //        //   loCell.BackgroundColor = iTextSharp.text.Color.LIGHT_GRAY;

            //        loTable.AddCell(loCell);
            //    }


            #endregion

            Paragraph p10noteleft = new Paragraph(" ", setFontsAll(9, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#ffffff"))));
            //  Chunk c10noteright = new Chunk("Annual, quarterly fee, and fee rate reflect a " + string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:P2}", Convert.ToDouble(strDiscount)) + " Discount", setFontsAll(8, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Chunk c10noteright = new Chunk("Assets under advisement fee reflects a " + string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:P2}", Convert.ToDouble(strDiscount)) + " Discount", setFontsAll(8, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Paragraph p10noteright = new Paragraph();

            p10noteright.SetAlignment("right");
            p10noteleft.Leading = 12f;
            p10noteleft.Leading = 12f;
            p10noteright.Add(asterisk);
            p10noteright.Add(c10noteright);




            /**** 11th Line ***/
            Paragraph p11 = new Paragraph("Note:  Numbers are rounded", setFontsAll(9, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //   p11.Leading = 12f;

            /*** 12th Line ***/
            Paragraph p12 = new Paragraph();
            Chunk p12text = new Chunk("If you have established automatic fee payment through your account with Fidelity please do not send a check.  Please compare the Quarterly Fee Due to the amount debited from your Fidelity Statement on      ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            Chunk p12date = new Chunk(strAutoDebitDt + ".", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            p12.Add(p12text);
            p12.Add(p12date);
            p12.Leading = 12f;

            /*** 13th Line ***/
            Paragraph p13 = new Paragraph("If you pay by check please make your check payable to �Gresham Partners, LLC� and return along with a copy of your invoice. ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            p13.Leading = 12f;
            /*** 14th Line ***/
            Paragraph p14 = new Paragraph("If you would like a detailed list of your assets under advisement or have questions, please contact your Advisor or David Salsburg at (312) 960-0221. ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            p14.Leading = 12f;

            Paragraph pT12 = new Paragraph();
            Chunk p12text1 = new Chunk("Your quarterly fees will accrue until initial investment at which time payment of the total accrual to date will be due. ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
            //Chunk p12date1 = new Chunk(strAutoDebitDt + ".", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

            pT12.Add(p12text1);

            pT12.Leading = 12f;

            cellHeaderleft.AddElement(pLetterDate);
            cellHeaderright.AddElement(pInvoice);

            cellHeaderleft.Colspan = 2;
            cellHeaderright.Colspan = 2;

            cellAddress.AddElement(pLetterDate);
            cellAddress.AddElement(PBlank);
            cellAddress.AddElement(PBlank);


            cellAddress.AddElement(pContactFullName);
            cellAddress.AddElement(pAddress1);
            if (strAddress2 != "")
                cellAddress.AddElement(pAddress2);

            if (strAddress3 != "")
                cellAddress.AddElement(pAddress3);
            cellAddress.AddElement(pCityState);

            if (!string.IsNullOrEmpty(strCountry))
                cellAddress.AddElement(pCountry);

            cellAddress.Colspan = 2;

            cellRef.AddElement(pREF);
            cellRef.Colspan = 2;

            cellline4.AddElement(p4);
            cellline4.Colspan = 2;

            cellAsOfDate.AddElement(pAsofdate);
            cellAsOfDate.Colspan = 2;

            celllineBlank.AddElement(PBlank);

            celllineBlank.Colspan = 2;

            if (BillAmt != 0)
            {
                if (Accrued != "True")
                {
                    //cellline6left_1.AddElement(p6left_1);
                    //cellline6left_2.AddElement(p6left_2);

                    //cellline6right_1.AddElement(p6right_1);
                    //cellline6right_2.AddElement(p6right_2);

                }
                else
                {
                    Paragraph leftcell1 = new Paragraph();
                    Paragraph rightcell1 = new Paragraph();
                    leftcell1.Leading = 12f;
                    rightcell1.Leading = 12f;
                    celllineBlank1.AddElement(leftcell1);
                    celllineBlankright1.AddElement(rightcell1);
                }

            }
            else
            {
                Paragraph leftcell1 = new Paragraph();
                Paragraph rightcell1 = new Paragraph();
                leftcell1.Leading = 12f;
                rightcell1.Leading = 12f;
                celllineBlank1.AddElement(leftcell1);
                celllineBlankright1.AddElement(rightcell1);
                //cellAsOfDate.PaddingBottom = 12f;
            }


            if (Accrued != "True")
            {
                //cellline7left.AddElement(p7left);
                //cellline7right.AddElement(p7right);
            }
            else
            {
                //cellline7left.AddElement(p7left);
                //cellline7right.AddElement(p7right);
            }


            cellline10noteleft.AddElement(p10noteleft);
            cellline10noteright.AddElement(p10noteright);

            cellline11.AddElement(p11);
            cellline11.Colspan = 2;

            if (Accrued != "True")
            {
                cellline12.AddElement(p12);
                cellline12.Colspan = 2;

                cellline13.AddElement(p13);
                cellline13.Colspan = 2;

                cellline14.AddElement(p14);
                cellline14.Colspan = 2;
            }
            else
            {
                celllineT12.AddElement(pT12);
                celllineT12.Colspan = 2;
            }

            cellline10noteright.Colspan = 2;



            // tblMain.AddCell(cellHeaderleft);
            // tblMain.AddCell(new PdfPCell());
            tblMain.AddCell(cellHeaderright);
            //tblMain.AddCell(cellHeaderleft);
            tblMain.AddCell(cellAddress);
            tblMain.AddCell(cellRef);
            tblMain.AddCell(cellline4);

            tblMain.AddCell(cellAsOfDate);


            tblMain.AddCell(celllineBlank);

            PdfPCell cellloheader = null;
            if (Accrued != "True")
            {
                cellloheader = new PdfPCell(loTable);
                cellloheader.Colspan = 2;
            }
            else
            {
                cellloheader = new PdfPCell(loAccuredTbl);
                cellloheader.Border = 0;
                cellloheader.Colspan = 2;
                //  cellloheader.Width = 60F;
            }


            cellloheader.Border = 0;




            tblMain.AddCell(cellloheader);

            if (BillAmt != 0)
            {
                if (Accrued != "True")
                {

                }
                else
                {
                    tblMain.AddCell(celllineBlank1);
                    tblMain.AddCell(celllineBlankright1);
                }

            }
            else
            {
                tblMain.AddCell(celllineBlank1);
                tblMain.AddCell(celllineBlankright1);

            }

            if (strDiscount != "" && strDiscount != "0.0")
                tblMain.AddCell(cellline10noteright);


            //if (strDiscount != "" && strDiscount != "0.0")
            //    tblMain.AddCell(cellline10noteright);

            tblMain.AddCell(cellline11);


            if (Accrued != "True")
            {
                tblMain.AddCell(cellline12);
                tblMain.AddCell(cellline13);
                tblMain.AddCell(cellline14);
            }
            else
            {
                tblMain.AddCell(celllineT12);
            }
            pdoc.Add(tblMain);

            pdoc.Close();

            if (DSCount > 0)
            {
                try
                {
                    FileInfo loFile = new FileInfo(ls);
                    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                }
                catch
                { }

            }
            else
            {
                fsFinalLocation = "";
            }

            return fsFinalLocation.Replace(".xls", ".pdf");
        }
        catch
        {
            return "";
        }
    }

    #region Old Function not in used
    //public string GetBillingInvoice()
    //{
    //    DB clsDB = new DB();
    //    String lsSQL = getFinalSp(ReportType.Invoice);//Store Procedure call


    //    DataSet newdataset = clsDB.getDataSet(lsSQL);
    //    string strBillableAUM = string.Empty;

    //    string underline = string.Empty;
    //    string underline1 = string.Empty;

    //    int DSC = newdataset.Tables.Count;
    //    int DSCount = newdataset.Tables[0].Rows.Count;
    //    // string strGUID = System.DateTime.Now.ToString("MMddyyhhmmssfff");
    //    // String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + "_Billing.pdf";

    //    var strGUID = Guid.NewGuid().ToString();

    //    String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + strGUID + "_Billing.pdf";

    //    String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + System.DateTime.Now.ToString("MMddyyhhmmssfff") + System.Guid.NewGuid().ToString() + "_Billing.pdf";

    //    iTextSharp.text.Document pdoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 48, 48, 31, 8);//10,10        
    //    PdfWriter writer = PdfWriter.GetInstance(pdoc, new FileStream(ls, FileMode.Create));
    //    //AddFooter(pdoc);
    //    //Footer
    //    Phrase footPhraseImg = new Phrase("Gresham Partners, LLC | 333 W. Wacker Dr. Suite 700 | Chicago, IL 60606 | P 312.960.0200 | F 312.960.0204 | www.greshampartners.com", setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
    //    HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
    //    footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
    //    footer.Alignment = Element.ALIGN_LEFT;
    //    pdoc.Footer = footer;

    //    pdoc.Open();

    //    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
    //    png.SetAbsolutePosition(48, 800);//540
    //    //png.ScaleToFit(288f, 42f);
    //    png.ScalePercent(10);
    //    pdoc.Add(png);

    //    string strLetterDate = Convert.ToString(newdataset.Tables[0].Rows[0]["LetterDate"]); //"November 12, 2015";
    //    if (strLetterDate != "")
    //        strLetterDate = Convert.ToDateTime(strLetterDate).ToString("MMMM dd,yyyy");

    //    string strInvYear = Convert.ToString(newdataset.Tables[0].Rows[0]["LetterDate"]);
    //    if (strInvYear != "")
    //        strInvYear = Convert.ToDateTime(strLetterDate).ToString("yyyy");

    //    string strInvNumber = Convert.ToString(newdataset.Tables[0].Rows[0]["InvoiceNumber"]); //"1111";
    //    string strContactFullname = Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_salutation_mail"]) + " " + Convert.ToString(newdataset.Tables[0].Rows[0]["ContactFullName"]);
    //    string strAddress1 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline1_mail"]);
    //    string strAddress2 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline2_mail"]);
    //    string strAddress3 = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_addressline3_mail"]);
    //    string strCity = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_city_mail"]);
    //    string strState = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_stateprovince_mail"]);
    //    string strZip = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_zipcode_mail"]);
    //    string strCountry = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_countryregion_mail"]);
    //    string strHouseHold = Convert.ToString(newdataset.Tables[0].Rows[0]["HouseHoldName"]); //"Adams Family";
    //    string strRefName = Convert.ToString(newdataset.Tables[0].Rows[0]["RefName"]); //"Adams Family";
    //    string strAsOfDate = Convert.ToString(newdataset.Tables[0].Rows[0]["DateRange"]); //"November 2015 � January 2016";

    //    string strAODmailmerge = Convert.ToString(newdataset.Tables[0].Rows[0]["AumAsOfDate"]);  //"September 30, 2015";
    //    if (strAODmailmerge != "")
    //        strAODmailmerge = Convert.ToDateTime(strAODmailmerge).ToString("MMMM dd,yyyy");

    //    strBillableAUM = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["BillableAumAmt"])); //"$ 34,262,588";


    //    string strBillableAUM1 = strBillableAUM.Replace("$", "").Replace(" ", "");


    //    string strFeeRate = Percentage(Convert.ToString(newdataset.Tables[0].Rows[0]["FeeRatePct"]));//"0.68%";
    //    string strAnnualFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["AnnualisedFeeAmt"]));//"$ 233,813";
    //    string strQuaterFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])); //"$58,453";

    //    //string strQuaterFee = RoundUp(Convert.ToString(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])); //"$58,453";


    //    //decimal strQuaterFee1 = Decimal.Round((Convert.ToDecimal(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])), 2, MidpointRounding.AwayFromZero);

    //    //string strQuaterFee2 = Convert.ToString(strQuaterFee1);

    //    // string strQuaterFee1 =((Convert.ToString(newdataset.Tables[0].Rows[0]["QuarterlyFeeAmt"])), 2, MidpointRounding.AwayFromZero);

    //    string strAdjustment = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Adjustment"]));
    //    string strAdjustedFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_AdjustedFee"]));
    //    string strAdjustmentReason = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_AdjustmentReason"]);


    //    //string strRelationshipFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_RelationshipFee"]));


    //    //string strFeeName = (Convert.ToString(newdataset.Tables[2].Rows[0]["ssi_name"]));


    //    //string strAmount = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[2].Rows[0]["ssi_Amount"]));
    //    //string strAmount1 = strAmount.Replace("$", "").Replace(" ", "");


    //    //string strSetupFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_SetUpFee"]));
    //    //string strSetupFee1 = strSetupFee.Replace("$", "").Replace(" ", "");


    //    string strAnnFeeBeforeAdditionalFee = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["AnnFeeBeforeAdditionalFee"]));

    //    string strAnnFeeBeforeAdditionalFee1 = strAnnFeeBeforeAdditionalFee.Replace("$", "").Replace(" ", "");

    //    string Accrued = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Accrued"]);

    //    string Monthly1 = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["Ssi_Month1Fee"]));

    //    string strDiscount = "0.0";

    //    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Discount"]) != "")
    //        strDiscount = Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Discount"]);


    //    string strAutoDebitDt = Convert.ToString(newdataset.Tables[0].Rows[0]["AutoDebitDt"]); //"November 19, 2015";

    //    if (strAutoDebitDt != "")
    //        strAutoDebitDt = Convert.ToDateTime(strAutoDebitDt).ToString("MMMM d, yyyy");

    //    PdfPTable tblMain = new PdfPTable(2);

    //    PdfPCell cellHeaderleft = new PdfPCell();
    //    PdfPCell cellHeaderright = new PdfPCell();
    //    PdfPCell cellAddress = new PdfPCell();
    //    PdfPCell cellRef = new PdfPCell();
    //    PdfPCell cellline4 = new PdfPCell();
    //    PdfPCell cellAsOfDate = new PdfPCell();
    //    PdfPCell cellline6left = new PdfPCell();
    //    PdfPCell cellline6right = new PdfPCell();

    //    PdfPCell cellline6left1 = new PdfPCell();
    //    PdfPCell cellline6right1 = new PdfPCell();


    //    PdfPCell celllineT12 = new PdfPCell();

    //    //----------test--------------------------------------

    //    PdfPCell celllineTestleft = new PdfPCell();
    //    PdfPCell celllineTest1left = new PdfPCell();
    //    PdfPCell celllineTest2left = new PdfPCell();
    //    PdfPCell celllineTest3left = new PdfPCell();
    //    PdfPCell celllineTest4left = new PdfPCell();

    //    PdfPCell celllineTest5left = new PdfPCell();

    //    PdfPCell celllineTest6left = new PdfPCell();

    //    PdfPCell celllineTest7left = new PdfPCell();

    //    PdfPCell celllineTest8left = new PdfPCell();

    //    PdfPCell celllineTest9left = new PdfPCell();

    //    PdfPCell celllineTest10left = new PdfPCell();

    //    PdfPCell celllineTest11left = new PdfPCell();

    //    PdfPCell celllineTest12left = new PdfPCell();

    //    PdfPCell celllineTest13left = new PdfPCell();

    //    PdfPCell celllineTest14left = new PdfPCell();

    //    PdfPCell celllineTest15left = new PdfPCell();

    //    PdfPCell celllineTest16left = new PdfPCell();

    //    PdfPCell celllineTest17left = new PdfPCell();

    //    PdfPCell celllineTest18left = new PdfPCell();

    //    PdfPCell celllineTest19left = new PdfPCell();

    //    PdfPCell celllineTest20left = new PdfPCell();

    //    PdfPCell celllineTest21left = new PdfPCell();

    //    PdfPCell celllineTest22left = new PdfPCell();

    //    PdfPCell celllineTestright = new PdfPCell();
    //    PdfPCell celllineTest1right = new PdfPCell();
    //    PdfPCell celllineTest2right = new PdfPCell();
    //    PdfPCell celllineTest3right = new PdfPCell();
    //    PdfPCell celllineTest4right = new PdfPCell();
    //    PdfPCell celllineTest5right = new PdfPCell();
    //    PdfPCell celllineTest6right = new PdfPCell();
    //    PdfPCell celllineTest7right = new PdfPCell();
    //    PdfPCell celllineTest8right = new PdfPCell();

    //    PdfPCell celllineTest9right = new PdfPCell();

    //    PdfPCell celllineTest10right = new PdfPCell();

    //    PdfPCell celllineTest11right = new PdfPCell();

    //    PdfPCell celllineTest12right = new PdfPCell();

    //    PdfPCell celllineTest13right = new PdfPCell();

    //    PdfPCell celllineTest14right = new PdfPCell();

    //    PdfPCell celllineTest15right = new PdfPCell();

    //    PdfPCell celllineTest16right = new PdfPCell();

    //    PdfPCell celllineTest17right = new PdfPCell();

    //    PdfPCell celllineTest18right = new PdfPCell();

    //    PdfPCell celllineTest19right = new PdfPCell();

    //    PdfPCell celllineTest20right = new PdfPCell();

    //    PdfPCell celllineTest21right = new PdfPCell();

    //    PdfPCell celllineTest22right = new PdfPCell();


    //    //----------test--------------------------------------

    //    PdfPCell cellline7left = new PdfPCell();
    //    PdfPCell cellline7right = new PdfPCell();
    //    PdfPCell cellline8left = new PdfPCell();
    //    PdfPCell cellline8right = new PdfPCell();
    //    PdfPCell cellline9 = new PdfPCell(); //Border above -- blank

    //    PdfPCell celllineBlank = new PdfPCell(); //Border above -- blank
    //    PdfPCell celllineBlankright = new PdfPCell();

    //    PdfPCell celllineBlank1 = new PdfPCell(); //Border above -- blank
    //    PdfPCell celllineBlankright1 = new PdfPCell();

    //    PdfPCell cellline10left = new PdfPCell();
    //    PdfPCell cellline10right = new PdfPCell();
    //    PdfPCell cellline10noteleft = new PdfPCell();
    //    PdfPCell cellline10noteright = new PdfPCell();

    //    //**
    //    //New Section Added for adjustment 
    //    //**
    //    PdfPCell cellline1AdjustmentLeft = new PdfPCell();
    //    PdfPCell cellline1AdjustmentRight = new PdfPCell();

    //    PdfPCell cellline2AdjustmentLeft = new PdfPCell();
    //    PdfPCell cellline2AdjustmentRight = new PdfPCell();
    //    //**

    //    PdfPCell cellline11 = new PdfPCell();
    //    PdfPCell cellline12 = new PdfPCell();
    //    PdfPCell cellline13 = new PdfPCell();
    //    PdfPCell cellline14 = new PdfPCell();


    //    cellHeaderleft.Border = 0;
    //    cellHeaderright.Border = 0;
    //    cellAddress.Border = 0;
    //    cellRef.Border = 0;
    //    cellline4.Border = 0;
    //    cellAsOfDate.Border = 0;
    //    cellline6left.Border = 0;
    //    cellline6right.Border = 0;

    //    cellline6left1.Border = 0;
    //    cellline6right1.Border = 0;

    //    celllineT12.Border = 0;

    //    //---------------------------------------------test--------------------------------------------------
    //    celllineTestleft.Border = 0;
    //    celllineTestright.Border = 0;

    //    celllineTest1left.Border = 0;
    //    celllineTest1right.Border = 0;

    //    celllineTest2left.Border = 0;
    //    celllineTest2right.Border = 0;

    //    celllineTest3left.Border = 0;
    //    celllineTest3right.Border = 0;

    //    celllineTest4left.Border = 0;
    //    celllineTest4right.Border = 0;

    //    celllineTest5left.Border = 0;
    //    celllineTest5right.Border = 0;

    //    celllineTest6left.Border = 0;
    //    celllineTest6right.Border = 0;

    //    celllineTest7left.Border = 0;
    //    celllineTest7right.Border = 0;

    //    celllineTest8left.Border = 0;
    //    celllineTest8right.Border = 0;

    //    celllineTest9left.Border = 0;
    //    celllineTest9right.Border = 0;

    //    celllineTest10left.Border = 0;
    //    celllineTest10right.Border = 0;

    //    celllineTest11left.Border = 0;
    //    celllineTest11right.Border = 0;

    //    celllineTest12left.Border = 0;
    //    celllineTest12right.Border = 0;

    //    celllineTest13left.Border = 0;
    //    celllineTest13right.Border = 0;

    //    celllineTest15left.Border = 0;
    //    celllineTest15right.Border = 0;


    //    celllineTest16left.Border = 0;
    //    celllineTest16right.Border = 0;

    //    celllineTest17left.Border = 0;
    //    celllineTest17right.Border = 0;

    //    celllineTest18left.Border = 0;
    //    celllineTest18right.Border = 0;

    //    celllineTest19left.Border = 0;
    //    celllineTest19right.Border = 0;

    //    celllineTest20left.Border = 0;
    //    celllineTest20right.Border = 0;

    //    celllineTest21left.Border = 0;
    //    celllineTest21right.Border = 0;

    //    celllineTest22left.Border = 0;
    //    celllineTest22right.Border = 0;



    //    //celllineTest5left.Border = 0;
    //    //celllineTest5right.Border = 0;

    //    //celllineTest4left.Border = 0;
    //    //celllineTest4right.Border = 0;

    //    //---------------------------------------------test--------------------------------------------------

    //    cellline7left.Border = 0;
    //    cellline7right.Border = 0;
    //    cellline8left.Border = 0;
    //    cellline8right.Border = 0;
    //    cellline9.Border = 0;

    //    celllineBlank.Border = 0;
    //    celllineBlankright.Border = 0;

    //    celllineBlank1.Border = 0;
    //    celllineBlankright1.Border = 0;

    //    cellline10left.Border = 0;
    //    cellline10right.Border = 0;

    //    cellline10noteleft.Border = 0;
    //    cellline10noteright.Border = 0;
    //    cellline1AdjustmentLeft.Border = 0;
    //    cellline1AdjustmentRight.Border = 0;
    //    cellline2AdjustmentLeft.Border = 0;
    //    cellline2AdjustmentRight.Border = 0;
    //    cellline11.Border = 0;
    //    cellline12.Border = 0;
    //    cellline13.Border = 0;
    //    cellline14.Border = 0;

    //    cellHeaderleft.PaddingTop = 35f;
    //    cellHeaderright.PaddingTop = 35f;

    //    cellAddress.PaddingTop = 60f;

    //    cellRef.PaddingTop = 60f;

    //    cellline4.PaddingTop = 36f;

    //    cellAsOfDate.PaddingTop = 10f;

    //    cellline6left.PaddingTop = 10f;
    //    cellline6right.PaddingTop = 10f;


    //    cellline6left1.PaddingTop = 10f;
    //    cellline6right1.PaddingTop = 10f;

    //    cellline8left.PaddingBottom = 10f;
    //    cellline8right.PaddingBottom = 10f;

    //    cellline11.PaddingTop = 10f;
    //    cellline12.PaddingTop = 10f;

    //    celllineT12.PaddingTop = 10f;

    //    cellline13.PaddingTop = 10f;
    //    cellline14.PaddingTop = 10f;

    //    int[] widthHeader = { 80, 20 };
    //    tblMain.SetWidths(widthHeader);
    //    tblMain.TotalWidth = 100f;
    //    tblMain.WidthPercentage = 100f;

    //    Chunk asterisk = new Chunk("*", setFontsAll(6, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    asterisk.SetTextRise(4);


    //    /***    First Line  ***/


    //    Paragraph pLetterDate = new Paragraph(strLetterDate.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph pInvoice = new Paragraph("Invoice " + strInvYear.ToString() + "-" + strInvNumber.ToString(), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    pInvoice.SetAlignment("right");
    //    pLetterDate.Leading = 12f;
    //    pInvoice.Leading = 12f;

    //    /*** Address ***/
    //    Paragraph pContactFullName = new Paragraph(strContactFullname.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph pAddress1 = new Paragraph(strAddress1.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph pAddress2 = new Paragraph(strAddress2.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph pAddress3 = new Paragraph(strAddress3.ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph pCityState = new Paragraph();
    //    Chunk pCity = new Chunk(strCity + ", ".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Chunk pState = new Chunk(strState + " ".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Chunk pZipCode = new Chunk(strZip + " ".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph pCountry = new Paragraph(strCountry + "".ToString().Replace("''", "'"), setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    if (!string.IsNullOrEmpty(strCity))
    //        pCityState.Add(pCity);

    //    if (!string.IsNullOrEmpty(strState))
    //        pCityState.Add(pState);

    //    if (!string.IsNullOrEmpty(strZip))
    //        pCityState.Add(pZipCode);

    //    //if (!string.IsNullOrEmpty(strCountry))
    //    //    pCityState.Add(pCountry);

    //    pContactFullName.Leading = 12f;
    //    pAddress1.Leading = 12f;
    //    pAddress2.Leading = 12f;
    //    pAddress3.Leading = 12f;
    //    pCityState.Leading = 12f;
    //    pCountry.Leading = 12f;
    //    /*** REF ***/
    //    Paragraph pREF = new Paragraph("REF: " + strRefName, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    pREF.Leading = 12f;

    //    /*** 4th Line ***/
    //    Paragraph p4 = new Paragraph("Invoice for financial counseling and investment services for the period: ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    p4.Leading = 12f;

    //    /*** As of Date ***/
    //    Paragraph pAsofdate = new Paragraph(strAsOfDate, setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    pAsofdate.SetAlignment("center");
    //    pAsofdate.Leading = 12f;
    //    /*** 6th Line ***/
    //    // Paragraph p6 = new Paragraph();

    //    strBillableAUM = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[0].Rows[0]["BillableAumAmt"]));

    //    decimal BillAmt = Convert.ToDecimal(strBillableAUM1);

    //    Paragraph p6left1 = null;
    //    Paragraph p6right1 = null;

    //    // p6left1 = new Paragraph("Agreed upon flat monthly fee to be accrued until intial investment " + strAODmailmerge.ToString().Replace("''", "'") + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    // p6right1 = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    //p6right1.SetAlignment("right");
    //    //p6left1.Leading = 10f;
    //    //p6right1.Leading = 10f;

    //    Paragraph p6left = null;
    //    Paragraph p6right = null;

    //    p6left = new Paragraph("Assets under advisement " + strAODmailmerge.ToString().Replace("''", "'") + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    p6right = new Paragraph(strBillableAUM, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    p6right.SetAlignment("right");
    //    p6left.Leading = 10f;
    //    p6right.Leading = 10f;





    //    // p6.Add(p6left);
    //    // p6.Add(p6right);

    //    //-------------------------------------------------test-------------------------------------------------------------------//
    //    Paragraph pTest3left = null;
    //    Paragraph pTest3right = null;


    //    Paragraph pTestleft = null;
    //    Paragraph pTestright = null;

    //    //if (Convert.ToDecimal(strRelationshipFee1) != 0)
    //    //{
    //    //    // Paragraph pTest = new Paragraph();
    //    //    pTestleft = new Paragraph("Relationship Fee " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    //    pTestright = new Paragraph(strRelationshipFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    //    pTestright.SetAlignment("right");
    //    //    pTestleft.Leading = 10f;
    //    //    pTestright.Leading = 10f;
    //    //    pTestleft.IndentationLeft = 30f;

    //    //} // pTest.Add(pTestleft);
    //    //// pTest.Add(pTestright);


    //    Paragraph pTest1left = null;
    //    Paragraph pTest1right = null;

    //    //if (Convert.ToDecimal(strSetupFee1) != 0)
    //    //{

    //    //    // Paragraph pTest1 = new Paragraph();
    //    //    pTest1left = new Paragraph("Setup Fee " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    //    pTest1right = new Paragraph(strSetupFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    //    pTest1right.SetAlignment("right");
    //    //    pTest1left.Leading = 10f;
    //    //    pTest1right.Leading = 10f;
    //    //    pTest1left.IndentationLeft = 30f;
    //    //}
    //    // pTest1.Add(pTest1left);
    //    // pTest1.Add(pTest1right);
    //    Paragraph pTest2left = null;
    //    Paragraph pTest2right = null;

    //    Paragraph pTest4left = null;
    //    Paragraph pTest4right = null;

    //    Paragraph pTest5left = null;
    //    Paragraph pTest5right = null;

    //    Paragraph pTest6left = null;
    //    Paragraph pTest6right = null;


    //    Paragraph pTest7left = null;
    //    Paragraph pTest7right = null;


    //    Paragraph pTest8left = null;
    //    Paragraph pTest8right = null;


    //    Paragraph pTest9left = null;
    //    Paragraph pTest9right = null;

    //    Paragraph pTest10left = null;
    //    Paragraph pTest10right = null;

    //    Paragraph pTest11left = null;
    //    Paragraph pTest11right = null;

    //    Paragraph pTest12left = null;
    //    Paragraph pTest12right = null;

    //    Paragraph pTest13left = null;
    //    Paragraph pTest13right = null;

    //    Paragraph pTest14left = null;
    //    Paragraph pTest14right = null;


    //    Paragraph pTest15left = null;
    //    Paragraph pTest15right = null;

    //    Paragraph pTest16left = null;
    //    Paragraph pTest16right = null;

    //    Paragraph pTest17left = null;
    //    Paragraph pTest17right = null;


    //    Paragraph pTest18left = null;
    //    Paragraph pTest18right = null;

    //    Paragraph pTest19left = null;
    //    Paragraph pTest19right = null;

    //    Paragraph pTest20left = null;
    //    Paragraph pTest20right = null;

    //    Paragraph pTest21left = null;
    //    Paragraph pTest21right = null;

    //    Paragraph pTest22left = null;
    //    Paragraph pTest22right = null;


    //    if (newdataset.Tables[1].Rows.Count > 0)
    //    {

    //        decimal Amount = 0;
    //        //int count = 0;

    //        for (int i = 0; i < newdataset.Tables[1].Rows.Count; i++)
    //        {
    //            //  Paragraph pTest1 = new Paragraph();
    //            //pTest2left = new Paragraph("Fidelity Fee abcd " + strAODmailmerge.ToString().Replace("''", "'") + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //            //pTest2right = new Paragraph("0", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //            //pTest2right.SetAlignment("right");
    //            //pTest2left.Leading = 10f;
    //            //pTest2right.Leading = 10f;

    //            string Fidility = Convert.ToString(newdataset.Tables[1].Rows[i]["ssi_Name"]);

    //            // string FidilityAmt = RoundToZeroDecimal(currencyFormat(Convert.ToString(newdataset.Tables[1].Rows[i]["ssi_Amount"])));

    //            string FidilityAmt = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[1].Rows[i]["ssi_Amount"]));


    //            string FidilityAmt1 = FidilityAmt.Replace("$", "").Replace(" ", "").Replace("(", "").Replace(")", "");
    //            underline = Convert.ToString((newdataset.Tables[1].Rows[i]["_UnderlineFlg"]));


    //            // string Amount = string.Empty;

    //            if (FidilityAmt.Contains("-"))
    //            {
    //                FidilityAmt = currencyFormat(FidilityAmt);




    //                // FidilityAmt = Amount.ToString(); ;
    //            }

    //            //decimal.TryParse(FidilityAmt.Trim(), out Amount);

    //            //    Amount = Convert.ToDecimal(FidilityAmt.Trim());

    //            if (FidilityAmt != "" && FidilityAmt != " ")
    //            {
    //                if (Convert.ToDecimal(FidilityAmt1) != 0)
    //                {
    //                    if (i == 0)
    //                    {
    //                        // string fiedility = Convert.ToString(newdataset.Tables[1].
    //                        pTest15left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        // pTest15right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk15 = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest15right = new Paragraph();


    //                        if (underline == "1")
    //                        {
    //                            chk15.SetUnderline(1f, -2f);
    //                        }

    //                        pTest15right.SetAlignment("right");
    //                        pTest15left.Leading = 10f;
    //                        pTest15right.Leading = 10f;
    //                        pTest15left.IndentationLeft = 30f;
    //                        pTest15right.Add(chk15);
    //                    }

    //                    if (i == 1)
    //                    {
    //                        pTest16left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest16right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        Chunk chk16 = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest16right = new Paragraph();


    //                        if (underline == "1")
    //                        {
    //                            chk16.SetUnderline(1f, -2f);
    //                        }


    //                        pTest16right.SetAlignment("right");
    //                        pTest16left.Leading = 10f;
    //                        pTest16right.Leading = 10f;
    //                        pTest16left.IndentationLeft = 30f;
    //                        pTest16right.Add(chk16);
    //                    }
    //                    if (i == 2)
    //                    {
    //                        pTest17left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        // pTest17right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk17 = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest17right = new Paragraph();


    //                        if (underline == "1")
    //                        {
    //                            chk17.SetUnderline(1f, -2f);
    //                        }

    //                        pTest17right.SetAlignment("right");
    //                        pTest17left.Leading = 10f;
    //                        pTest17right.Leading = 10f;
    //                        pTest17left.IndentationLeft = 30f;
    //                        pTest17right.Add(chk17);
    //                    }
    //                    if (i == 3)
    //                    {
    //                        pTest18left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest18right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk18 = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest18right = new Paragraph();


    //                        if (underline == "1")
    //                        {
    //                            chk18.SetUnderline(1f, -2f);
    //                        }


    //                        pTest18right.SetAlignment("right");
    //                        pTest18left.Leading = 10f;
    //                        pTest18right.Leading = 10f;
    //                        pTest18left.IndentationLeft = 30f;
    //                        pTest18right.Add(chk18);
    //                    }

    //                    if (i == 4)
    //                    {
    //                        pTest19left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        // pTest19right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk19 = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest19right = new Paragraph();


    //                        if (underline == "1")
    //                        {
    //                            chk19.SetUnderline(1f, -2f);
    //                        }

    //                        pTest19right.SetAlignment("right");
    //                        pTest19left.Leading = 10f;
    //                        pTest19right.Leading = 10f;
    //                        pTest19left.IndentationLeft = 30f;
    //                        pTest19right.Add(chk19);
    //                    }

    //                    if (i == 5)
    //                    {
    //                        pTest20left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest20right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk20 = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest20right = new Paragraph();


    //                        if (underline == "1")
    //                        {
    //                            chk20.SetUnderline(1f, -2f);
    //                        }

    //                        pTest20right.SetAlignment("right");
    //                        pTest20left.Leading = 10f;
    //                        pTest20right.Leading = 10f;
    //                        pTest20left.IndentationLeft = 30f;
    //                        pTest20right.Add(chk20);



    //                    }

    //                    if (i == 6)
    //                    {
    //                        pTest21left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        // pTest21right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk21 = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest20right = new Paragraph();


    //                        if (underline == "1")
    //                        {
    //                            chk21.SetUnderline(1f, -2f);
    //                        }

    //                        pTest21right.SetAlignment("right");
    //                        pTest21left.Leading = 10f;
    //                        pTest21right.Leading = 10f;
    //                        pTest21left.IndentationLeft = 30f;
    //                        pTest21right.Add(chk21);
    //                    }

    //                }

    //            }

    //        }




    //    }


    //    ///************************************************FLAT FEE LOGIC ********************************************************///
    //    if (DSC > 1)
    //    {
    //        int count = newdataset.Tables[0].Rows.Count;

    //        if (Convert.ToDecimal(strAnnFeeBeforeAdditionalFee1) != 0)
    //        {
    //            if (newdataset.Tables[2].Rows.Count > 0)
    //            {


    //                //int count = 0;

    //                for (int i = 0; i < newdataset.Tables[2].Rows.Count; i++)
    //                {
    //                    //  Paragraph pTest1 = new Paragraph();
    //                    //pTest2left = new Paragraph("Fidelity Fee abcd " + strAODmailmerge.ToString().Replace("''", "'") + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                    //pTest2right = new Paragraph("0", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                    //pTest2right.SetAlignment("right");
    //                    //pTest2left.Leading = 10f;
    //                    //pTest2right.Leading = 10f;

    //                    string Fidility = Convert.ToString(newdataset.Tables[2].Rows[i]["ssi_Name"]);

    //                    string FidilityAmt = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[2].Rows[i]["ssi_Amount"]));

    //                    string FidilityAmt1 = FidilityAmt.Replace("$", "").Replace(" ", "").Replace("(", "").Replace(")", "");

    //                    underline1 = Convert.ToString((newdataset.Tables[2].Rows[i]["_UnderlineFlg"]));

    //                    // string FidilityAmt1 = FidilityAmt.Replace("$", "").Replace(" ", "");

    //                    //decimal FlatFee;
    //                    //if (decimal.TryParse(FidilityAmt1.Trim(), out FlatFee))
    //                    //    FlatFee = Convert.ToDecimal(FidilityAmt1.Trim());
    //                    //else
    //                    //    FlatFee = 0;
    //                    //string FidilityAmt3 = (RoundToZeroDecimal(Convert.ToString(newdataset.Tables[2].Rows[i]["ssi_Amount"])));
    //                    if (FidilityAmt.Contains("-"))
    //                    {
    //                        FidilityAmt = currencyFormat(Convert.ToString(newdataset.Tables[2].Rows[i]["ssi_Amount"]));

    //                    }

    //                    if (Convert.ToDecimal(FidilityAmt1) != 0 || count != 0)
    //                    {

    //                        ////decimal AnnFeeBeforeAdditional;
    //                        ////if (decimal.TryParse(strAnnFeeBeforeAdditionalFee1.Trim(), out FlatFee))
    //                        ////    AnnFeeBeforeAdditional = Convert.ToDecimal(strAnnFeeBeforeAdditionalFee1.Trim());
    //                        ////else
    //                        //    AnnFeeBeforeAdditional = 0;

    //                        if (strAnnFeeBeforeAdditionalFee1 != "")
    //                        {
    //                            if (Convert.ToDecimal(strAnnFeeBeforeAdditionalFee1) != 0)
    //                            {
    //                                pTest3left = new Paragraph("Assets under advisement Fee for the Period " + strAsOfDate + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                                pTest3right = new Paragraph(strAnnFeeBeforeAdditionalFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                                pTest3right.SetAlignment("right");
    //                                pTest3left.Leading = 10f;
    //                                pTest3right.Leading = 10f;
    //                                if (strDiscount != "" && strDiscount != "0.0")
    //                                {
    //                                    pTest3right.Add(asterisk);
    //                                }
    //                            }
    //                        }
    //                    }

    //                    if (i == 0)
    //                    {
    //                        // string fiedility = Convert.ToString(newdataset.Tables[1].
    //                        pTest2left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        // pTest2right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk2left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest2right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk2left.SetUnderline(1f, -2f);
    //                        }

    //                        pTest2right.SetAlignment("right");
    //                        pTest2left.Leading = 10f;
    //                        pTest2right.Leading = 10f;
    //                        pTest2left.IndentationLeft = 30f;
    //                        pTest2right.Add(chk2left);


    //                    }

    //                    if (i == 1)
    //                    {
    //                        pTest4left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest4right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk4left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest4right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk4left.SetUnderline(1f, -2f);
    //                        }

    //                        pTest4right.SetAlignment("right");
    //                        pTest4left.Leading = 10f;
    //                        pTest4right.Leading = 10f;
    //                        pTest4left.IndentationLeft = 30f;
    //                        pTest4right.Add(chk4left);
    //                    }
    //                    if (i == 2)
    //                    {
    //                        pTest5left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest5right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk5left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest5right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk5left.SetUnderline(1f, -2f);
    //                        }

    //                        pTest5right.SetAlignment("right");
    //                        pTest5left.Leading = 10f;
    //                        pTest5right.Leading = 10f;
    //                        pTest5left.IndentationLeft = 30f;
    //                        pTest5right.Add(chk5left);

    //                    }
    //                    if (i == 3)
    //                    {
    //                        pTest6left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        // pTest6right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk6left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest6right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk6left.SetUnderline(1f, -2f);
    //                        }

    //                        pTest6right.SetAlignment("right");
    //                        pTest6left.Leading = 10f;
    //                        pTest6right.Leading = 10f;
    //                        pTest6left.IndentationLeft = 30f;
    //                        pTest6right.Add(chk6left);
    //                    }

    //                    if (i == 4)
    //                    {
    //                        pTest7left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest7right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk7left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest7right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk7left.SetUnderline(1f, -2f);
    //                        }

    //                        pTest7right.SetAlignment("right");
    //                        pTest7left.Leading = 10f;
    //                        pTest7right.Leading = 10f;
    //                        pTest7left.IndentationLeft = 30f;
    //                        pTest7right.Add(chk7left);

    //                    }

    //                    if (i == 5)
    //                    {
    //                        pTest8left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest8right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));


    //                        Chunk chk8left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest8right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk8left.SetUnderline(1f, -2f);
    //                        }
    //                        pTest8right.SetAlignment("right");
    //                        pTest8left.Leading = 10f;
    //                        pTest8right.Leading = 10f;
    //                        pTest8left.IndentationLeft = 30f;
    //                        pTest8right.Add(chk8left);
    //                    }

    //                    if (i == 6)
    //                    {
    //                        pTest9left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest9right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        Chunk chk9left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest9right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk9left.SetUnderline(1f, -2f);
    //                        }
    //                        pTest9right.SetAlignment("right");
    //                        pTest9left.Leading = 10f;
    //                        pTest9right.Leading = 10f;
    //                        pTest9left.IndentationLeft = 30f;

    //                        pTest9right.Add(chk9left);
    //                    }

    //                    if (i == 7)
    //                    {
    //                        pTest10left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //pTest10right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk10left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest10right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk10left.SetUnderline(1f, -2f);
    //                        }

    //                        pTest10right.SetAlignment("right");
    //                        pTest10left.Leading = 10f;
    //                        pTest10right.Leading = 10f;
    //                        pTest10left.IndentationLeft = 30f;

    //                        pTest10right.Add(chk10left);
    //                    }

    //                    if (i == 8)
    //                    {
    //                        pTest11left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        // pTest11right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        Chunk chk11left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest11right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk11left.SetUnderline(1f, -2f);
    //                        }
    //                        pTest11right.SetAlignment("right");
    //                        pTest11left.Leading = 10f;
    //                        pTest11right.Leading = 10f;
    //                        pTest11left.IndentationLeft = 30f;

    //                        pTest11right.Add(chk11left);
    //                    }

    //                    if (i == 9)
    //                    {
    //                        pTest12left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        //  pTest12right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        Chunk chk12left = new Chunk(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest12right = new Paragraph();


    //                        if (underline1 == "1")
    //                        {
    //                            chk12left.SetUnderline(1f, -2f);
    //                        }

    //                        pTest12right.SetAlignment("right");
    //                        pTest12left.Leading = 10f;
    //                        pTest12right.Leading = 10f;
    //                        pTest12left.IndentationLeft = 30f;

    //                        pTest12right.Add(chk12left);
    //                    }

    //                    if (i == 10)
    //                    {
    //                        pTest13left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest13right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        pTest13right.SetAlignment("right");
    //                        pTest13left.Leading = 10f;
    //                        pTest13right.Leading = 10f;
    //                        pTest13left.IndentationLeft = 30f;
    //                    }

    //                    if (i == 11)
    //                    {
    //                        pTest14left = new Paragraph(Fidility + " " + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //                        pTest14right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //                        pTest14right.SetAlignment("right");
    //                        pTest14left.Leading = 10f;
    //                        pTest14right.Leading = 10f;
    //                        pTest14left.IndentationLeft = 30f;
    //                    }



    //                }


    //            }
    //        }
    //        // pTest2.Add(pTest2left);
    //        // pTest2.Add(pTest2right);
    //    }





    //    //-------------------------------------------------FLAT FEE LOGIC -------------------------------------------------------------------//












    //    /*** 7th Line ***/
    //    //  Paragraph p7 = new Paragraph();




    //    Paragraph p7left = null;
    //    Paragraph p7right = null;

    //    if (Accrued != "True")
    //    {
    //        p7left = new Paragraph("Annualized Total Fee for the period " + strAsOfDate, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        Chunk chk7 = new Chunk(strAnnualFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //        p7right = new Paragraph();

    //        p7right.SetAlignment("right");
    //        p7left.Leading = 10f;
    //        p7right.Leading = 10f;

    //        p7right.Add(chk7);

    //        if (strDiscount != "" && strDiscount != "0.0")
    //            p7right.Add(asterisk);
    //    }
    //    else
    //    {
    //        p7left = new Paragraph("Agreed upon flat montly fee to be accrued until initial investment ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        Chunk chk7 = new Chunk(Monthly1, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //        p7right = new Paragraph();

    //        p7right.SetAlignment("right");
    //        p7left.Leading = 10f;
    //        p7right.Leading = 10f;

    //        p7right.Add(chk7);

    //        if (strDiscount != "" && strDiscount != "0.0")
    //            p7right.Add(asterisk);
    //    }
    //    /*** 8th Line ***/
    //    // Paragraph p8 = new Paragraph();

    //    Paragraph p8left = new Paragraph("Agreed upon fee rate: ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Chunk chk8 = new Chunk(strFeeRate, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    chk8.SetUnderline(1f, -2f);
    //    Paragraph p8right = new Paragraph();

    //    p8right.SetAlignment("right");
    //    p8left.Leading = 10f;
    //    p8right.Leading = 10f;

    //    p8right.Add(chk8);

    //    if (strDiscount != "" && strDiscount != "0.0")
    //        p8right.Add(asterisk);

    //    /**** 9th Line ***/
    //    Paragraph pBlank = new Paragraph("T", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#FFFFFFF"))));
    //    pBlank.Leading = 10f;

    //    /**** 10th Line ***/
    //    // Paragraph p10 = new Paragraph();

    //    Paragraph p10left = null;
    //    Paragraph p10right = null;

    //    if (Accrued != "True")
    //    {
    //        p10left = new Paragraph("Quarterly Fee Due At This Time ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        Chunk chk10right = new Chunk(strQuaterFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        //Chunk chk10right = new Chunk(strQuaterFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        p10right = new Paragraph();

    //        p10right.SetAlignment("right");
    //        p10left.Leading = 12f;
    //        p10right.Leading = 12f;

    //        p10right.Add(chk10right);

    //        if (strDiscount != "" && strDiscount != "0.0")
    //            p10right.Add(asterisk);
    //    }
    //    else
    //    {
    //        p10left = new Paragraph("Quarterly Accrued Fee ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        Chunk chk10right = new Chunk(strQuaterFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        //Chunk chk10right = new Chunk(strQuaterFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //        p10right = new Paragraph();

    //        p10right.SetAlignment("right");
    //        p10left.Leading = 12f;
    //        p10right.Leading = 12f;

    //        p10right.Add(chk10right);

    //        if (strDiscount != "" && strDiscount != "0.0")
    //            p10right.Add(asterisk);
    //    }
    //    /***** 10th line note for discount ****/

    //    Paragraph p10noteleft = new Paragraph(" ", setFontsAll(9, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#ffffff"))));
    //    //  Chunk c10noteright = new Chunk("Annual, quarterly fee, and fee rate reflect a " + string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:P2}", Convert.ToDouble(strDiscount)) + " Discount", setFontsAll(8, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Chunk c10noteright = new Chunk("Assets under advisement fee reflects a " + string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:P2}", Convert.ToDouble(strDiscount)) + " Discount", setFontsAll(8, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph p10noteright = new Paragraph();

    //    p10noteright.SetAlignment("right");
    //    p10noteleft.Leading = 12f;
    //    p10noteleft.Leading = 12f;
    //    p10noteright.Add(asterisk);
    //    p10noteright.Add(c10noteright);


    //    //-------UnderLine after SecurityFee/Flat Fee----------

    //    //Paragraph pTest22 = new Paragraph();
    //    //Chunk chTest22 = new Chunk();
    //    //chTest22.SetUnderline(1f, -2f);
    //    //pTest22.SetAlignment("right");
    //    //pTest22.Leading = 10f;
    //    //pTest22.Add(chTest22);

    //    //-----------------------------


    //    /***** Adjustment 1st line -- new added ***/
    //    string strsign = "";
    //    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Adjustment"]) != "")
    //    {
    //        if (Convert.ToDecimal(newdataset.Tables[0].Rows[0]["ssi_Adjustment"]) < 0)
    //            strsign = "-";
    //        else
    //            strsign = "+";
    //    }
    //    Paragraph pAd1Left = new Paragraph(strsign + " " + strAdjustmentReason, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Chunk chkad1 = new Chunk(strAdjustment, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    chkad1.SetUnderline(1f, -2f);
    //    Paragraph pAd1Right = new Paragraph(chkad1);

    //    pAd1Right.SetAlignment("right");
    //    pAd1Left.Leading = 12f;
    //    pAd1Left.Leading = 12f;

    //    /***** Adjustment 2nd line -- new added ***/
    //    Paragraph pAd2Left = new Paragraph("Net adjusted quarterly fee due  at this time ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Paragraph pAd2Right = new Paragraph(strAdjustedFee, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    pAd2Right.SetAlignment("right");
    //    pAd2Left.Leading = 12f;
    //    pAd2Left.Leading = 12f;




    //    /*** 12th Line ***/
    //    Paragraph pT12 = new Paragraph();
    //    Chunk p12text1 = new Chunk("Your quarterly fees will accrue until initial investment at which time payment of the total accrual to date will be due. ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    //Chunk p12date1 = new Chunk(strAutoDebitDt + ".", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    pT12.Add(p12text1);

    //    pT12.Leading = 12f;

    //    /**** 11th Line ***/
    //    Paragraph p11 = new Paragraph("Note:  Numbers are rounded", setFontsAll(9, 0, 1, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    p11.Leading = 12f;

    //    /*** 12th Line ***/
    //    Paragraph p12 = new Paragraph();
    //    Chunk p12text = new Chunk("If you have established automatic fee payment through your account with Fidelity please do not send a check.  Please compare the Quarterly Fee Due to the amount debited from your Fidelity Statement on      ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    Chunk p12date = new Chunk(strAutoDebitDt + ".", setFontsAll(9, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    p12.Add(p12text);
    //    p12.Add(p12date);
    //    p12.Leading = 12f;

    //    /*** 13th Line ***/
    //    Paragraph p13 = new Paragraph("If you pay by check please make your check payable to �Gresham Partners, LLC� and return along with a copy of your invoice. ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    p13.Leading = 12f;
    //    /*** 14th Line ***/
    //    Paragraph p14 = new Paragraph("If you would like a detailed list of your assets under advisement or have questions, please contact your Advisor or Kate Warner at (312) 960-0214. ", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    p14.Leading = 12f;

    //    cellHeaderleft.AddElement(pLetterDate);
    //    cellHeaderright.AddElement(pInvoice);

    //    cellAddress.AddElement(pContactFullName);
    //    cellAddress.AddElement(pAddress1);
    //    if (strAddress2 != "")
    //        cellAddress.AddElement(pAddress2);

    //    if (strAddress3 != "")
    //        cellAddress.AddElement(pAddress3);
    //    cellAddress.AddElement(pCityState);

    //    if (!string.IsNullOrEmpty(strCountry))
    //        cellAddress.AddElement(pCountry);

    //    cellAddress.Colspan = 2;

    //    cellRef.AddElement(pREF);
    //    cellRef.Colspan = 2;

    //    cellline4.AddElement(p4);
    //    cellline4.Colspan = 2;

    //    cellAsOfDate.AddElement(pAsofdate);
    //    cellAsOfDate.Colspan = 2;



    //    if (BillAmt != 0)
    //    {
    //        if (Accrued != "True")
    //        {
    //            cellline6left.AddElement(p6left);
    //            cellline6right.AddElement(p6right);
    //        }
    //        else
    //        {
    //            Paragraph leftcell1 = new Paragraph();
    //            Paragraph rightcell1 = new Paragraph();
    //            leftcell1.Leading = 12f;
    //            rightcell1.Leading = 12f;
    //            celllineBlank1.AddElement(leftcell1);
    //            celllineBlankright1.AddElement(rightcell1);
    //        }

    //    }
    //    else
    //    {
    //        Paragraph leftcell1 = new Paragraph();
    //        Paragraph rightcell1 = new Paragraph();
    //        leftcell1.Leading = 12f;
    //        rightcell1.Leading = 12f;
    //        celllineBlank1.AddElement(leftcell1);
    //        celllineBlankright1.AddElement(rightcell1);
    //        //cellAsOfDate.PaddingBottom = 12f;
    //    }



    //    //------------------test------------------------------------
    //    if (pTest3left != null)
    //    {
    //        celllineTest3left.AddElement(pTest3left);
    //        celllineTest3right.AddElement(pTest3right);
    //    }

    //    if (pTestleft != null)
    //    {
    //        celllineTestleft.AddElement(pTestleft);
    //        celllineTestright.AddElement(pTestright);
    //    }

    //    if (pTest1left != null)
    //    {
    //        celllineTest1left.AddElement(pTest1left);
    //        celllineTest1right.AddElement(pTest1right);
    //    }


    //    //if (DSC > 1)
    //    //{
    //    //    if (newdataset.Tables[1].Rows.Count > 0)
    //    //    {
    //    //        for (int i = 0; i < newdataset.Tables[1].Rows.Count; i++)
    //    //        {
    //    //            string Fidility = Convert.ToString(newdataset.Tables[1].Rows[i]["ssi_Name"]);

    //    //            string FidilityAmt = RoundToZeroDecimal(Convert.ToString(newdataset.Tables[1].Rows[i]["ssi_Amount"]));
    //    //            //Paragraph pTest2 = new Paragraph();
    //    //            Paragraph pTest2left = new Paragraph(Fidility + ":", setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));
    //    //            Paragraph pTest2right = new Paragraph(FidilityAmt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#000000"))));

    //    //            pTest2right.SetAlignment("right");
    //    //            pTest2left.Leading = 10f;
    //    //            pTest2right.Leading = 10f;
    //    //            celllineTest2left.AddElement(pTest2left);
    //    //            celllineTest2right.AddElement(pTest2right);
    //    //            tblMain.AddCell(celllineTest2left);
    //    //            tblMain.AddCell(celllineTest2right);


    //    //        }
    //    //    }
    //    //}

    //    if (pTest2left != null)
    //    {
    //        celllineTest2left.AddElement(pTest2left);
    //        celllineTest2right.AddElement(pTest2right);
    //    }

    //    if (pTest4left != null)
    //    {
    //        celllineTest4left.AddElement(pTest4left);
    //        celllineTest4right.AddElement(pTest4right);
    //    }


    //    if (pTest5left != null)
    //    {
    //        celllineTest5left.AddElement(pTest5left);
    //        celllineTest5right.AddElement(pTest5right);
    //    }


    //    if (pTest6left != null)
    //    {
    //        celllineTest6left.AddElement(pTest6left);
    //        celllineTest6right.AddElement(pTest6right);
    //    }

    //    if (pTest7left != null)
    //    {
    //        celllineTest7left.AddElement(pTest7left);
    //        celllineTest7right.AddElement(pTest7right);
    //    }

    //    if (pTest8left != null)
    //    {
    //        celllineTest8left.AddElement(pTest8left);
    //        celllineTest8right.AddElement(pTest8right);
    //    }

    //    if (pTest9left != null)
    //    {
    //        celllineTest9left.AddElement(pTest9left);
    //        celllineTest9right.AddElement(pTest9right);
    //    }

    //    if (pTest10left != null)
    //    {
    //        celllineTest10left.AddElement(pTest10left);
    //        celllineTest10right.AddElement(pTest10right);
    //    }

    //    if (pTest11left != null)
    //    {
    //        celllineTest11left.AddElement(pTest11left);
    //        celllineTest11right.AddElement(pTest11right);
    //    }

    //    if (pTest12left != null)
    //    {
    //        celllineTest12left.AddElement(pTest12left);
    //        celllineTest12right.AddElement(pTest12right);
    //    }
    //    if (pTest13left != null)
    //    {
    //        celllineTest13left.AddElement(pTest13left);
    //        celllineTest13right.AddElement(pTest13right);
    //    }

    //    if (pTest14left != null)
    //    {
    //        celllineTest14left.AddElement(pTest14left);
    //        celllineTest14right.AddElement(pTest14right);
    //    }


    //    //if (underline == "1")
    //    //{
    //    //    celllineTest22right.AddElement(pTest22);
    //    //}

    //    if (pTest15left != null)
    //    {
    //        celllineTest15left.AddElement(pTest15left);
    //        celllineTest15right.AddElement(pTest15right);
    //    }

    //    if (pTest16left != null)
    //    {
    //        celllineTest16left.AddElement(pTest16left);
    //        celllineTest16right.AddElement(pTest16right);
    //    }

    //    if (pTest17left != null)
    //    {
    //        celllineTest17left.AddElement(pTest17left);
    //        celllineTest17right.AddElement(pTest17right);
    //    }

    //    if (pTest18left != null)
    //    {
    //        celllineTest18left.AddElement(pTest18left);
    //        celllineTest18right.AddElement(pTest18right);
    //    }



    //    if (pTest19left != null)
    //    {
    //        celllineTest19left.AddElement(pTest19left);
    //        celllineTest19right.AddElement(pTest19right);
    //    }

    //    if (pTest20left != null)
    //    {
    //        celllineTest20left.AddElement(pTest20left);
    //        celllineTest20right.AddElement(pTest20right);
    //    }

    //    if (pTest21left != null)
    //    {
    //        celllineTest21left.AddElement(pTest21left);
    //        celllineTest21right.AddElement(pTest21right);
    //    }



    //    //if (underline1 == "1")
    //    //{
    //    //    celllineTest22right.AddElement(pTest22);
    //    //}

    //    //celllineTest22right.AddElement();
    //    //-----------------------test-----------------------------

    //    if (Accrued != "True")
    //    {
    //        cellline7left.AddElement(p7left);
    //        cellline7right.AddElement(p7right);
    //    }
    //    else
    //    {
    //        cellline7left.AddElement(p7left);
    //        cellline7right.AddElement(p7right);
    //    }
    //    //cellline7left.Leading = 12F;


    //    if (strFeeRate != "")
    //    {
    //        if (Accrued != "True")
    //        {
    //            cellline8left.AddElement(p8left);
    //            cellline8right.AddElement(p8right);

    //        }
    //        else
    //        {
    //            Paragraph leftcell = new Paragraph();
    //            Paragraph rightcell = new Paragraph();
    //            leftcell.Leading = 12f;
    //            rightcell.Leading = 12f;
    //            celllineBlank.AddElement(leftcell);
    //            celllineBlankright.AddElement(rightcell);

    //            cellAsOfDate.PaddingBottom = 12f;
    //        }
    //    }

    //    else
    //    {
    //        //Chunk lochunk = new Chunk("\n");
    //        //Paragraph BlankRow = new Paragraph();
    //        //BlankRow.Add(lochunk);
    //        //cellline7left.AddElement(p7left);
    //        //cellline7right.AddElement(p7right);
    //        Paragraph leftcell = new Paragraph();
    //        Paragraph rightcell = new Paragraph();
    //        leftcell.Leading = 12f;
    //        rightcell.Leading = 12f;
    //        celllineBlank.AddElement(leftcell);
    //        celllineBlankright.AddElement(rightcell);

    //        cellAsOfDate.PaddingBottom = 12f;
    //        //celllineBlank.Colspan = 2;
    //        //celllineBlankright.Colspan = 2;
    //        //pdoc.Add(BlankRow);
    //    }
    //    //if(strFeeRate!="" || BillAmt!=0)
    //    //{
    //    //cellline9.AddElement(pBlank);
    //    //cellline9.Colspan = 2;
    //    //cellline9.BorderWidthTop = 1f;
    //    //}

    //    cellline9.AddElement(pBlank);
    //    cellline9.Colspan = 2;
    //    cellline9.BorderWidthTop = 1f;

    //    if (Accrued != "True")
    //    {
    //        cellline10left.AddElement(p10left);
    //        cellline10right.AddElement(p10right);
    //    }
    //    else
    //    {
    //        cellline10left.AddElement(p10left);
    //        cellline10right.AddElement(p10right);
    //    }

    //    if (Accrued != "True")
    //    {
    //        cellline10noteleft.AddElement(p10noteleft);
    //        cellline10noteright.AddElement(p10noteright);
    //    }

    //    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Adjustment"]) != "")
    //    {
    //        cellline1AdjustmentLeft.AddElement(pAd1Left);
    //        cellline1AdjustmentRight.AddElement(pAd1Right);

    //        cellline2AdjustmentLeft.AddElement(pAd2Left);
    //        cellline2AdjustmentRight.AddElement(pAd2Right);
    //    }

    //    cellline11.AddElement(p11);
    //    cellline11.Colspan = 2;

    //    if (Accrued != "True")
    //    {
    //        cellline12.AddElement(p12);
    //        cellline12.Colspan = 2;

    //        cellline13.AddElement(p13);
    //        cellline13.Colspan = 2;

    //        cellline14.AddElement(p14);
    //        cellline14.Colspan = 2;
    //    }
    //    else
    //    {
    //        celllineT12.AddElement(pT12);
    //        celllineT12.Colspan = 2;

    //    }
    //    cellline10noteright.Colspan = 2;

    //    tblMain.AddCell(cellHeaderleft);
    //    tblMain.AddCell(cellHeaderright);
    //    tblMain.AddCell(cellAddress);
    //    tblMain.AddCell(cellRef);
    //    tblMain.AddCell(cellline4);

    //    tblMain.AddCell(cellAsOfDate);
    //    if (BillAmt != 0)
    //    {
    //        if (Accrued != "True")
    //        {
    //            tblMain.AddCell(cellline6left);
    //            tblMain.AddCell(cellline6right);
    //        }
    //        else
    //        {
    //            tblMain.AddCell(celllineBlank1);
    //            tblMain.AddCell(celllineBlankright1);
    //        }


    //    }
    //    else
    //    {
    //        tblMain.AddCell(celllineBlank1);
    //        tblMain.AddCell(celllineBlankright1);

    //    }
    //    //----------------------------test-----------------------------------------

    //    //if (Convert.ToDecimal(strAnnFeeBeforeAdditionalFee1) != 0)
    //    //{
    //    //    tblMain.AddCell(celllineTest3left);
    //    //    tblMain.AddCell(celllineTest3right);
    //    //}


    //    //if (Convert.ToDecimal(strRelationshipFee1) != 0)
    //    //{

    //    //    tblMain.AddCell(celllineTestleft);
    //    //    tblMain.AddCell(celllineTestright);
    //    //}

    //    //if (Convert.ToDecimal(strSetupFee1) != 0)
    //    //{


    //    //    tblMain.AddCell(celllineTest1left);
    //    //    tblMain.AddCell(celllineTest1right);
    //    //}
    //    if (pTest3left != null)
    //    {
    //        tblMain.AddCell(celllineTest3left);
    //        tblMain.AddCell(celllineTest3right);
    //    }
    //    if (pTestleft != null)
    //    {
    //        tblMain.AddCell(celllineTestleft);
    //        tblMain.AddCell(celllineTestright);
    //    }

    //    if (pTest1left != null)
    //    {
    //        tblMain.AddCell(celllineTest1left);
    //        tblMain.AddCell(celllineTest1right);
    //    }

    //    if (pTest2left != null)
    //    {
    //        tblMain.AddCell(celllineTest2left);
    //        tblMain.AddCell(celllineTest2right);
    //    }

    //    if (pTest4left != null)
    //    {
    //        tblMain.AddCell(celllineTest4left);
    //        tblMain.AddCell(celllineTest4right);
    //    }
    //    if (pTest5left != null)
    //    {
    //        tblMain.AddCell(celllineTest5left);
    //        tblMain.AddCell(celllineTest5right);
    //    }
    //    if (pTest6left != null)
    //    {
    //        tblMain.AddCell(celllineTest6left);
    //        tblMain.AddCell(celllineTest6right);
    //    }
    //    if (pTest7left != null)
    //    {
    //        tblMain.AddCell(celllineTest7left);
    //        tblMain.AddCell(celllineTest7right);
    //    }
    //    if (pTest8left != null)
    //    {
    //        tblMain.AddCell(celllineTest8left);
    //        tblMain.AddCell(celllineTest8right);
    //    }


    //    if (pTest9left != null)
    //    {
    //        tblMain.AddCell(celllineTest9left);
    //        tblMain.AddCell(celllineTest9right);
    //    }

    //    if (pTest10left != null)
    //    {
    //        tblMain.AddCell(celllineTest10left);
    //        tblMain.AddCell(celllineTest10right);
    //    }

    //    if (pTest11left != null)
    //    {
    //        tblMain.AddCell(celllineTest11left);
    //        tblMain.AddCell(celllineTest11right);
    //    }

    //    if (pTest12left != null)
    //    {
    //        tblMain.AddCell(celllineTest12left);
    //        tblMain.AddCell(celllineTest12right);
    //    }

    //    if (pTest13left != null)
    //    {
    //        tblMain.AddCell(celllineTest13left);
    //        tblMain.AddCell(celllineTest13right);
    //    }

    //    if (pTest14left != null)
    //    {
    //        tblMain.AddCell(celllineTest14left);
    //        tblMain.AddCell(celllineTest14right);
    //    }

    //    //if(underline1=="1")
    //    //{
    //    //    tblMain.AddCell(celllineTest22right);
    //    //}
    //    if (pTest15left != null)
    //    {
    //        tblMain.AddCell(celllineTest15left);
    //        tblMain.AddCell(celllineTest15right);
    //    }

    //    if (pTest16left != null)
    //    {
    //        tblMain.AddCell(celllineTest16left);
    //        tblMain.AddCell(celllineTest16right);
    //    }

    //    if (pTest17left != null)
    //    {
    //        tblMain.AddCell(celllineTest17left);
    //        tblMain.AddCell(celllineTest17right);
    //    }

    //    if (pTest18left != null)
    //    {
    //        tblMain.AddCell(celllineTest18left);
    //        tblMain.AddCell(celllineTest18right);
    //    }

    //    if (pTest19left != null)
    //    {
    //        tblMain.AddCell(celllineTest19left);
    //        tblMain.AddCell(celllineTest19right);
    //    }

    //    if (pTest20left != null)
    //    {
    //        tblMain.AddCell(celllineTest20left);
    //        tblMain.AddCell(celllineTest20right);
    //    }


    //    if (pTest21left != null)
    //    {
    //        tblMain.AddCell(celllineTest21left);
    //        tblMain.AddCell(celllineTest21right);
    //    }

    //    //if (underline == "1")
    //    //{
    //    //    tblMain.AddCell(celllineTest22right);
    //    //}

    //    //----------------------------------test----------------------------------------
    //    if (Accrued != "True")
    //    {
    //        tblMain.AddCell(cellline7left);
    //        tblMain.AddCell(cellline7right);
    //    }
    //    else
    //    {
    //        tblMain.AddCell(cellline7left);
    //        tblMain.AddCell(cellline7right);
    //    }

    //    if (strFeeRate != "")
    //    {
    //        if (Accrued != "True")
    //        {
    //            tblMain.AddCell(cellline8left);
    //            tblMain.AddCell(cellline8right);
    //        }
    //        else
    //        {
    //            tblMain.AddCell(celllineBlank);
    //            tblMain.AddCell(celllineBlankright);
    //        }
    //    }
    //    else
    //    {
    //        tblMain.AddCell(celllineBlank);
    //        tblMain.AddCell(celllineBlankright);
    //    }


    //    tblMain.AddCell(cellline9);

    //    if (Accrued != "True")
    //    {
    //        tblMain.AddCell(cellline10left);
    //        tblMain.AddCell(cellline10right);
    //    }
    //    else
    //    {
    //        tblMain.AddCell(cellline10left);
    //        tblMain.AddCell(cellline10right);
    //    }
    //    //   tblMain.AddCell(cellline10noteleft);
    //    if (strDiscount != "" && strDiscount != "0.0")
    //        tblMain.AddCell(cellline10noteright);
    //    if (Convert.ToString(newdataset.Tables[0].Rows[0]["ssi_Adjustment"]) != "")
    //    {
    //        tblMain.AddCell(cellline1AdjustmentLeft);
    //        tblMain.AddCell(cellline1AdjustmentRight);
    //        tblMain.AddCell(cellline2AdjustmentLeft);
    //        tblMain.AddCell(cellline2AdjustmentRight);
    //    }
    //    tblMain.AddCell(cellline11);
    //    if (Accrued != "True")
    //    {
    //        tblMain.AddCell(cellline12);
    //        tblMain.AddCell(cellline13);
    //        tblMain.AddCell(cellline14);
    //    }
    //    else
    //    {
    //        tblMain.AddCell(celllineT12);
    //    }
    //    pdoc.Add(tblMain);

    //    pdoc.Close();

    //    if (DSCount > 0)
    //    {
    //        try
    //        {
    //            FileInfo loFile = new FileInfo(ls);
    //            loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
    //        }
    //        catch
    //        { }

    //    }
    //    else
    //    {
    //        fsFinalLocation = "";
    //    }

    //    return fsFinalLocation.Replace(".xls", ".pdf");
    //}

    #endregion
    public string currencyFormat(string Value)
    {

        string value = Value.Replace(",", "").Replace("$", "").Replace("%", "").Replace("(", "-").Replace(")", "");


        if (value != "")
        {
            decimal ul = 0;

            ul = Convert.ToDecimal(value);

            value = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-US"), "{0:C0}", ul);
        }

        return value;
    }

    #endregion



    #region Common Methods
    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic, iTextSharp.text.Color foColor)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        }
        if (italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanai.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\Fverdanaz.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        }

        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //}
        //if (bold == 1 && italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        //}
        return font;
        #endregion
    }
    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanai.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanaz.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        }


        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    //FTI_____.PFM
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //}

        return font;
        #endregion
    }


    public iTextSharp.text.Font setFontsAllFrutiger(int size, int bold, int italic)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = HttpContext.Current.Server.MapPath(".");

        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        }

        return font;
        #endregion
    }

    public iTextSharp.text.Font setFontsAllTimesNewRoman(int size, int bold, int italic)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM Times NEW ROMAN
        string fontpath = HttpContext.Current.Server.MapPath(".");

        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\times.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\timesbd.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\timesi.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\TimesNewRoman\\timesbi.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        }

        return font;
        #endregion
    }


    public iTextSharp.text.Font setFontsverdana()
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, 9, iTextSharp.text.Font.NORMAL);



        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    //FTI_____.PFM
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //}

        return font;
        #endregion
    }


    public iTextSharp.text.Font setFontsAll1(int size, int bold)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = HttpContext.Current.Server.MapPath(".");

        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);

            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }


        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);

        //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}


        return font;
        #endregion
    }

    public void setTableProperty(iTextSharp.text.Table fotable, ReportType Type)
    {
        //int[] headerwidths = { 28, 9, 9, 9, 9, 9, 9, 9, 7 };

        setWidthsoftable(fotable);

        //fotable.Width = 100;
        if (Type == ReportType.FundMemorandum)
        {
            fotable.Alignment = 0;
        }
        else
        {
            fotable.Alignment = 1;
        }
        fotable.Border = 0;
        fotable.Cellspacing = 0;
        fotable.Cellpadding = 3;
        fotable.Locked = false;

    }

    public void setWidthsoftable(iTextSharp.text.Table fotable)
    {

        switch (lsTotalNumberofColumns)
        {
            case "2":
                int[] headerwidths2 = { 70, 18 };
                fotable.SetWidths(headerwidths2);
                fotable.Width = 88;
                break;
            case "3":
                //int[] headerwidths3 = { 30, 9, 9 };
                int[] headerwidths3 = { 70, 18, 22 };//changed on 6 feb 2015
                fotable.SetWidths(headerwidths3);
                fotable.Width = 92;
                break;
            case "4":
                int[] headerwidths4 = { 45, 5, 15, 15 };
                fotable.SetWidths(headerwidths4);
                fotable.Width = 80;
                break;
            case "5":
                int[] headerwidths5 = { 30, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths5);
                fotable.Width = 67;
                break;
            case "6":
                int[] headerwidths6 = { 27, 11, 11, 8, 5, 7 };
                fotable.SetWidths(headerwidths6);
                fotable.Width = 70;
                break;
            case "7":
                int[] headerwidths7 = { 30, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths7);
                fotable.Width = 85;
                break;
            case "8":
                int[] headerwidths8 = { 30, 9, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths8);
                fotable.Width = 94;
                break;
            case "9":
                int[] headerwidths9 = { 27, 9, 9, 9, 9, 9, 9, 9, 7 };
                fotable.SetWidths(headerwidths9);
                fotable.Width = 97;
                break;

            case "10":
                int[] headerwidths10 = { 25, 8, 8, 8, 8, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths10);
                fotable.Width = 97; break;
            case "11":
                //int[] headerwidths11 = { 25, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7 };
                int[] headerwidths11 = { 25, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths11);
                fotable.Width = 95; break;
            case "21":
                int[] headerwidths12 = { 7, 7, 12, 7, 7, 7, 7, 35, 7, 7, 7, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                fotable.SetWidths(headerwidths12);
                fotable.Width = 150; break;
            case "13":
                int[] headerwidths13 = { 12, 2, 15, 2, 20, 20, 2, 15, 5, 15, 15, 2, 15 };
                fotable.SetWidths(headerwidths13);
                fotable.Width = 100; break;
            case "14":
                int[] headerwidths14 = { 30, 9 };
                fotable.SetWidths(headerwidths14);
                fotable.Width = 100;
                break;
            case "15":
                int[] headerwidths15 = { 30, 9 };
                fotable.SetWidths(headerwidths15);
                fotable.Width = 39;
                break;
            case "16":
                int[] headerwidths16 = { 30, 9 };
                fotable.SetWidths(headerwidths16);
                fotable.Width = 39;
                break;
            case "17":
                int[] headerwidths17 = { 30, 9 };
                fotable.SetWidths(headerwidths17);
                fotable.Width = 39;
                break;
            case "18":
                int[] headerwidths18 = { 30, 9 };
                fotable.SetWidths(headerwidths18);
                fotable.Width = 39;
                break;
            case "19":
                int[] headerwidths19 = { 30, 9 };
                fotable.SetWidths(headerwidths19);
                fotable.Width = 39;
                break;
            case "29":
                int[] headerwidths20 = { 30, 9 };
                fotable.SetWidths(headerwidths20);
                fotable.Width = 39;
                break;

        }
    }


    public iTextSharp.text.Font Font8Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(8, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(8, 0, 0);
    }

    public iTextSharp.text.Font Font7GreyItalic()
    {
        return setFontsAll(7, 0, 1, new iTextSharp.text.Color(216, 216, 216));
    }




    public void AddFooter(iTextSharp.text.Document document)
    {
        Phrase footPhraseImg = new Phrase("Gresham Advisors, LLC | 333 W. Wacker Dr. Suite 700 | Chicago, IL 60606 | P 312.960.0200 | F 312.960.0204 | www.greshampartners.com", setFontsAll(6, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter footer = new HeaderFooter(footPhraseImg, false);
        footer.Border = iTextSharp.text.Rectangle.NO_BORDER;
        footer.Alignment = Element.ALIGN_CENTER;
        document.Footer = footer;
    }

    public void AddHeader(iTextSharp.text.Document document)
    {
        Paragraph pimage = new Paragraph();
        iTextSharp.text.Image SignatureJpg = iTextSharp.text.Image.GetInstance(iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png"));
        SignatureJpg.SetAbsolutePosition(45, 557);//540
        //SignatureJpg.ScaleToFit(45, 557);
        SignatureJpg.ScalePercent(10);
        pimage.Add(SignatureJpg);

        //Phrase footPhraseImg = new Phrase(, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        HeaderFooter Header = new HeaderFooter(pimage, false);
        Header.Border = iTextSharp.text.Rectangle.NO_BORDER;
        Header.Alignment = Element.ALIGN_CENTER;
        document.Header = Header;
    }

    public void setHeaderCapitalCallStatementCustom(Document foDocument)
    {
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(3);   // 2 rows, 3 columns   
        lsTotalNumberofColumns = "3";
        setTableProperty(loTable, ReportType.CapitalCallStatementCustom);

        iTextSharp.text.Chunk lochunk2 = new Chunk("Partnership Name", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk6 = new Chunk("", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell2 = new Cell();
        loCell2.Add(lochunk2);
        loCell2.Add(lochunk6);
        loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell2.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell2.Border = 0;
        loCell2.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell2.MaxLines = 3;
        loCell2.Leading = 10f;
        loTable.AddCell(loCell2);

        iTextSharp.text.Chunk lochunk3 = new Chunk("Percent of \n Commitment \n Called                                 ", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk7 = new Chunk("", setFontsAll(9, 1, 0));
        iTextSharp.text.Cell loCell3 = new Cell();
        loCell3.Add(lochunk3);
        //loCell3.Add(lochunk7);
        loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell3.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell3.Border = 0;
        //loCell3.EnableBorderSide(2);
        loCell3.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell3.MaxLines = 3;
        loCell3.Leading = 10f;
        loTable.AddCell(loCell3);

        iTextSharp.text.Chunk lochunk8 = new Chunk("Current Capital Call", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk9 = new Chunk("", setFontsAll(9, 1, 0));
        iTextSharp.text.Cell loCell4 = new Cell();
        loCell4.Add(lochunk8);
        //loCell4.Add(lochunk9);
        loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell4.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell4.Border = 0;
        //loCell4.EnableBorderSide(2);
        loCell4.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell4.MaxLines = 3;
        loCell4.Leading = 10f;
        loTable.AddCell(loCell4);

        foDocument.Add(loTable);

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //png.SetAbsolutePosition(45, 557);//540
        ////png.ScaleToFit(288f, 42f);
        //png.ScalePercent(10);
        //foDocument.Add(png);
    }
    public void getTotalAmountCapitalCallStatementCustom(Document foDocument, string strTotalAmount)
    {
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(3);   // 2 rows, 3 columns   
        lsTotalNumberofColumns = "3";
        setTableProperty(loTable, ReportType.CapitalCallStatementCustom);

        iTextSharp.text.Chunk lochunk2 = new Chunk(" ", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk6 = new Chunk("", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell2 = new Cell();
        loCell2.Add(lochunk2);
        loCell2.Add(lochunk6);
        loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell2.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell2.Border = 0;
        loCell2.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell2.MaxLines = 3;
        loCell2.Leading = 10f;
        loTable.AddCell(loCell2);

        iTextSharp.text.Chunk lochunk3 = new Chunk("Total", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk7 = new Chunk("", setFontsAll(9, 1, 0));
        iTextSharp.text.Cell loCell3 = new Cell();
        loCell3.Add(lochunk3);
        //loCell3.Add(lochunk7);
        loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell3.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell3.Border = 0;
        //loCell3.EnableBorderSide(2);
        loCell3.VerticalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell3.MaxLines = 3;
        loCell3.Leading = 10f;
        loTable.AddCell(loCell3);

        iTextSharp.text.Chunk lochunk8 = new Chunk(RoundUp(strTotalAmount), setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk9 = new Chunk("", setFontsAll(9, 1, 0));
        iTextSharp.text.Cell loCell4 = new Cell();
        loCell4.Add(lochunk8);
        //loCell4.Add(lochunk9);
        loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell4.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell4.Border = 0;
        //loCell4.EnableBorderSide(2);
        loCell4.VerticalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        //loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell4.MaxLines = 3;
        loCell4.Leading = 10f;
        loTable.AddCell(loCell4);

        foDocument.Add(loTable);

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //png.SetAbsolutePosition(45, 557);//540
        ////png.ScaleToFit(288f, 42f);
        //png.ScalePercent(10);
        //foDocument.Add(png);
    }

    public void setHeaderDistributionStatementCustom(Document foDocument)
    {
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(3);   // 2 rows, 2 columns  
        lsTotalNumberofColumns = "3";
        setTableProperty(loTable, ReportType.CapitalCallStatementCustom);

        iTextSharp.text.Chunk lochunk2 = new Chunk("Partnership Name", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk6 = new Chunk("", setFontsAll(7, 1, 0));
        iTextSharp.text.Cell loCell2 = new Cell();
        loCell2.Add(lochunk2);
        loCell2.Add(lochunk6);
        loCell2.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell2.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell2.Border = 0;
        loCell2.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell2.MaxLines = 3;
        loCell2.Leading = 10f;
        loTable.AddCell(loCell2);


        //iTextSharp.text.Chunk lochunk3 = new Chunk("Percent of \n Commitment \n Distributed                                  ", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk3 = new Chunk("Percent of Commitment Distributed", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk7 = new Chunk("", setFontsAll(9, 1, 0));
        iTextSharp.text.Cell loCell3 = new Cell();
        loCell3.Add(lochunk3);
        //loCell3.Add(lochunk7);
        loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell3.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell3.Border = 0;
        //loCell3.EnableBorderSide(2);
        loCell3.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        //loCell3.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_RIGHT;
        loCell3.MaxLines = 3;
        loCell3.Leading = 10f;
        loTable.AddCell(loCell3);


        iTextSharp.text.Chunk lochunk4 = new Chunk("Current Distribution", setFontsAll(9, 1, 0));
        iTextSharp.text.Chunk lochunk8 = new Chunk("", setFontsAll(9, 1, 0));
        iTextSharp.text.Cell loCell4 = new Cell();
        loCell4.Add(lochunk4);
        loCell4.Add(lochunk8);
        loCell4.HorizontalAlignment = iTextSharp.text.Cell.ALIGN_CENTER;
        loCell4.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
        loCell4.Border = 0;
        loCell4.VerticalAlignment = iTextSharp.text.Cell.ALIGN_BOTTOM;
        loCell4.MaxLines = 3;
        loCell4.Leading = 10f;
        loTable.AddCell(loCell4);

        foDocument.Add(loTable);

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
        //png.SetAbsolutePosition(45, 557);//540
        ////png.ScaleToFit(288f, 42f);
        //png.ScalePercent(10);
        //foDocument.Add(png);

    }


    public string RemoveStyle(string html)
    {
        //start by completely removing all unwanted tags 
        //System.Text.RegularExpressions.Regex.Replace(html, @"<(.|\n)*?>", string.Empty);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"<[/]?(font|span|xml|del|ins|[ovwxp]:\w+)[^>]*?>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //then run another pass over the html (twice), removing unwanted attributes 
        html = System.Text.RegularExpressions.Regex.Replace(html, @"(mso-bidi-font-style: normal)*<( )*font-family: Verdana( )*(/)*>", "\r", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"(mso-bidi-font-weight: normal)*<( )*font-family: Verdana( )*>", "\r", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        html = System.Text.RegularExpressions.Regex.Replace(html, @"<([^>]*)(?:class|[ovwxp]:\w+)=(?:'[^']*'|""[^""]*""|[^>]+)([^>]*)>", "<$1$2>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        html = System.Text.RegularExpressions.Regex.Replace(html, @"<([^>]*)(?:class|[ovwxp]:\w+)=(?:'[^']*'|""[^""]*""|[^>]+)([^>]*)>", "<$1$2>", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        //html.Replace("<p", "\n").Replace("</p>", "\n").Replace("<li", "\n").Replace("</li>", "\n").Replace("b", "@").Replace("</b>", "@").Replace("<i", "@").Replace("</i>", "@");
        return html;
    }

    public void AddStyle(string selector, string styles)
    {

        string STYLE_DEFAULT_TYPE = "style";
        this._Styles.LoadTagStyle(selector, STYLE_DEFAULT_TYPE, styles);
    }

    #endregion
}

    #endregion