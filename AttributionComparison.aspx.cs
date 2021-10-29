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
using System.Security.Principal;
using System.Data.SqlClient;
//using CrmSdk;
using System.IO;
using Spire.Xls;
using System.Data.Common;
using System.Xml;

public partial class AttributionComparison : System.Web.UI.Page
{
    GeneralMethods clsGM = new GeneralMethods();
    DB clsdb = new DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //BindHouseHolds();
            //FillYear();
            //Bind_AssetClass();

            //if (Request.QueryString.Count == 0)
            //{
            //    trMarketableGA.Style.Add("display", "none");
            //    trAssetClass.Style.Add("display", "none");
            //}
        }
    }

    private void FillYear()
    {
        //for (int i = DateTime.Now.Year; i > DateTime.Now.Year - 15; i--)
        //{
        //    ddlYear.Items.Add(Convert.ToString(i));
        //}

        //ddlYear.SelectedValue = DateTime.Now.Year.ToString();
    }

    public void Bind_AssetClass()
    {
        //string sqlstr = "SP_S_ASSET_CLASS";
        //clsGM.getListForBindListBox(lstAssetClass, sqlstr, "sas_name", "sas_assetclassId");

        //lstAssetClass.Items.Insert(0, "All");
        //lstAssetClass.Items[0].Value = "0";
        //lstAssetClass.SelectedIndex = 0;

        /*
        for (int i = 0; i < lstAssetClass.Items.Count; i++)
        {
            if (lstAssetClass.Items[i].Value.ToString() == "e2a78beb-d604-de11-a38c-001d09665e8f")//Domestic Equity
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "028b5efb-d604-de11-a38c-001d09665e8f")//Fixed Income
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "8413896b-4925-df11-b686-001d09665e8f")//Global Opportunistic
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "c2a2d71c-d704-de11-a38c-001d09665e8f")//Illiquid Real Assets
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "42b39247-d704-de11-a38c-001d09665e8f")//International Equity
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "0332530a-1ad3-df11-9789-0019b9e7ee05")//Liquid Real Assets
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "2287692a-d704-de11-a38c-001d09665e8f")//Low Volatility Hedged Strategies
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "02ffe912-d704-de11-a38c-001d09665e8f")//Private Equity
                lstAssetClass.Items[i].Selected = true;
            if (lstAssetClass.Items[i].Value.ToString() == "9776259d-0392-4de0-8a12-0399724abf8d") //Cash and Equivalents
                lstAssetClass.Items[i].Selected = true;
        }
        */
        //lstAssetClass.Items[3].Selected = true;
        //lstAssetClass.Items[4].Selected = true;
        //lstAssetClass.Items[5].Selected = true;
        //lstAssetClass.Items[6].Selected = true;
        //lstAssetClass.Items[7].Selected = true;
        //lstAssetClass.Items[9].Selected = true;
        //lstAssetClass.Items[12].Selected = true;
        //lstAssetClass.Items[16].Selected = true;
    }

    public void BindHouseHolds()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        //DB clsDB = new DB();
        //DataSet loDataset = clsDB.getDataSet("sp_s_Get_HouseHoldName");
        //ddlHousehold.Items.Clear();
        //ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        //for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        //{
        //    ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        //}

    }
    public void FillReportRollUpGroup()
    {
        //string HHUID = ddlHousehold.SelectedValue == "" ? "" : "'" + ddlHousehold.SelectedValue + "'";
        //DB clsDB = new DB();
        //ddlReportRollupgrp.Items.Clear();
        //ddlReportRollupgrp.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        //if (HHUID != "")
        //{
        //    DataSet loDataset = clsDB.getDataSet("SP_S_GROUPNAME  @MarkovType = 2,@HHUUID =" + HHUID + "");
        //    for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        //    {
        //        ddlReportRollupgrp.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["GroupName"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["sas_reportrollupgroupid"])));
        //    }
        //}
    }
    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        //FillAllocationGroup();
        FillReportRollUpGroup();
        lblMessage.Text = "";
    }

    //public void FillAllocationGroup()
    //{
    //    DB clsDB = new DB();
    //    ddlAllocationGroup.Items.Clear();
    //    ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
    //    DataSet loDataset = clsDB.getDataSet("SP_S_Advent_Allocation_Group  @Householdname ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "'");
    //    for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
    //    {
    //        ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"])));
    //    }

    //}

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        //generatesExcelsheets();
        GenerateExcelReport();
    }

    private DataSet GetReportData()
    {
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery = string.Empty;

        //object Household = ddlHousehold.SelectedValue == "0" ? "null" : "'" + ddlHousehold.SelectedValue + "'";
        object FamilyGAGroupFlg = ddlGroups.SelectedValue == "9" ? "null" : "'" + ddlGroups.SelectedValue + "'";
        //object Year = ddlYear.SelectedItem.Text;
        object AsofDate = txtAsOfDate.Text == "" ? "null" : "'" + txtAsOfDate.Text + "'";
        //object ReportRollupgrp = ddlReportRollupgrp.SelectedValue == "0" ? "null" : "'" + ddlReportRollupgrp.SelectedItem.Text.Replace("'", "''") + "'";

        //string strAssetClass = lstAssetClass.SelectedValue == "0" ? "null" : "'" + GetAllItemsTextFromListBox(lstAssetClass, true) + "'";


        try
        {
            greshamquery = "EXEC SP_R_ATTRIBUTION_COMPARISON "
                            + "@AsOfDate = " + AsofDate + ""
                             + ",@FamilyGAGroupFlg = " + FamilyGAGroupFlg + ""; ;

            ds_gresham = clsdb.getDataSet(greshamquery);
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {

            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Detail.InnerText;
        }
        catch (Exception exc)
        {
            lblMessage.Text = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exc.Message;
        }

        return ds_gresham;
    }

    private void GenerateExcelReport()
    {

        #region Spire License Code
        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion


        lblMessage.Text = "";
        DataSet lodataset = GetReportData();
        DataSet ds = lodataset.Copy();

        if (ds.Tables.Count > 0)
        {
            if (ds.Tables[0].Rows.Count < 1)
            {
                lblMessage.Text = "No Records found.";
                return;
            }
        }
        else
        {
            if (lblMessage.Text == "")
                lblMessage.Text = "No Records found.";
            return;
        }

        String lsFileNamforFinalXls = System.DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\AttributionComparison.xlsx");
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);
        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls.Replace("xlsx", "xml"));
        FileInfo loFile = new FileInfo(strDirectory1);
        loFile.CopyTo(strDirectory, true);
        if (ds.Tables.Count > 0)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(strDirectory);
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                DataTable t = ds.Tables[i];
                //export datatable to excel

                Worksheet sheet = book.Worksheets[i];
                sheet.InsertDataTable(t, true, 1, 1, -1, -1);
                sheet.Range[1, 1, 8000, 200].RowHeight = 16.5; // all data row height
                sheet.Range[1, 1, 8000, 200].AutoFitColumns();
                sheet.Range[1, 1, 1, 200].Style.Font.IsBold = true;

                for (int exlrow = 1; exlrow < ds.Tables[0].Rows.Count + 2; exlrow++)
                {
                    if (Convert.ToString(sheet.Range[exlrow, 7].Value) == "")
                    {
                        sheet.Range[exlrow, 7].Text = "-";
                        sheet.Range[exlrow, 7].HorizontalAlignment = HorizontalAlignType.Center;
                    }
                }

                sheet.Range[2, 4, 8000, 4].NumberFormat = "mmm-yy";
                sheet.Range[2, 5, 8000, 6].NumberFormat = "$#,##0.00_);($#,##0.00)";
                sheet.Range[2, 5, 8000, 7].NumberFormat = "#,##0.00_);(#,##0.00)";
                sheet.Range[2, 8, 8000, 10].NumberFormat = "0.00%";










            }

            //book.SaveToFile(ExceptionFilePath);

            //book.SaveAsXml(strDirectory2);
            //book = null;
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
            //loFile = new FileInfo(strDirectory2);
            //loFile.Delete();
            book.SaveToFile(strDirectory, ExcelVersion.Version2016);
        }

       

        #region New xls to xlsx code
        //Workbook workbook1 = new Workbook();
        //workbook1.LoadFromXml(strDirectory2);
        //workbook1.SaveToFile(strDirectory, ExcelVersion.Version2010);

        //workbook1 = new Workbook();
        ////  workbook1.LoadFromFile(strDirectory.Replace("xls", "xlsx"));
        //workbook1.LoadFromFile(strDirectory);
        //Worksheet sheet1 = workbook1.Worksheets[0];
        //sheet1.Range[6, 1, 6, 5].Style.Color = System.Drawing.Color.FromArgb(216, 216, 216);
        ////workbook1.SaveToFile(strDirectory.Replace("xls", "xlsx"), ExcelVersion.Version2010);
        //workbook1.SaveToFile(strDirectory, ExcelVersion.Version2010);
        #region PageSetup
        Workbook workbook1 = new Workbook();
        workbook1.LoadFromFile(strDirectory);
        Worksheet sheet1 = workbook1.Worksheets[0];
        workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);

        var setup = sheet1.PageSetup;
        setup.FitToPagesWide = 1;
        //setup.FitToPagesTall = 1;
        setup.IsFitToPage = true;
        setup.PaperSize = PaperSizeType.PaperA4;
        setup.Orientation = PageOrientationType.Landscape;
        setup.FitToPagesWide = 1;
        setup.FitToPagesTall = 0;
        setup.CenterHorizontally = true;
        setup.CenterVertically = false;
        workbook1.SaveToFile(strDirectory, ExcelVersion.Version2016);
        #endregion

        loFile = new FileInfo(strDirectory2);
        loFile.Delete();
        loFile = null;
     
        lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + lsFileNamforFinalXls;
        #endregion

      //  lsFileNamforFinalXls = "./ExcelTemplate/" + lsFileNamforFinalXls;
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + Path.GetFileName(lsFileNamforFinalXls) + "");
        Response.TransmitFile(lsFileNamforFinalXls);
        Response.End();
    }


    public string GetAllItemsTextFromListBox(ListBox lstBox, bool IsOnlySelected)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 0; i < lstBox.Items.Count; i++)
            {
                if (IsOnlySelected)
                {
                    if (lstBox.Items[i].Selected)
                    {
                        lstselecteditems = lstselecteditems + "|" + lstBox.Items[i].Text;
                        //insert command
                    }
                }
                else
                {
                    lstselecteditems = lstselecteditems + "|" + lstBox.Items[i].Text;
                }
            }
            if (lstselecteditems != "")
            {
                lstselecteditems = lstselecteditems.Substring(1);
            }
        }


        return lstselecteditems;
    }
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     