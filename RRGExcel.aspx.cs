using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Spire.Xls;

public partial class DumpExcel : System.Web.UI.Page
{
    string Server = AppLogic.GetParam(AppLogic.ConfigParam.Server);
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {


            fillHousehold();
        }
    }
    public void fillHousehold()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("SP_S_HouseHoldName @IncludeClassB = 1, @AdvisorId = null, @BatchId = null, @AssociateId = null, @RecipientId = null, @BatchType =null");
        ddlHouseHold.Items.Clear();
        ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }
    protected void ddlHouseHold_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Visible = false;
        lblError.Text = "";
        lblMessage.Visible = false;
        lblMessage.Text = "";
        // string HouseholdTxt = ddlHouseHold.SelectedItem.ToString();
        object HouseholdTxt = ddlHouseHold.SelectedItem.ToString() == "Please Select" || ddlHouseHold.SelectedItem.ToString() == "" ? "null" : "'" + ddlHouseHold.SelectedItem.ToString() + "'";

        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("exec [dbo].[SP_S_RRGFamily] @HouseholdNameTxt =" + HouseholdTxt);
        ddlGAGroup.Items.Clear();
        //   ddlGAGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            //  string val = loDataset.Tables[0].Rows[liCounter][1].ToString();

            //  ddlGAGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
            ddlGAGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(liCounter)));
        }

        ddlTIAGroup.Items.Clear();
        // ddlTIAGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[1].Rows.Count; liCounter++)
        {
            // ddlTIAGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[1].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[1].Rows[liCounter][0])));
            ddlTIAGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[1].Rows[liCounter][0]), Convert.ToString(liCounter)));
        }
        if (loDataset.Tables[0].Rows.Count > 0 && loDataset.Tables[1].Rows.Count > 0)
        {
            btnSumbitTop.Enabled = true;
        }
        else
        {
            btnSumbitTop.Enabled = false;
        }
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        try
        {

            lblError.Text = "";
            lblMessage.Text = "";
            lblError.Visible = false;
            lblMessage.Visible = false;
            object HouseholdTxt = ddlHouseHold.SelectedItem.ToString() == "Please Select" || ddlHouseHold.SelectedItem.ToString() == "" ? "null" : "'" + ddlHouseHold.SelectedItem.ToString() + "'";
            object GATxt = ddlGAGroup.SelectedItem.ToString() == "Please Select" || ddlGAGroup.SelectedItem.ToString() == "" ? "null" : "'" + ddlGAGroup.SelectedItem.ToString() + "'";
            object TIATxt = ddlTIAGroup.SelectedItem.ToString() == "Please Select" || ddlTIAGroup.SelectedItem.ToString() == "" ? "null" : "'" + ddlTIAGroup.SelectedItem.ToString() + "'";

            DB clsDB = new DB();
            DataSet loDataset = clsDB.getDataSet("[dbo].SP_S_RRG_Data @HouseHoldNameTxt=" + HouseholdTxt + ", @FamilyGARRGNameTxt=" + GATxt + ", @FamilyTIARRGNameTxt=" + TIATxt + "");
            if (loDataset.Tables.Count > 0)
            {




                string ReportOpFolder = Request.MapPath("ExcelTemplate\\TempFolder\\");
                string FileName = "";

                DateTime dtAsOfDate = Convert.ToDateTime(DateTime.Now);
                string strYear = dtAsOfDate.Year.ToString().Length < 2 ? "0" + dtAsOfDate.Year.ToString() : dtAsOfDate.Year.ToString();
                string strMonth = dtAsOfDate.Month.ToString().Length < 2 ? "0" + dtAsOfDate.Month.ToString() : dtAsOfDate.Month.ToString();
                string strDay = dtAsOfDate.Day.ToString().Length < 2 ? "0" + dtAsOfDate.Day.ToString() : dtAsOfDate.Day.ToString();
                string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
                string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
                string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
                string strMilliSecond = DateTime.Now.Millisecond.ToString().Length < 2 ? "0" + DateTime.Now.Millisecond.ToString() : DateTime.Now.Millisecond.ToString();

                if (Server.ToLower() != "prod")//added 7_6_219
                {
                    FileName = GeneralMethods.RemoveSpecialCharacters(ddlHouseHold.SelectedItem.ToString()) + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond + "_TEST" + ".xlsx";
                }
                else
                {
                    FileName = GeneralMethods.RemoveSpecialCharacters(ddlHouseHold.SelectedItem.ToString()) + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond + ".xlsx";
                }




                // GeneralMethods.RemoveSpecialCharacters(HouseholdTxt.ToString());
                //Create Excel File
                if (!Directory.Exists(ReportOpFolder))
                {
                    Directory.CreateDirectory(ReportOpFolder);
                }

                String lsFileNamforFinalXls = FileName;
                string ExcelFilePath = ReportOpFolder + lsFileNamforFinalXls;
                bool Success = GenerateExcel(ExcelFilePath, loDataset);
                if (Success)
                {

                    //Response.Write("<script>");
                    //lsFileNamforFinalXls = ExcelFilePath;
                    //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    //Response.Write("</script>");
                    Response.Write("<script>");
                    lsFileNamforFinalXls = "./ExcelTemplate/TempFolder/" + lsFileNamforFinalXls;
                    Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                    Response.Write("</script>");


                    lblMessage.Visible = true;
                    lblMessage.Text = "Report Generated Successfully";
                }
            }
            else
            {
                lblError.Visible = true;
                lblError.Text = "No Records Found";
            }
        }
        catch (Exception exx)
        {
            lblError.Visible = true;
            lblError.Text = "Error Occurred" + exx.Message.ToString();
        }




    }

    public bool GenerateExcel(string ExcelFilePath, DataSet ds)
    {
        bool isSuccess = false;
        try
        {
            #region Spire License Code
            string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
            Spire.License.LicenseProvider.SetLicenseKey(License);
            Spire.License.LicenseProvider.LoadLicense();
            #endregion
            if (System.IO.File.Exists(ExcelFilePath))
            {
                System.IO.File.Delete(ExcelFilePath);
            }
            int SheetNo = 0;
            Workbook book = new Workbook();
            book.Version = ExcelVersion.Version2016;

            string TemplateFile = Request.MapPath("ExcelTemplate\\") + "RRGTemplate.xlsx";
            // book.CreateEmptySheets(ds.Tables.Count);
            book.LoadFromFile(TemplateFile);
            for (int i = 0; i < ds.Tables.Count; i++)
            {

                //  string SheetNme = ds.Tables[i].Rows[0][0].ToString();
                // string GroupName = ds.Tables[i].Rows[0][0].ToString();
                // i++;
                // ds.Tables[i].Columns.Add("GroupName");



                Worksheet sheet = book.Worksheets[SheetNo];
                // sheet.Name = SheetNme;
                if (ds.Tables[i].Rows.Count > 0)
                {
                    // sheet.Range[1, 1, 1, ds.Tables[i].Columns.Count].Style.Font.IsBold = true;
                    if (i == 1)
                    {
                        sheet.InsertDataTable(ds.Tables[i], false, 3, 1);
                    }
                    else if (i == 7)
                    {
                        sheet.InsertDataTable(ds.Tables[i], true, 4, 1);
                    }
                    else if (i == 8)
                    {
                        sheet.InsertDataTable(ds.Tables[i], false, 5, 1);
                    }
                    else if (i == 9)
                    {
                        sheet.InsertDataTable(ds.Tables[i], true, 4, 2);
                    }
                    else if (i == 10)
                    {
                        sheet.InsertDataTable(ds.Tables[i], true, 1, 3);
                    }
                    else if (i == 11)
                    {
                        sheet.InsertDataTable(ds.Tables[i], false, 3, 1);
                    }
                    else
                    {
                        sheet.InsertDataTable(ds.Tables[i], false, 2, 1);
                    }

                    ///  sheet.Range[1, 1, ds.Tables[i].Rows.Count, ds.Tables[i].Columns.Count].AutoFitColumns();
                    //  sheet.Range[1, 1, ds.Tables[i].Rows.Count, ds.Tables[i].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;
                }
                if (i != 10)
                {
                    SheetNo++;
                }
            }

            // book.SaveToFile(ExcelFilePath, ExcelVersion.Version2016);
            book.SaveToFile(ExcelFilePath);
            isSuccess = true;
        }
        catch (Exception ex)
        {
            lblError.Visible = true;
            lblError.Text = "Error Occurred" + ex.Message.ToString();
        }
        return isSuccess;
    }
}