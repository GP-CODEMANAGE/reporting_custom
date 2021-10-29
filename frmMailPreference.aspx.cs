using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class frmTaskNote : System.Web.UI.Page
{
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            BindHouseHold(lstHouseHold);
            BindAssociate(ddlAssociate);
            BindMailType(lstMailType);
        }
    }

    protected void ddlAssociate_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindHouseHold(lstHouseHold);
    }

    public void BindHouseHold(ListBox lstBox)
    {
        object AssociatedId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";

        lstBox.Items.Clear();

        //String sqlstr = "SP_S_HouseHoldName @IncludeClassB = 1,@AdvisorId=" + AdvisorId + ",@BatchId=" + BatchId + ",@AssociateId=" + AssociatedId + ",@RecipientId=" + RecipientId + ",@BatchType=" + BatchType;

        String sqlstr = "SP_S_HouseHoldName @IncludeClassB = 1,@AssociateId=" + AssociatedId;
        clsGM.getListForBindListBox(lstBox, sqlstr, "Name", "Accountid");

        if (lstBox.Items.Count == 1)
        {
            if (lstBox.Items[0].Value == "0")
                lstBox.Items.Remove(lstBox.Items[0]);
        }
        lstBox.Items.Insert(0, "All");
        lstBox.Items[0].Value = "0";
        lstBox.SelectedIndex = 0;

    }

    public void BindMailType(ListBox lstBox)
    {
        object AssociatedId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";

        lstBox.Items.Clear();

        //String sqlstr = "SP_S_HouseHoldName @IncludeClassB = 1,@AdvisorId=" + AdvisorId + ",@BatchId=" + BatchId + ",@AssociateId=" + AssociatedId + ",@RecipientId=" + RecipientId + ",@BatchType=" + BatchType;

        String sqlstr = "SP_S_MAILTYPE";
        clsGM.getListForBindListBox(lstBox, sqlstr, "ssi_name", "ssi_mailid");

        if (lstBox.Items.Count == 1)
        {
            if (lstBox.Items[0].Value == "0")
                lstBox.Items.Remove(lstBox.Items[0]);
        }
        lstBox.Items.Insert(0, "All");
        lstBox.Items[0].Value = "0";
        lstBox.SelectedIndex = 0;

    }

    public void BindAssociate(DropDownList ddl)
    {
        object OwnerId = "null";
        ddl.Items.Clear();

        string sqlstr = "SP_S_ASSOCIATE @OwnerId=" + OwnerId;//;///SP_S_BATCH_ASSOCIATE//
        clsGM.getListForBindDDL(ddl, sqlstr, "Ssi_SecondaryOwnerIdName", "Ssi_SecondaryOwnerId");

        if (ddl.Items.Count == 1)
        {
            if (ddl.Items[0].Value == "0")
                ddl.Items.Remove(ddl.Items[0]);
        }
        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    protected void btnsubmit_Click(object sender, EventArgs e)
    {
        object AssociateId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object HouseHold = lstHouseHold.SelectedValue == "0" || lstHouseHold.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstHouseHold) + "'";
        object MailType = lstMailType.SelectedValue == "0" || lstMailType.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstMailType) + "'";

        string str = "SP_R_MailingList_Preference @AssociateUUID=" + AssociateId + ",@HouseHoldUUIDListTxt=" + HouseHold + ",@MailTypeUUIDListTxt=" + MailType;

        #region Spire License Code 

        string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();

        #endregion

        Workbook book = new Workbook();
        Worksheet sheet = book.Worksheets[0];

        DataSet dataSet = clsDB.getDataSet(str);
        DataTable datatable = dataSet.Tables[0];

        ExcelFont fontBold = book.CreateFont();
        fontBold.IsBold = true;

        sheet.InsertDataTable(datatable, true, 1, 1);

        sheet.AllocatedRange.AutoFitColumns();
        sheet.AllocatedRange.AutoFitRows();

        sheet.Range["A1:E1"].Style.Font.IsBold = true;

        sheet.Range["A1:E1"].Style.HorizontalAlignment = HorizontalAlignType.Center;


        string ExcelFilePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\";
        string FileName = "MailPreferenceReportExcel_"+ DateTime.Now.ToString("yyyy_MM_dd_hhmmss") + ".xlsx";
        string fn = FileName;
        book.SaveToFile(ExcelFilePath + FileName, ExcelVersion.Version2016);
        // System.Diagnostics.Process.Start(ExcelFilePath + FileName);

        FileInfo file = new FileInfo(ExcelFilePath + FileName);

        if (file.Exists)
        {
            Response.Clear();
            Response.AddHeader("Content-Disposition", "attachment; filename=" + file.Name);
            Response.AddHeader("Content-Length", file.Length.ToString());
            Response.ContentType = "Application/x-msexcel";
            Response.Flush();
            Response.TransmitFile(file.FullName);
            Response.End();
        }
        else
        {
            lblError.Text = "Requested file is not available to download";
        }
    }
}