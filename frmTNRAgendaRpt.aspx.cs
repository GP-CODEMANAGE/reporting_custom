using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using Microsoft.Reporting.WebForms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Spire.Xls;

public partial class frmTNRAgendaRpt : System.Web.UI.Page
{

    object StartDate = null;
    object EndDate = null;
    object ShowPdf = null;
    object FundGrp = null;
    GeneralMethods clsGM = new GeneralMethods();
    string UUID = "";

    //string strServerURL = "http://gp-crmsql1/ReportServer";//Gp-crm2016
    //string strServerURL = "http://gp-db1/ReportServer";//Gp-crm1
    //string strServerURL = "http://gp-crmsql1/ReportServer";//Gp-PRODDB // added 4_17_2019
    string strServerURL = "http://gp-testdb/ReportServer_GPTESTDB/";//Gp-TESTCRM // added 6_11_2019

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            trexception.Style.Add("display", "none");
            BindFundType();
            //  Session["ExceptionReport"] = "";
            //if (Convert.ToString(Session["ExceptionReport"]) == "" && txtStartDate.Text != "" && txtendDate.Text != "")
            //    trexception.Style.Add("display", "none");

            //if (Convert.ToString(Session["ExceptionReport"]) != "" && txtStartDate.Text == "" && txtendDate.Text == "")
            //{
            //    Session["ExceptionReport"] = "";
            //    trexception.Style.Add("display", "none");
            //}

            //if (Convert.ToString(Session["ExceptionReport"]) != "")
            //    trexception.Style.Add("display", "");
            lstFund.SelectedIndex = 0;

        }

    }

    public void BindFundType()
    {

        string sqlstr = "SP_S_FUND_PARTNERSHIPTYPE";
        clsGM.getListForBindListBox(lstFund, sqlstr, "TypeNametxt", "TypeIdNmb");

        lstFund.Items.Insert(0, "All");
        lstFund.Items[0].Value = "0";

    }

    protected void btnGeneratePDF_Click(object sender, EventArgs e)
    {
        try
        {
            lblMessage.Text = "";
            Warning[] warnings;
            string[] streamids;
            string mimeType;
            string encoding;
            string extension;
            DB clsDB = new DB();
            string Template1 = "";

            string strGUID = DateTime.Now.ToString("MMddyyhhmmss");

            //////// Set Header Title details for PDF  ////////        
            string strStartdate = string.Empty;
            string strEndDate = string.Empty;
            string strShowPdf = string.Empty;


            UUID = Guid.NewGuid().ToString();
            string lsSQL = getFinalSp(0, UUID);
            //Response.Write(lsSQL);
            DataSet newdataset = clsDB.getDataSet(lsSQL);

            if (newdataset.Tables.Count > 0)
            {
                if (newdataset.Tables[1].Rows.Count > 0)
                {
                    if (Convert.ToInt16(Convert.ToString(newdataset.Tables[1].Rows[0]["CountInserted"])) == 0)
                    {
                        lblMessage.ForeColor = System.Drawing.Color.Red;
                        lblMessage.Text = "No Records were found";
                        return;
                    }
                }
                //Response.Write("Count " + newdataset.Tables[1].Rows.Count);
                if (newdataset.Tables[0].Rows.Count > 0)
                {
                    string ExcelfileName = "Exception_Report_" + txtStartDate.Text.Replace("/", "").ToString() + "-" + txtendDate.Text.ToString().Replace("/", "") + ".xls";
                    Session["ExceptionReport"] = ExcelfileName.ToString();

                    trexception.Style.Add("display", "");
                    ExportExcel(newdataset.Tables[0], ExcelfileName);
                }
                else
                {
                    trexception.Style.Add("display", "none");
                    Session["ExceptionReport"] = "";
                }
            }

            if (RadioButtonList1.SelectedValue == "1") //PDF
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Test", "showexcrpt();", true);
                Random rnd = new Random();
                int rannum = rnd.Next(1, 9999);
                //Fetch Values from UI controls
                getValues();

                ReportParameter[] param = new ReportParameter[6];
                param[0] = new ReportParameter("Start_Date", Convert.ToString(StartDate));
                param[1] = new ReportParameter("End_Date", Convert.ToString(EndDate));
                param[2] = new ReportParameter("UUID", Convert.ToString(UUID));
                param[3] = new ReportParameter("Fund_Group", Convert.ToString(FundGrp));
                param[4] = new ReportParameter("DocType", Convert.ToString(1));
                param[5] = new ReportParameter("PDF", Convert.ToString(1));


                ReportViewer viewer = new ReportViewer();

                viewer.ProcessingMode = ProcessingMode.Remote;
                viewer.ServerReport.ReportServerCredentials = new ReportServerNetworkCredentials();

                viewer.ServerReport.ReportServerUrl = new Uri(strServerURL); //report server for TEST SERVER

                viewer.ServerReport.ReportPath = "/TNR Reports/PCA Report 2"; // rdl name

                viewer.ServerReport.SetParameters(param);
                viewer.ServerReport.Refresh();
                //viewer.LocalReport.DataSources.Add(rds);

                byte[] bytes = viewer.ServerReport.Render("PDF", null, out mimeType, out encoding, out extension, out streamids, out warnings);


                if (bytes != null)
                {

                    Response.AddHeader("content-disposition", "attachment;filename= " + strGUID + "_TNR_Agenda_Report.pdf");
                    Response.ContentType = "application/octectstream";
                    Response.BinaryWrite(bytes);
                    Response.End();
                    // this.Context.ApplicationInstance.CompleteRequest();
                }
                // ViewState["DownloadData"] = bytes;
                // Page.ClientScript.RegisterStartupScript(this.GetType(), "Test", "ShowAlert();", true);




            }
            else
            {
                Page.ClientScript.RegisterStartupScript(this.GetType(), "Test", "showexcrpt();", true);
                Random rnd = new Random();
                int rannum = rnd.Next(1, 9999);
                //Fetch Values from UI controls
                getValues();

                ReportParameter[] param = new ReportParameter[5];
                param[0] = new ReportParameter("MasterStartDate", Convert.ToString(StartDate));
                param[1] = new ReportParameter("MasterEndDate", Convert.ToString(EndDate));
                param[2] = new ReportParameter("MasterUUID", Convert.ToString(UUID));
                param[3] = new ReportParameter("MasterFund_Group", Convert.ToString(FundGrp));
                param[4] = new ReportParameter("MasterDocType", Convert.ToString(2));


                ReportViewer viewer = new ReportViewer();

                viewer.ProcessingMode = ProcessingMode.Remote;
                viewer.ServerReport.ReportServerCredentials = new ReportServerNetworkCredentials();

                viewer.ServerReport.ReportServerUrl = new Uri(strServerURL); //report server for TEST SERVER

                viewer.ServerReport.ReportPath = "/TNR Reports/Master Report"; // rdl name

                viewer.ServerReport.SetParameters(param);
                viewer.ServerReport.Refresh();
                //viewer.LocalReport.DataSources.Add(rds);

                byte[] bytes = viewer.ServerReport.Render("Excel", null, out mimeType, out encoding, out extension, out streamids, out warnings);
                // if (Convert.ToString(Session["ExceptionReport"]) != "")
                //this.Page.ClientScript.RegisterStartupScript(this.GetType(), "alert", "showexcrpt();", true);
                //Response.Write("<script>showexcrpt()</script>");
                // string outputPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\TNR_Agenda_Report_" + strGUID + "_" + rannum + ".xls";

                //using (FileStream fs = new FileStream(outputPath, FileMode.Create))
                //{
                //    fs.Write(bytes, 0, bytes.Length);
                //    fs.Close();
                //}
                //Session["OUTPUTFILE"] = outputPath;

                //    ScriptManager.RegisterStartupScript(this, typeof(string), "OPEN_WINDOW", "var Mleft = (screen.width/2)-(760/2);var Mtop = (screen.height/2)-(700/2);window.open( 'file.aspx', null, 'height=700,width=760,status=yes,toolbar=no,scrollbars=yes,menubar=no,location=no,top=\'+Mtop+\', left=\'+Mleft+\'' );", true);

                //  byte[] bytesPDF = System.IO.File.ReadAllBytes(outputPath);

                if (bytes != null)
                {

                    Response.AddHeader("content-disposition", "attachment;filename= " + strGUID + "_TNR_Agenda_Report.xls");
                    Response.ContentType = "application/vnd.ms-excel";
                    Response.BinaryWrite(bytes);
                    Response.End();
                    // this.Context.ApplicationInstance.CompleteRequest();
                }


            }

            #region Commented
            //if (RadioButtonList1.SelectedValue == "1") // for Pdf files
            //{

            //    string[] strTemplate = null;
            //    string DestinationPath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\TNR_Agenda_Report_" + strGUID + ".pdf";
            //    string FinalPath = "./ExcelTemplate/pdfOutput/TNR_Agenda_Report_" + strGUID + ".pdf";

            //    Template1 = Template1.Substring(1, Template1.Length - 1);
            //    strTemplate = Template1.Split('|');
            //    if (DestinationPath != "")
            //    {
            //        PDFMerge PDF = new PDFMerge();
            //        PDF.MergeFiles(DestinationPath, strTemplate);
            //    }

            //    byte[] bytesPDF = System.IO.File.ReadAllBytes(DestinationPath);

            //    if (bytesPDF != null)
            //    {

            //        Response.AddHeader("content-disposition", "attachment;filename= " + strGUID + "_TNR_Agenda_Report.pdf");
            //        Response.ContentType = "application/octectstream";
            //        Response.BinaryWrite(bytesPDF);
            //        Response.End();
            //    }

            //    //  lnkException.Visible = true;
            //}
            //else
            //{

            //    #region Commented
            //    //String lsFileNamforFinalXls = System.DateTime.Now.ToString("MMddyyhhmmss") + ".xls";
            //    //string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\pdfOutput\fffffff.xls");
            //    //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + lsFileNamforFinalXls);

            //    //load the first workbook
            //    //  Workbook workbook = new Workbook();
            //    //  workbook.LoadFromFile(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\TNR_Agenda_Report_051616083333_8081_temp.xls", ExcelVersion.Version97to2003);

            //    //load the second workbook
            //    //   Workbook workbook2 = new Workbook(); 
            //    //   workbook2.LoadFromFile(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\TNR_Agenda_Report_051616083333_2639_temp.xls", ExcelVersion.Version97to2003);

            //    //insert the second workbook's worksheet into the first workbook using a dataTable
            //    //  Worksheet sheet2 = workbook2.Worksheets[0];
            //    // DataTable dataTable = sheet2.ExportDataTable();
            //    // Worksheet sheetAdd = workbook.CreateEmptySheet("TEST"); 
            //    // sheetAdd.InsertDataTable(dataTable, true, 1, 1);
            //    // Worksheet newsheet = (Worksheet)sheet2.Clone(sheet2.Parent);
            //    // workbook.Worksheets.Add(sheet2);
            //    //  sheet2.Range["B02:L02"].Style.Color = Color.Blue;
            //    //   sheet2.Range["B02:L02"].Style.Font.Color = Color.White;
            //    //    workbook.Worksheets.AddCopyAfter(sheet2);
            //    //    workbook.Worksheets.Insert(1, sheet2);
            //    // workbook.Worksheets.Add("copied sheet1");
            //    //   workbook.Worksheets[0].CopyFrom(workbook2.Worksheets[0]);
            //    //save the workbook
            //    //  workbook.SaveToFile(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\FilesMerge4.xls", ExcelVersion.Version97to2003);

            //    //launch the workbook
            //    //  System.Diagnostics.Process.Start(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\FilesMerge2.xls");
            //    // System.Diagnostics.Process.Start(DestinationPath);


            //    //Workbook workbook = new Workbook();
            //    //workbook.LoadFromFile(@"..\copy worksheets.xls");
            //    //Worksheet worksheet = workbook.Worksheets[0];
            //    ////add worksheets and name them
            //    //workbook.Worksheets.Add("copied sheet1");
            //    //workbook.Worksheets.Add("copied sheet2");
            //    ////copy worksheet to the new added worksheets
            //    //workbook.Worksheets[1].CopyFrom(workbook.Worksheets[0]);
            //    //workbook.Worksheets[2].CopyFrom(workbook.Worksheets[0]);
            //    //workbook.SaveToFile(@"..\copy worksheets.xls");
            //    //System.Diagnostics.Process.Start(@"..\copy worksheets.xls");


            //    //byte[] bytesPDF = System.IO.File.ReadAllBytes(excel1);

            //    //if (bytesPDF != null)
            //    //{

            //    //    Response.AddHeader("content-disposition", "attachment;filename= " + strGUID + "_TNR_Agenda_Report1.xls");
            //    //    Response.ContentType = "application/vnd.ms-excel";
            //    //    Response.BinaryWrite(bytesPDF);
            //    //    Response.End();
            //    //}

            //    //byte[] bytesPDF1 = System.IO.File.ReadAllBytes(excel2);

            //    //if (bytesPDF1 != null)
            //    //{

            //    //    Response.AddHeader("content-disposition", "attachment;filename= " + strGUID + "_TNR_Agenda_Report2.xls");
            //    //    Response.ContentType = "application/vnd.ms-excel";
            //    //    Response.BinaryWrite(bytesPDF1);
            //    //    Response.End();
            //    //}

            //    #endregion

            //}
            #endregion


        }
        catch (Exception ex)
        {
            lblMessage.ForeColor = System.Drawing.Color.Red;
            lblMessage.Text = "Error Occured while generating the report. Details: " + ex.Message;
        }
    }

    public void getValues()
    {

        if (txtStartDate.Text == "")
            StartDate = "null";
        else
            StartDate = txtStartDate.Text;

        if (txtendDate.Text == "")
            EndDate = "null";
        else
            EndDate = txtendDate.Text;

        if (RadioButtonList1.SelectedValue == "1")
            ShowPdf = "1";
        else
            ShowPdf = "0";

        //if (ddlFund.SelectedValue == "0")
        //    FundGrp = "null";
        //else
        //    FundGrp = ddlFund.SelectedValue;


        if (lstFund.SelectedIndex == 0)
            FundGrp = "null";
        else
        {
            FundGrp = GetMultipleSelectedItemsFromListBox(lstFund);
        }


    }

    public string GetMultipleSelectedItemsFromListBox(ListBox lstBox)
    {
        string lstselecteditems = "";
        if (lstBox.Items.Count > 0)
        {
            for (int i = 0; i < lstBox.Items.Count; i++)
            {
                if (lstBox.Items[i].Selected)
                {
                    lstselecteditems = lstselecteditems + "|" + lstBox.Items[i].Value;

                }
            }
            if (lstselecteditems != "")
            {
                lstselecteditems = lstselecteditems.Substring(1);
            }
        }


        return lstselecteditems;
    }

    public string getFinalSp(int flg, string UUID)
    {

        String lsSQL = "";

        string StartDt = txtStartDate.Text == "" ? "null" : "'" + txtStartDate.Text + "'";
        string Enddt = txtendDate.Text == "" ? "null" : "'" + txtendDate.Text + "'";
        string Fund = "";
        if (lstFund.SelectedIndex == 0)
            Fund = "null";
        else
        {
            Fund = "'" + GetMultipleSelectedItemsFromListBox(lstFund) + "'";
        }

        lsSQL = "SP_R_TNR_PRICING_COMMITTEE_AGENDA_REPORT @Start_Date=" + StartDt +
                                                        ",@End_Date=" + Enddt +
                                                        ",@Fund_Group=" + Fund +
                                                         ",@DocType=1" +
                                                        ",@UUID='" + UUID +
                                                        "',@PDF=" + flg + "";


        return lsSQL;
    }

    private void ExportExcel(DataTable p_dsSrc, string p_strPath)
    {

        using (ExcelPackage objExcelPackage = new ExcelPackage())
        {

            //Create the worksheet    
            ExcelWorksheet objWorksheet = objExcelPackage.Workbook.Worksheets.Add("Sheet1");
            //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1    
            objWorksheet.Cells["A1"].LoadFromDataTable(p_dsSrc, true);
            objWorksheet.Cells.Style.Font.SetFromFont(new Font("Calibri", 10));
            objWorksheet.Cells.AutoFitColumns();
            //Format the header    
            using (ExcelRange objRange = objWorksheet.Cells["A1:XFD1"])
            {
                objRange.Style.Font.Bold = true;
                objRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                objRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                objRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                objRange.Style.Fill.BackgroundColor.SetColor(Color.White);
            }

            //Write it back to the client    
            if (File.Exists(@"" + Server.MapPath("./ExcelTemplate/" + p_strPath + "") + ""))
            {
                File.Delete(@"" + Server.MapPath("./ExcelTemplate/" + p_strPath + "") + "");
            }

            //Create excel file on physical disk    
            FileStream objFileStrm = File.Create(Server.MapPath("./ExcelTemplate/" + p_strPath + ""));
            objFileStrm.Close();

            //Write content to excel file    
            File.WriteAllBytes(Server.MapPath("./ExcelTemplate/" + p_strPath + ""), objExcelPackage.GetAsByteArray());

        }
    }


    protected void lnkPCAReport_Click(object sender, EventArgs e)
    {
        String Filename = Convert.ToString(ViewState["PCAReport"]);
        string filePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + Filename + "";
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + Filename + "");
        Response.TransmitFile(filePath);
        Response.End();
    }
    protected void lnkPCAReport2_Click(object sender, EventArgs e)
    {
        String Filename = Convert.ToString(ViewState["PCAReport2"]);
        string filePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + Filename + "";
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + Filename + "");
        Response.TransmitFile(filePath);
        Response.End();
    }
    protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        //  lnkException.Visible = false;
        trexception.Style.Add("display", "none");
        //if (RadioButtonList1.SelectedValue == "1")
        //{
        //    lnkPCAReport.Visible = false;
        //    lnkPCAReport2.Visible = false;
        //}

    }
    protected void lnkException_Click(object sender, EventArgs e)
    {
        String Filename = Convert.ToString(Session["ExceptionReport"]);
        string filePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\" + Filename + "";
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + Filename + "");
        Response.TransmitFile(filePath);
        Response.End();
    }

    //protected void btnHidden_Click(object sender, EventArgs e)
    //{

    //  //  byte[] certificateBytes = (byte[])ViewState["DownloadData"];
    //    string filePath = Convert.ToString(Session["OUTPUTFILE"]);
    //    Response.ContentType = "application/octet-stream";
    //    Response.AddHeader("Content-Disposition", "attachment; filename=Test.pdf");
    //    Response.TransmitFile(filePath);
    //    Response.End();

    //}
}