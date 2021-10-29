using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using Spire.Xls;

public partial class Develpment : System.Web.UI.Page
{
    Logs lg = new Logs();
    public StreamWriter sw = null;

    protected void Page_Load(object sender, EventArgs e)
    {
        // ReadFile();
    }
    public string[] ReadFile()
    {
        string[] lines = null;
        try
        {
            string Filepath = Server.MapPath("") + @"\File.txt";

            var list = new List<string>();
            // var fileStream = new FileStream(@"D:\Back Data 19-08-2016\D Data\Development\GreshamPartners\Reporting Custom\Check\file.txt", FileMode.Open, FileAccess.Read);
            var fileStream = new FileStream(Filepath, FileMode.Open, FileAccess.Read);
            using (var streamReader = new StreamReader(fileStream, System.Text.Encoding.UTF8))
            {
                string line;
                while ((line = streamReader.ReadLine()) != null)
                {
                    list.Add(line);
                }
            }
            lines = list.ToArray();
        }
        catch (Exception ex)
        {
            lg.AddinLogFile(Session["Filename"].ToString(), "Error Occureed :" + ex.Message.ToString() + "---" + DateTime.Now);
        }
        return lines;
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        Random rnd = new Random();
        string strRndNumber = Convert.ToString(rnd.Next(9999));
        DateTime dtmain = DateTime.Now;
        string LogFileName = string.Empty;
        LogFileName = "Log_" + strRndNumber + "_" + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        sw = new StreamWriter(LogFileName);
        sw.Close();
        HttpContext.Current.Session["Filename"] = LogFileName;
        ViewState["Filename"] = LogFileName;


        LogFileName = (string)ViewState["Filename"];

        Session["Filename"] = LogFileName;

        string filenale = (string)Session["Filename"];

        lg.AddinLogFile(Session["Filename"].ToString(), "----------------------------------PROCESS START------------------------- " + dtmain);

        Process();
    }
    public void Process()
    {
        Dictionary<string, string> found = new Dictionary<string, string>();
        string line;
        bool textFound = false;
        int FileFound = 0;
        int TextFound_in_file = 0;
        //int Total_Proc = 0;
        DataTable dt = new DataTable();
        dt.Columns.Add("ProcName");
        dt.Columns.Add("FileName");
        dt.Columns.Add("Line");

        DataRow dr = null;
        lg.AddinLogFile(Session["Filename"].ToString(), "ReadFile " + "---" + DateTime.Now);

        try
        {


            string[] ProcName = ReadFile();

            if (ProcName.Length > 0)
            {
                lg.AddinLogFile(Session["Filename"].ToString(), "ProcName.Length " + ProcName.Length + "---" + DateTime.Now);

                foreach (string ProcName_ in ProcName)
                {
                    //  Total_Proc++;

                    foreach (string filename in Directory.GetFiles(@"D:\Back Data 19-08-2016\D Data\Development\GreshamPartners\Reporting Custom\Check\", "*", SearchOption.AllDirectories))
                    {
                        lg.AddinLogFile(Session["Filename"].ToString(), "File Checked: " + filename + "---" + DateTime.Now);
                        FileFound++;
                        using (StreamReader file = new StreamReader(filename))
                        {
                            while ((line = file.ReadLine()) != null)
                            {
                                //Check if ProcName exist in each file.
                                if (line.Contains(ProcName_))
                                {
                                    found.Add(line, filename);

                                    #region Add Filename and line to Table 
                                    dr = dt.NewRow();
                                    dr["ProcName"] = Path.GetFileName(ProcName_);
                                    dr["FileName"] = Path.GetFileName(filename);
                                    dr["Line"] = line;
                                    dt.Rows.Add(dr);
                                    #endregion

                                    textFound = true;
                                    TextFound_in_file++;
                                }

                            }
                        }
                    }
                }
            }

            if (!textFound)
            {
                lblError.Text = "No Words Found" + "File Count = " + TextFound_in_file;
            }
            else
            {
                lblError.Text = "File Count = " + TextFound_in_file;

                string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
                string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
                string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
                string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
                string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
                string strMilliSecond = DateTime.Now.Millisecond.ToString().Length < 2 ? "0" + DateTime.Now.Millisecond.ToString() : DateTime.Now.Millisecond.ToString();

                string ExcelFileName = strYear + "_" + strMonth + "_" + strDay + strHour + "_" + strMinute + "_" + strSecond + "_" + strMilliSecond + "_TEST" + ".xlsx";
                string Excel_FilePath = GenerateExcel(ExcelFileName, dt);
                lg.AddinLogFile(Session["Filename"].ToString(), "Excel_FilePath: " + Excel_FilePath + "---" + DateTime.Now);
                if (Excel_FilePath != "")
                {

                    //Response.Write("<script>");
                    //Response.Write("window.open('" + Excel_FilePath + "', 'mywindow')");
                    //Response.Write("</script>");

                    //string baseUrl = Request.Url.GetLeftPart(UriPartial.Authority);
                    //Response.Redirect(baseUrl + Excel_FilePath);


                    Response.ContentType = "application/octet-stream";
                    Response.AddHeader("Content-Disposition", "attachment;filename=" + Excel_FilePath);
                    Response.TransmitFile(Server.MapPath(Excel_FilePath));
                    Response.End();
                }
            }
        }
        catch (Exception ex)
        {
            lblError.Text = "Error in PROCESS: " + ex.Message.ToString();
            lg.AddinLogFile(Session["Filename"].ToString(), "Error in PROCESS: " + ex.Message.ToString()+ "---" + DateTime.Now);

        }
    }

    public string GenerateExcel(string FileName, DataTable dt)
    {
        #region Spire License Code
        //string License = ConfigurationSettings.AppSettings["SpireLicense"].ToString();
        //Spire.License.LicenseProvider.SetLicenseKey(License);
        //Spire.License.LicenseProvider.LoadLicense();
        #endregion

        try
        {
            if (!Directory.Exists(Server.MapPath("") + "\\ExcelTemplate\\TempFolder\\"))
            {
                Directory.CreateDirectory(Server.MapPath("") + "\\ExcelTemplate\\TempFolder\\");
            }

            //string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            //string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            //string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            String lsFileNamforFinalXls = FileName;// "ManagerPerformance_" + strMonth + "_" + strDay + "_" + strYear + ".xlsx";
            string ExcelFilePath = Server.MapPath("") + "\\ExcelTemplate\\TempFolder\\" + lsFileNamforFinalXls;

            if (System.IO.File.Exists(ExcelFilePath))
            {
                System.IO.File.Delete(ExcelFilePath);
            }




            int SheetNo = 0;
            Workbook book = new Workbook();
            // book.CreateEmptySheets(ds.Tables.Count / 2);
            //for (int i = 0; i < ds.Tables.Count; i++)
            //{


            Worksheet sheet = book.Worksheets[SheetNo];
            sheet.Name = "DataFetched";
            if (dt.Rows.Count > 0)
            {
                sheet.Range[1, 1, 1, dt.Columns.Count].Style.Font.IsBold = true;

                sheet.InsertDataTable(dt, true, 1, 1);
                sheet.Range[1, 1, dt.Rows.Count, dt.Columns.Count].AutoFitColumns();
                sheet.Range[1, 1, dt.Rows.Count, dt.Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;
            }
            SheetNo++;
            //}

            book.SaveToFile(ExcelFilePath, ExcelVersion.Version2016);

            string vContain = "Excel Report Generated Succesfully ";
            //sw.WriteLine(vContain);
            return ExcelFilePath;
        }
        catch (Exception e)
        {

            string vContain = "Excel Report Genration Fail,  Error " + e.ToString();
            //sw.WriteLine(vContain);
            //LG.AddinLogFile(Form1.vLogFile, vContain);
            //   lblMsg.Text = vContain;

            return "";
        }
    }
}