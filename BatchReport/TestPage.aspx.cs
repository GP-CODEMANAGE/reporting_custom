using iTextSharp.text;
using iTextSharp.text.pdf;
using java.io;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class BatchReport_TestPage : System.Web.UI.Page
{
    PDFMerge PDF = new PDFMerge();
    protected void Page_Load(object sender, EventArgs e)
    {
        string Gresham_String1 = "Password=slater6;Persist Security Info=False;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=gp-proddb";
        SqlConnection Gresham_con1 = new SqlConnection(Gresham_String1);
        DataSet ds_gresham1 = new DataSet();
        SqlDataAdapter dagreshamFund = new SqlDataAdapter();
        DataSet ds_greshamFund = new DataSet();
        DataSet ds1 = new DataSet();

        string greshamquery1 = string.Empty;
        greshamquery1 = "EXEC SP_S_FundList_SP";
        dagreshamFund = new SqlDataAdapter(greshamquery1, Gresham_con1);
        ds_greshamFund = new DataSet();
        dagreshamFund.Fill(ds_greshamFund);

        // string FundFilePath = GenerateCsv(ds_greshamFund, true);
        string FundFilePath2 = GenerateExcel("test.xlsx", ds_greshamFund);

    }
    public string GenerateExcel(string FileName, DataSet ds)
    {
        #region Spire License Code
        string License = "RXBfKi8BALJLc4hMAXEUQZAUbU2CX1R+LkUf5JIXi7aZue6hXI5ljcKLtYiQe0Y8fO6LNtmhAAol7LvqkdFm3IiJNsy7tKOrHCXBlMyJ8w00NrFqplYOV1aExL/fDDPJoRh31RzsMFiP/gM8i+uXonNjW9Q+5UGjYr0vMzWE8L6XLtTp65NswSYdO8kYV95m+QOOKfL8JvWOAZr5V4GCCPo4/ESVGW8O5Ciy3XkaM5JrzT+SF4gE+dyeu1DunZJMZf6gashLnOkQS44xLGqQvDTi+w6Yh1xp3sBlCfd0oeW9c2hCVzb8OsYSj5Zfl9oik1ia9y2z58y1fl+NfKKwJSiIkXY2dmwvm44zsyJ+CAhOhh63oiR3Ju5l1riG7ZO7AuGM8twEyUfjETAzfoLWnTtbh2u/KopSFHjpMh3f3NbgfPsvSiRE+UYWGk4gag3ZyLDdcxKnFyADu8NTkyrHnmT+gD1DKiOxbIOQ4S+tq1Ptkv7Lr0iVSgOriwn45p8J3D4mFirXXynOQ6c6RqdN72XGetut2WA7MX/c1YOJHAkwbAH3quB2rrIuOyJIf8H9caTppOhqS7Fj6XwoPxzzT/rE3P+bgjn89l6977+rsCebMLoZkNwAhatVhu+IdvHYhZOsnw64EfeccCVPzwsY9oVgOPmQSyf1+k7aE2AyaFeZmj+ZkIryU+lMjTyaBY/VBgkdhi8Fb05dw217vBoWaECSxu4D6H6ml4ymodegICj2pwFFA//tuwwLUsSVm0XPWfdG4KR6GZjiClE+eNmNJwH9tpmW6gVmqFCdE9h2b3U6FZX+DHwJju3oGlBubz0egnArXxCj/34xXBE9SSYwIyGgTCyLnrho2zZeIe7xZAp3XS2zIC2LJbO10AiOlgChpues7gutpsddbzyr7adW1d8e9L92b26LIifYFQtyX+rFFOUr88B425QINj09P3HzAB2PHXCAW5P6EaS8abqLTYldhl7J8cIUHA0HmWyGNBh1Wv522e4Wz+/XknxEEia3y8YhQmvmfKiTXSu1RdP7dLYgLDP4IHhbeIO+Wx1YSmQ86SZKmqBhwM6vgZnwmoQ4BWO6fXtZK/h0pyp4WlxvYLShuHXN+uuJnFhSKuDrG7S7Qt+yEMCWgjwF06Jx3k4lchKmzpSm8r5GKETQ+0prX9WynirB7GPx0iRrAJtS1aJK0nca3lX12PqZdw3FkN6b9yRg+fugaNZkfluwNVnXEXeifwJOXWRp0UmGfKUrqlNoQDxKEbVzSYHJn9czXC8shbJavgRwUmAOUNkYhTBqrj4BBxeBpB5km6R/zrfMnUPJ0mMVGkxVxE2ivvuaVzrDITWmBx2STeMkAs4E8frg2vBNXeAP22yo+ho9URfvDS7Itq8s3XJ4zuiLNByWYdtTFAKWD9SzqvTfpKKwifp+Sl4upfC6gGylh3Tzyc2KD4XdANT+rlZ2VfpBQlo7DHMPKPgNB37WP4OSNEq1viV8U9JmSrCybZFbqd8+v1h3ygnUOR5wusbC5Q==";
        Spire.License.LicenseProvider.SetLicenseKey(License);
        Spire.License.LicenseProvider.LoadLicense();
        #endregion
        // DataSet ds=null;
        // DataTable dtTable1 = null;
        try
        {
            if (!Directory.Exists(Server.MapPath("") + @"\ExcelTemplate\TempFolder\"))
            {
                Directory.CreateDirectory(Server.MapPath("") + @"\ExcelTemplate\TempFolder\");
            }

            //string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            //string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            //string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            String lsFileNamforFinalXls = FileName;// "ManagerPerformance_" + strMonth + "_" + strDay + "_" + strYear + ".xlsx";
            string ExcelFilePath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\\" + lsFileNamforFinalXls;

            if (System.IO.File.Exists(ExcelFilePath))
            {
                System.IO.File.Delete(ExcelFilePath);
            }

            //  string ExcelFilePath = ExcelfilePath + "TradingAppRecon" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";

            #region EPP code
            //FileInfo newFile = new FileInfo(ExcelFilePath);

            //using (OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(newFile))
            //{
            //    //if (ds.Tables["Table1"] != null)
            //    //{
            //    //    OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add("sheet1");
            //    //    ws.Cells["A1"].LoadFromDataTable(ds.Tables["Table1"], true);
            //    //    WorksheetFormatting(ws);
            //    //}
            //    //if (ds.Tables["Table2"] != null)
            //    //{
            //    //    OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add("sheet2");
            //    //    ws.Cells["A1"].LoadFromDataTable(ds.Tables["Table2"], true);
            //    //    WorksheetFormatting(ws);
            //    //}
            //    for (int i = 0; i < ds.Tables.Count; i++)
            //    {
            //        string SheetNme = ds.Tables[i].Rows[0][0].ToString();
            //        string GroupName = ds.Tables[i].Rows[0][1].ToString();
            //        i++;
            //        ds.Tables[i].Columns.Add("GroupName");
            //        for (int k = 0; k < ds.Tables[i].Rows.Count; k++)
            //        {
            //            ds.Tables[i].Rows[k]["GroupName"] = GroupName;
            //        }
            //        OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.Add(SheetNme);
            //        if (ds.Tables[i].Rows.Count > 0)
            //        {
            //            ws.Cells["A1"].LoadFromDataTable(ds.Tables[i], true);
            //            WorksheetFormatting(ws);
            //        }
            //        else
            //        {
            //            ws.Cells["J9:L10"].Merge = true;
            //            ws.Cells["J9:L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            //            ws.Cells["J9:L10"].Value = "No Data Found";
            //            ws.Cells["J9:L10"].Style.Font.Size = 16;
            //            //ws.Cells["J9:L10"].Style.Fill.BackgroundColor.SetColor(Color.Red);
            //        }

            //    }
            //    pck.Save();
            //}
            #endregion
            int SheetNo = 0;
            Workbook book = new Workbook();
            book.Version = ExcelVersion.Version2016;
            //book.CreateEmptySheets(ds.Tables.Count / 2);
            for (int i = 0; i < ds.Tables.Count; i++)
            {

                //  string SheetNme = ds.Tables[i].Rows[0][0].ToString();
                // string GroupName = ds.Tables[i].Rows[0][1].ToString();
                //  i++;
                //   ds.Tables[i].Columns.Add("GroupName");
                //   for (int k = 0; k < ds.Tables[i].Rows.Count; k++)
                //{
                //    ds.Tables[i].Rows[k]["GroupName"] = GroupName;
                //}


                Worksheet sheet = book.Worksheets[SheetNo];
                sheet.Name = "SheetNme";
                if (ds.Tables[i].Rows.Count > 0)
                {
                    sheet.Range[1, 1, 1, ds.Tables[i].Columns.Count].Style.Font.IsBold = true;

                    sheet.InsertDataTable(ds.Tables[i], false, 1, 1);
                    sheet.Range[1, 1, ds.Tables[i].Rows.Count, ds.Tables[i].Columns.Count].AutoFitColumns();
                    sheet.Range[1, 1, ds.Tables[i].Rows.Count, ds.Tables[i].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;
                    int j = 0;
                    int rownum = 0;
                    int count = 0;
                    int lastrow = ds.Tables[i].Rows.Count
                    foreach (CellRange cs in sheet.Range[1, 1, ds.Tables[i].Rows.Count, ds.Tables[i].Columns.Count])
                    {
                      
                        count++;
                        if(count = )
                   
                        string check = ds.Tables[i].Rows[j][1].ToString();
                        string text = cs.Text;
                        j++;
                        if (cs.Text.StartsWith("J") == true)

                        {
                            ExcelFont fontBold = book.CreateFont();
                            fontBold.IsBold = true;

                            RichText richText = cs.RichText;
                            //  richText.Text = "It is in Bold";
                            richText.SetFont(0, richText.Text.Length - 1, fontBold);
                        }


                    }
                    SheetNo++;
                }

                // book.SaveToFile(ExcelFilePath, ExcelVersion.Version2016);
                book.SaveToFile(ExcelFilePath);

                string vContain = "Excel Report Generated Succesfully ";
                //sw.WriteLine(vContain);
                
            }
            return ExcelFilePath;
        }
        catch (Exception e)
        {

            string vContain = "Excel Report Genration Fail,  Error " + e.ToString();
            //  sw.WriteLine(vContain);
            //LG.AddinLogFile(Form1.vLogFile, vContain);
            //   lblMsg.Text = vContain;

            return "";
        }
    }
    public string GenerateCsv(DataSet ds, bool isFund)
    {
        string lsFileNamforFinalXls = string.Empty;
        string ExcelFilePath = string.Empty;
        //if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\ReportOutput"))
        //{
        //    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\ReportOutput");
        //}  
        if (!Directory.Exists(Server.MapPath("") + @"\ExcelTemplate\TempFolder\"))
        {
            Directory.CreateDirectory(Server.MapPath("") + @"\ExcelTemplate\TempFolder\");
        }


        string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
        string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
        string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
        #region Commented 1_16_2109
        //if (SystemName.ToLower() != "gp-crm1")
        //{
        //    lsFileNamforFinalXls = "LegalEntityFriendlyCode_" + strMonth + "_" + strDay + "_" + strYear + "_TEST" + ".csv";
        //}
        //else
        //{

        //    lsFileNamforFinalXls = "LegalEntityFriendlyCode_" + strMonth + "_" + strDay + "_" + strYear + ".csv";
        //}
        #endregion
        if (isFund)
        {
            //if (server.ToLower() == "prod" || server.ToLower() == "crm1")
            //{
            //    lsFileNamforFinalXls = "FundShortNameDrop_" + strYear + strMonth + strDay + ".csv";
            //}
            //else
            // {
            lsFileNamforFinalXls = "FundShortNameDrop_" + strYear + strMonth + strDay + ".csv";

            //}
        }



        ExcelFilePath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls;

        if (System.IO.File.Exists(ExcelFilePath))
        {
            System.IO.File.Delete(ExcelFilePath);
        }
        try
        {

            if (ds.Tables.Count > 0)
            {
                System.Data.DataTable datatable = ds.Tables[0];

                // LG.AddinLogFile(Form1.vLogFile, " -------------------" + datatable.Rows.Count + DateTime.Now.ToString("yyyyMMddHHmmss"));
                // sw.WriteLine(" -------------------" + datatable.Rows.Count);
                if (datatable.Rows.Count > 0)
                {
                    //Build the CSV file data as a Comma separated string.
                    string csv = string.Empty;

                    foreach (DataColumn column in datatable.Columns)
                    {
                        //Add the Header row for CSV file.
                        csv += column.ColumnName + ',';
                    }

                    //Add new line.
                    csv += "\r\n";

                    foreach (DataRow row in datatable.Rows)
                    {
                        foreach (DataColumn column in datatable.Columns)
                        {
                            //Add the Data rows.
                            csv += AddEscapeSequenceInCsvField(row[column.ColumnName].ToString()) + ",";
                        }

                        //Add new line.
                        csv += "\r\n";
                    }

                    using (StreamWriter objWriter = new StreamWriter(ExcelFilePath))
                    {
                        objWriter.WriteLine(csv);
                    }
                    string vContain1 = "Csv Report Generated Succesfully ";


                    // sw.WriteLine(vContain1);
                    //  LG.AddinLogFile(Form1.vLogFile, vContain1 + DateTime.Now.ToString("yyyyMMddHHmmss"));

                    //}
                    // else
                    // {
                    if (csv == null)
                    {
                        string vContain12 = "No Data ";
                        // sw.WriteLine(vContain12);
                        //  LG.AddinLogFile(Form1.vLogFile, vContain12 + DateTime.Now.ToString("yyyyMMddHHmmss"));
                    }
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("No data Found");
                    using (StreamWriter objWriter = new StreamWriter(ExcelFilePath))
                    {
                        objWriter.WriteLine(sb);
                    }
                }
            }
            //else
            //{
            //    StringBuilder sb = new StringBuilder();
            //    sb.Append("No data Found");
            //    using (StreamWriter objWriter = new StreamWriter(ExcelFilePath))
            //    {
            //        objWriter.WriteLine(sb);
            //    }
            //}

        }
        catch (Exception ex)
        {
            string vContain = "CSV Report Genration Fail,  Error " + ex.ToString();
            // sw.WriteLine(vContain);
            // LG.AddinLogFile(Form1.vLogFile, vContain + DateTime.Now.ToString("yyyyMMddHHmmss"));

            return "";
        }
        return ExcelFilePath;
    }
    private string AddEscapeSequenceInCsvField(string ValueToEscape)
    {
        if (ValueToEscape.Contains(","))
        {
            return "\"" + ValueToEscape + "\"";
        }
        else
        {
            return ValueToEscape;
        }
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        //string[] SourceFileArray = null;
        //SourceFileArray = new string[2];
        //SourceFileArray[0] = (Server.MapPath("") + @"\ExcelTemplate\PDF test 1.pdf");
        //SourceFileArray[1] = (Server.MapPath("") + @"\ExcelTemplate\PDF test 2.pdf");


        //string DestinationPath = Request.MapPath("ExcelTemplate\\") + GeneralMethods.RemoveSpecialCharacters("Final.pdf");
        //PDF.MergeFiles1(DestinationPath, SourceFileArray);


        string oldFile = Server.MapPath("") + @"\ExcelTemplate\Final.pdf";
        string newFile = Server.MapPath("") + @"\ExcelTemplate\newFile1.pdf";

        // open the reader
        PdfReader reader = new PdfReader(oldFile);
        Rectangle size = reader.GetPageSizeWithRotation(1);
        Document document = new Document(size);

        // open the writer
        FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
        PdfWriter writer = PdfWriter.GetInstance(document, fs);
        document.Open();

        // the pdf content
        PdfContentByte cb = writer.DirectContent;

        // select the font properties
        BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
        cb.SetColorFill(Color.RED);
        cb.SetFontAndSize(bf, 8);

        // write the text in the pdf content
        cb.BeginText();
        string text = "Some random blablablabla...";
        // put the alignment and coordinates here
        cb.ShowTextAligned(1, text, 540, 620, 0);
        cb.EndText();
        cb.BeginText();
        text = "Other random blabla...";
        // put the alignment and coordinates here
        cb.ShowTextAligned(2, text, 100, 200, 0);
        cb.EndText();

        // create the new page and add it to the pdf
        PdfImportedPage page = writer.GetImportedPage(reader, 1);
        cb.AddTemplate(page, 100, 200);

        // close the streams and voilá the file should be changed :)
        document.Close();
        fs.Close();
        writer.Close();
        reader.Close();

    }

    protected void Button2_Click(object sender, EventArgs e)
    {
        try
        {
            PdfPCell CellR3Chart2 = new PdfPCell();
            PdfPTable LoR3Row2 = new PdfPTable(1);

            iTextSharp.text.Document pdoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER.Rotate(), -23, -20, 43, 0);//10,10
            string strGUID = Guid.NewGuid().ToString();
            string fsFinalLocation = Path.Combine(Server.MapPath("") + @"\ExcelTemplate" + "\\" + "ls_" + strGUID + ".pdf");

            PdfWriter writer = PdfWriter.GetInstance(pdoc, new FileStream(fsFinalLocation, FileMode.Create));

            pdoc.Open();

            iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(HttpContext.Current.Server.MapPath("") + @"\images\Gresham_Logo.png");
            png.SetAbsolutePosition(45, 557);//540
                                             //png.ScaleToFit(288f, 42f);
            png.ScalePercent(10);
            pdoc.Add(png);

            iTextSharp.text.Image chartimg1 = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\ExcelTemplate\PDF test 1.pdf");


            chartimg1.ScalePercent(25);
            CellR3Chart2.AddElement(chartimg1);
            LoR3Row2.AddCell(CellR3Chart2);
            pdoc.Add(LoR3Row2);
            pdoc.Close();
        }
        catch (Exception ex)
        {

        }
    }

    protected void Button3_Click(object sender, EventArgs e)
    {
        try
        {

            iTextSharp.text.Document document = null;

            PdfReader reader = new PdfReader(Server.MapPath("") + @"\ExcelTemplate\Final.pdf");
            PdfReader reader1 = new PdfReader(Server.MapPath("") + @"\ExcelTemplate\Final.pdf");
            //  string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\test_" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
            string filename = Server.MapPath("") + @"\ExcelTemplate" + "\\" + "test_" + Guid.NewGuid().ToString() + ".pdf";

            FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
            MemoryStream stream = new MemoryStream();
            //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 27, 31, 8);//10,10
            var writer = PdfWriter.GetInstance(document, fileStream);
            PdfStamper stamper = new PdfStamper(reader, stream);

            document.Open();

            //  for (var i = 1; i <= reader.NumberOfPages; i++)
            // {
            HttpContext.Current.Session["NumberofPages"] = reader.NumberOfPages;

            document.SetPageSize(PageSize.A4);
            document.NewPage();
            string fontpath = HttpContext.Current.Server.MapPath(".");
            var baseFont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            var importedPage = writer.GetImportedPage(reader, 1);
            var importedPage1 = writer.GetImportedPage(reader1, 1);
            var contentByte = writer.DirectContent;
            // contentByte.BeginText();
            //  contentByte.SetFontAndSize(baseFont, 2);

            String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();


            PdfContentByte cb1 = writer.DirectContent;
            ColumnText ct1 = new ColumnText(cb1);
            ct1.SetSimpleColumn(new Phrase(new Chunk(lsDateTime, setFontsAll(8, 0, 1, new iTextSharp.text.Color(128, 128, 128)))), 800, 15, 725, 40, 25, Element.ALIGN_RIGHT | Element.ALIGN_BOTTOM);
            ct1.Go();


            contentByte.AddTemplate(importedPage, 1f, 0, 0, 1f, 0, 0);
            contentByte.AddTemplate(importedPage1, 1f, 0, 0, 1f, 0, 0);
            document.Close();
            writer.Close();
            //  }
        }
        catch (Exception exx)
        {

        }
    }
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
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        }
        if (italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        }
        return font;
        #endregion
    }
}