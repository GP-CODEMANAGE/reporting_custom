using System;
using iTextSharp.text;

//using iTextSharp.text.html;
using System.IO;
//using iTextSharp.text.html.simpleparser;
//using iTextSharp.tool.xml;
using System.Text;
using System.Web;
using System.Collections;
using System.Text.RegularExpressions;
//using GemBox.Document;
//using iTextSharp.text.html.simpleparser;


public partial class CuteSoftText : System.Web.UI.Page
{
    clsReportTemplate cls = new clsReportTemplate();
    public string DOCUMENT_HTML_START = "<html><head></head><body>";
    public string DOCUMENT_HTML_END = "</body></html>";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            string sourcepath = @"C: \Users\byadav\Desktop\Newfolder(4)\Gresham Docs\Test\Sep 2019 Distribution Letter.docx";

            string destpath = @"C: \Users\byadav\Desktop\Newfolder(4)\Gresham Docs\Test\covnertedpdf.pdf";

            // ConvertDocument(sourcepath, destpath);
          

        }
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
         string str = txtLetterText.Value;

        string path = createPdftemp(str);


        //   CrreatePdf();


    }



    public void CrreatePdf()
    {
        //string str = txtLetterText.Value;

        //StringBuilder sb = new StringBuilder(str);


        //using (MemoryStream ms = new MemoryStream())
        //{

        //    Document document = new Document(PageSize.A4, 25, 25, 30, 30);


        //    String ls = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\testfile.pdf";

        //    iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));

        //    document.Open();

        //    // HTMLWorker hw = new HTMLWorker(document);


        //    iTextSharp.text.html.simpleparser.HTMLWorker hw = new HTMLWorker(document);




        //    hw.Parse(new StringReader(str.ToString()));

        //    document.Add(new Paragraph(str));

        //    document.Close();

        //    writer.Close();
        //    Response.ContentType = "pdf /application";

        //    Response.AddHeader("content-disposition",
        //    "attachment;filename=First PDF document.pdf");

        //    Response.OutputStream.Write(ms.GetBuffer(), 0, ms.GetBuffer().Length);

        //}

        //using (StringWriter sw = new StringWriter(sb))
        //{
        //    using (HtmlTextWriter hw = new HtmlTextWriter(sw))
        //    {
        //        // GridView1.RenderControl(hw);
        //        StringReader sr = new StringReader(sw.ToString());
        //        Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
        //        PdfWriter writer = PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
        //        pdfDoc.Open();
        //        XMLWorkerHelper.GetInstance().ParseXHtml(writer, pdfDoc, sr);


        //        pdfDoc.Close();
        //        Response.ContentType = "application/pdf";
        //        Response.AddHeader("content-disposition", "attachment;filename=GridViewExport.pdf");
        //        Response.Cache.SetCacheability(HttpCacheability.NoCache);
        //        Response.Write(pdfDoc);
        //        Response.End();
        //    }
        //}


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

    private string ConvertDocument(string strSourcePath, string strDestPath)
    {
        //try
        //{

        //    ComponentInfo.SetLicense("D7OT-O3KE-PMVU-IXWZ");
        //    //ComponentInfo.FreeLimitReached += (sender1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
        //    DocumentModel document = DocumentModel.Load(strSourcePath);

        //    document.Save(strDestPath.Replace(".docx", ".pdf"));

        //    return strDestPath.Replace(".pdf", ".docx");


        //}
        //catch (Exception ex)
        //{
        //    Response.Write(ex.ToString());
        //    return "";

        //}
        return "";
    }
}