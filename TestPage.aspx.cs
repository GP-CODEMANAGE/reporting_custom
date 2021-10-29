using GemBox.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using Karamasoft.WebControls.UltimateEditor;
public partial class TestPage : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if(!IsPostBack)
        {
            // Get the document from EditorContent.htm file
            string editorFile = Page.MapPath("~/EditorContent.htm");
            StreamReader sr = File.OpenText(editorFile);
            Ultimateeditor1.EditorHtml = sr.ReadToEnd();

            sr.Close();
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        string DestinationPath = Request.MapPath("ExcelTemplate\\TempFolder\\Test1Page.pdf");
        string DestinationPath1 = Request.MapPath("ExcelTemplate\\TempFolder\\Test2Page.pdf");
        HtmlToDOc("","", DestinationPath, DestinationPath1);
    }
    private void HtmlToDOc(string strSourcePath,string strSourcePath1, string strDestPath,string strDestPath1)
    {
        try
        {
            string strSourcePath5 = string.Empty;
            ComponentInfo.SetLicense("D7OT-O3KE-PMVU-IXWZ");
            strSourcePath = "<span style='font - family: Verdana, sans - serif; font - size: 9pt;'>In total, GPESsssssss 13 managers called approximately $690,000 during the first quarter. This call will fund capital activity that occurred during the quarter and replenish cash balances in anticipation of upcoming capital calls. A description of GPES 13 activity is below:</span>< p class='MsoNormal'><span style = 'font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif' >< o:p>&nbsp;</o:p></span></p><p class='MsoNormal' style='margin-left:.25in'><b style = 'mso-bidi-font-weight:&#10;normal' >< i style='mso-bidi-font-style:normal'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > Atlas Capital Resources II(&ldquo; Atlas II&rdquo;)</span></i></b><b style = 'mso-bidi-font-weight:normal' >< span style='font-size:9.0pt;font-family:&#10;&quot;Verdana&quot;,sans-serif'> <o:p></o:p></span></b></p><p class='MsoNormal' style='margin-left:.5in'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > GPES 13 received capital calls totaling $550,000 from its $4 million commitment to Atlas II, a deep value private equity firm.The proceeds funded several new investments.Following these calls, GPES 13 will have funded 100% of its commitment.<o:p></o:p></span></p><p class='MsoNormal' style='margin-left:.25in'><b style = 'mso-bidi-font-weight:&#10;normal' >< i style='mso-bidi-font-style:normal'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > Grey Mountain Partners III(&ldquo; Grey Mountain III&rdquo;)</span></i></b><b style = 'mso-bidi-font-weight:normal' >< span style='font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif'> <o:p></o:p></span></b></p><p class='MsoNormal' style='margin-left:.5in'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > GPES 13 received capital calls totaling $50,000 from its $5 million commitment to Grey Mountain III, a U.S.lower-middle market buyout firm.<span style= 'mso-spacerun:yes' > &nbsp; </span>The proceeds funded new investments.<span style = 'mso-spacerun:yes' > &nbsp; </span>Following these calls, GPES 13 will have funded 84% of its commitment. <o:p></o:p></span></p><p class='MsoNormal' style='margin-left:.25in'><b style = 'mso-bidi-font-weight:&#10;normal' >< i style='mso-bidi-font-style:normal'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > Trilantic Capital Partners V(&ldquo; Trilantic V&rdquo;)</span></i></b><b style = 'mso-bidi-font-weight:normal' >< span style='font-size:9.0pt;font-family:&#10;&quot;Verdana&quot;,sans-serif'> <o:p></o:p></span></b></p><p class='MsoNormal' style='margin-left:.5in'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > GPES 13 received a capital call of $90,000 from its $5 million commitment to Trilantic V, an established U.S.middle market buyout firm.<span style= 'mso-spacerun:yes' > &nbsp; </span>The proceeds will fund new and follow-on investments.<span style= 'mso-spacerun:yes' > &nbsp; </span>Following this call, GPES 13 will have funded 100% of its commitment. <o:p></o:p></span></p>";

            DataTable dt = GetDataTable();
            strSourcePath= dt.Rows[0]["ssi_FundTxt"].ToString(); // from the database
            strSourcePath5= txtFundDesc1.Value; // from the editor -- Working

            DocumentModel document = new DocumentModel();
            DocumentModel document1 = new DocumentModel();
            //document.Content.LoadText(strSourcePath);


            // Set the content for the whole document
            //document.Content.LoadText(strSourcePath, LoadOptions.HtmlDefault);

            //document.Content.LoadText(strSourcePath, LoadOptions.DocDefault); // not working
            //document.Content.LoadText(strSourcePath, LoadOptions.DocxDefault);// not working
            //document.Content.LoadText(strSourcePath, LoadOptions.PdfDefault);// not working
            //document.Content.LoadText(strSourcePath, LoadOptions.RtfDefault);
            //  document.Content.LoadText(strSourcePath, LoadOptions.TxtDefault);

            //document.Parent.Content.Delete();
            Section section = new Section(document);
            document.Sections.Add(section);

            Paragraph paragraph = new Paragraph(document);
            section.Blocks.Add(paragraph);

            Run run = new Run(document, "Hello World!");
            paragraph.Inlines.Add(run);

            var htmlLoadOptions = new HtmlLoadOptions();
            //using (var htmlStream = new MemoryStream(htmlLoadOptions.Encoding.GetBytes(strSourcePath)))
            //{
            //    // Load input HTML text as stream.
            //    document = DocumentModel.Load(htmlStream, htmlLoadOptions);
            //    // Save output PDF file.
            //    document.Save(strDestPath);
            //}

            document.Content.Start.LoadText("This is a plain text.", new CharacterFormat() { FontName = "Arial"});

            Paragraph pAsOfDate = new Paragraph();
            pAsOfDate.CharacterFormatForParagraphMark..Add(lochunkAsOfDate);
            pAsOfDate.Alignment = iTextSharp.text.Cell.ALIGN_LEFT;
            document.Add(pAsOfDate);

            // Insert HTML formatted text after the previous text.
            var position = document.Content.End.LoadText(strSourcePath5,
                LoadOptions.HtmlDefault);

            document.Save(strDestPath);


            //=====================================================================================
            strSourcePath1 = Ultimateeditor1.EditorHtml;


            var htmlLoadOptions1 = new HtmlLoadOptions();
            using (var htmlStream1 = new MemoryStream(htmlLoadOptions.Encoding.GetBytes(strSourcePath1)))
            {
                // Load input HTML text as stream.
                document1 = DocumentModel.Load(htmlStream1, htmlLoadOptions);
                // Save output PDF file.
                document1.Save(strDestPath1);
            }



            var bold = new CharacterFormat()
            {
                Bold = true
            };

            // Set the content for the 2nd paragraph
            // document.Sections[0].Blocks[1].Content.LoadText("Bold paragraph 2", bold);

            // Set the content for 3rd and 4th paragraph to be the same as the content of 1st and 2nd paragraph
            // var para3 = document.Sections[0].Blocks[2];
            // var para4 = document.Sections[0].Blocks[3];
            //  var destinationRange = new ContentRange(para3.Content.Start, para4.Content.End);
            //  var para1 = document.Sections[0].Blocks[0];
            //  var para2 = document.Sections[0].Blocks[1];
            // var sourceRange = new ContentRange(para1.Content.Start, para2.Content.End);
            //  destinationRange.Set(sourceRange);

            // Set content using HTML tags
            //  document.Sections[0].Blocks[4].Content.LoadText("Paragraph 5 <b>(part of this paragraph is bold)</b>", LoadOptions.HtmlDefault);

          //  document.Save(strDestPath, SaveOptions.PdfDefault);

        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            return;
        }
    }
    private DataTable GetDataTable()
    {
        string greshamquery;
        int totalCount = 0;

        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);


        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {

            // greshamquery = "select top (1) ssi_fundtxt,* from Ssi_templatefund";
            //  greshamquery = "select top (1) ssi_fundtxt,* from Ssi_templatefund where Ssi_TemplateIDName='2019-0930-Editor Test 1-GBhagia'";
            greshamquery = "select top (1) ssi_fundtxt,* from Ssi_templatefund where Ssi_TemplateIDName='2019-1231-EditorTest1_10-GBhagia'";

            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;

        }


        catch (Exception exc)
        {
            totalCount = 0;
            Response.Write(" sp fails error desc:" + exc.Message);
        }

        return ds_gresham.Tables[0];
    }
}