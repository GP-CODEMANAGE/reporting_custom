using System;
using System.IO;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;

/// <summary>
/// Summary description for PDFMerge
/// </summary>
public class PDFMerge
{
    public void MergeFiles1(string destinationFile, string[] sourceFiles)
    {
        try
        {
            //New Logic for SLOA -- Existing Logic not working for SLOA 
            if (destinationFile.ToUpper().Contains("SLOA"))
            {
                MergeNew(destinationFile, sourceFiles);
                return;
            }
            int f = 0;
            // we create a reader for a certain document
            PdfReader reader = new PdfReader(sourceFiles[f]);
            // we retrieve the total number of pages
            int n = reader.NumberOfPages;
            //Console.WriteLine("There are " + n + " pages in the original file.");
            // step 1: creation of a document-object
            Document document = new Document(reader.GetPageSizeWithRotation(1));
            // step 2: we create a writer that listens to the document
            //FileInfo file = new FileInfo();
            //file.FullName = "e:\\repots\\1.txt";
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
                    if (rotation == 0 && f == 0)// && n == 0)
                    {
                        int pageinCoverLetter = reader.NumberOfPages;
                        HttpContext.Current.Session["pageinCoverLetter"] = pageinCoverLetter;
                        cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                    }
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
        }

        catch (Exception e)
        {
            string strOb = e.Message;
            HttpContext.Current.Session["ErrorinMerge"] = e.Message.ToString();
        }
    }
    public void MergeFiles_watermark(string destinationFile, string[] sourceFiles)
    {
        // Document document = null;
        try
        {
            //New Logic for SLOA -- Existing Logic not working for SLOA 
            if (destinationFile.ToUpper().Contains("SLOA"))
            {
                MergeNew(destinationFile, sourceFiles);
                return;
            }
            int f = 0;
            // we create a reader for a certain document
            PdfReader reader = new PdfReader(sourceFiles[f]);
            // we retrieve the total number of pages
            int n = reader.NumberOfPages;
            //Console.WriteLine("There are " + n + " pages in the original file.");
            // step 1: creation of a document-object
            Document document = new Document(reader.GetPageSizeWithRotation(1));
            
            //    document = new Document(reader.GetPageSizeWithRotation(1));

            // step 2: we create a writer that listens to the document
            //FileInfo file = new FileInfo();
            //file.FullName = "e:\\repots\\1.txt";
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
                    //WATERMARK 9_7_2017-sasmit
                    writer.PageEvent = new PdfWriterEvents("Testing"); 
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
        }
        catch (Exception e)
        {
            // document.Close();
            string strOb = e.Message;
        }
    }
    public void MergeFiles(string destinationFile, string[] sourceFiles)
    {
       // Document document = null;
        try
        {
            //New Logic for SLOA -- Existing Logic not working for SLOA 
            if (destinationFile.ToUpper().Contains("SLOA"))
            {
                MergeNew(destinationFile, sourceFiles);
                return;
            }
            int f = 0;
            // we create a reader for a certain document
            PdfReader reader = new PdfReader(sourceFiles[f]);
            // we retrieve the total number of pages
            int n = reader.NumberOfPages;
            //Console.WriteLine("There are " + n + " pages in the original file.");
            // step 1: creation of a document-object
              Document document = new Document(reader.GetPageSizeWithRotation(1));
         //    document = new Document(reader.GetPageSizeWithRotation(1));

            // step 2: we create a writer that listens to the document
            //FileInfo file = new FileInfo();
            //file.FullName = "e:\\repots\\1.txt";
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
        }
        catch (Exception e)
        {
           // document.Close();
            string strOb = e.Message;
        }
    }

    public int CountPageNo(string strFileName)
    {
        // we create a reader for a certain document
        PdfReader reader = new PdfReader(strFileName);
        // we retrieve the total number of pages
        return reader.NumberOfPages;
    }

    public void MergeNew(string destinationFile, string[] lstFiles)
    {       
        PdfReader reader = null;
        Document sourceDocument = null;
        PdfCopy pdfCopyProvider = null;
        PdfImportedPage importedPage;
        string outputPdfPath = destinationFile;


        sourceDocument = new Document();
        pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));

        //Open the output file
        sourceDocument.Open();

        try
        {
            //Loop through the files list
            for (int f = 0; f <= lstFiles.Length - 1; f++)
            {
                int pages = get_pageCcount(lstFiles[f]);

                reader = new PdfReader(lstFiles[f]);
                //Add pages of current file
                for (int i = 1; i <= pages; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }

                reader.Close();
            }
            //At the end save the output file
            sourceDocument.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }


    }
    public int get_pageCcount(string file)
    {
        using (StreamReader sr = new StreamReader(System.IO.File.OpenRead(file)))
        {
            Regex regex = new Regex(@"/Type\s*/Page[^s]");
            MatchCollection matches = regex.Matches(sr.ReadToEnd());

            return matches.Count;
        }
    }

}
public class PdfWriterEvents : IPdfPageEvent
{
    string watermarkText = string.Empty;

    public PdfWriterEvents(string watermark)
    {
        watermarkText = watermark;
    }
    public void OnStartPage(PdfWriter writer, Document document)
    {
        float fontSize = 80;
        float xPosition = iTextSharp.text.PageSize.A4.Width / 2;
        float yPosition = (iTextSharp.text.PageSize.A4.Height - 140f) / 2;
        float angle = 45;
        try
        {
            PdfContentByte under = writer.DirectContent;
            BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
            under.BeginText();
            under.SetColorFill(Color.LIGHT_GRAY);
            under.SetFontAndSize(baseFont, fontSize);
            under.ShowTextAligned(PdfContentByte.ALIGN_CENTER, watermarkText, xPosition, yPosition, angle);
            under.EndText();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
        }
    }
    public void OnEndPage(PdfWriter writer, Document document) { }
    public void OnParagraph(PdfWriter writer, Document document, float paragraphPosition) { }
    public void OnParagraphEnd(PdfWriter writer, Document document, float paragraphPosition) { }
    public void OnChapter(PdfWriter writer, Document document, float paragraphPosition, Paragraph title) { }
    public void OnChapterEnd(PdfWriter writer, Document document, float paragraphPosition) { }
    public void OnSection(PdfWriter writer, Document document, float paragraphPosition, int depth, Paragraph title) { }
    public void OnSectionEnd(PdfWriter writer, Document document, float paragraphPosition) { }
    public void OnGenericTag(PdfWriter writer, Document document, Rectangle rect, String text) { }
    public void OnOpenDocument(PdfWriter writer, Document document) { }
    public void OnCloseDocument(PdfWriter writer, Document document) { }
}
