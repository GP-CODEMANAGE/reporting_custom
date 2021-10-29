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
using System.IO;

public partial class ViewReport : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string mimeType = "Application/pdf";
        string fileName = string.Empty;
        string filePath = string.Empty;
        string str = "";
        if (Session["id"] != null)
        {
            //Response.Write(Session["id"].ToString());
            str = Session["id"].ToString();
            Session.RemoveAll();

        }
        else
        {
            str = Convert.ToString(Request.QueryString[0]);
        }


        fileName = str;
        filePath = "./ExcelTemplate/TempFolder/" + fileName;
        filePath = Server.MapPath(filePath);

        bool isFileExist = File.Exists(filePath);
        //Response.Write(isFileExist.ToString());
        if (isFileExist)
        {
            fileName = "\"" + fileName + "\"";
            Response.Clear();
            Response.Buffer = false; //transmitfile self buffers
            //Response.Clear();
            Response.ClearContent();
            Response.ClearHeaders();
            Response.ContentType = "application/pdf";
            Response.AddHeader("Content-Disposition", "inline;filename=" + fileName);
            Response.WriteFile(filePath); //transmitfile keeps entire file from loading into memory
            Response.End();

            /*
            Stream stream = (); // Assuming you have a method that does this.
            BinaryReader reader = new BinaryReader(stream);

            HttpResponse response = HttpContext.Current.Response;
            response.ContentType = "application/pdf";
            response.AddHeader("Content-Disposition", "attachment; filename=file.pdf");
            response.ClearContent();
            response.OutputStream.Write(reader.ReadBytes(1000000), 0, 1000000);

            // End the response to prevent further work by the page processor.
            response.End();


            FileInfo file = new FileInfo(filePath);
            Response.Clear();

            if (mimeType.Equals("Application/pdf"))
            {
                Response.ContentType = mimeType;
                //Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName);//fileName
                Response.AddHeader("Content-Disposition", "inline;filename=" + fileName);//fileName
            }
            else
            {
                Response.ContentType = mimeType;
                Response.AddHeader("Content-Disposition", "inline;filename=" + fileName);//fileName
            }

            #region commented code

            //FileStream fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);
            ////Code for streaming the object while writing
            //long numBytes = file.Length;
            //BinaryReader br = new BinaryReader(fs);

            //const int ChunkSize = 1024;
            //byte[] binary = br.ReadBytes((int)numBytes);

            //MemoryStream ms = new MemoryStream(binary);
            //int SizeToWrite = ChunkSize;

            //for (int i = 0; i < binary.GetUpperBound(0) - 1; i = i + ChunkSize)
            //{
            //    if (!Response.IsClientConnected) return;
            //    if (i + ChunkSize >= binary.Length) SizeToWrite = binary.Length - i;
            //    byte[] chunk = new byte[SizeToWrite];
            //    ms.Read(chunk, 0, SizeToWrite);
            //    Response.BinaryWrite(chunk);
            //}
            //br.Close();
            //fs.Close();

            #endregion

            //Output file to response stream
            byte[] buffer = new byte[1024];
            int bytesRead = -1;

            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            Stream output = this.Context.Response.OutputStream;
            
            //
            Stream stream = output; // Assuming you have a method that does this.
            BinaryReader reader = new BinaryReader(stream);

            HttpResponse response = HttpContext.Current.Response;
            response.ContentType = "application/pdf";
            response.AddHeader("Content-Disposition", "attachment; filename=file.pdf");
            response.ClearContent();
            response.OutputStream.Write(reader.ReadBytes(1000000), 0, 1000000);

            // End the response to prevent further work by the page processor.
            response.End();
            //

            //while ((bytesRead = fs.Read(buffer, 0, buffer.Length)) > 0)
            //{
            //    output.Write(buffer, 0, bytesRead);
            //}

            //output.Close();
            //fs.Close();

            //Response.Flush();
            //Response.Close();
             * */
        }
    }
}
