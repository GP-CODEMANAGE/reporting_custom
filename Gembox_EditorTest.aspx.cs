using GemBox.Document;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Winthusiasm.HtmlEditor;

public partial class Gembox_EditorTest : System.Web.UI.Page
{
    GeneralMethods clsGM = new GeneralMethods();
    Logs lg = new Logs();
      public StreamWriter sw = null;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {


            DateTime dtmain = DateTime.Now;
            string LogFileName = string.Empty;
            LogFileName = "Log-" + DateTime.Now;
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
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        try
        {
           

            lg.AddinLogFile(Session["Filename"].ToString(), "CREATE PDF Start  " + DateTime.Now.ToString());


            string FileName = "EditorTest" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
            string FinalDestinationFileName = "EditorTestFinal" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";

            string DestinationPath = Request.MapPath("ExcelTemplate\\TempFolder\\" + FileName);
            string FinalDestinationPath = Request.MapPath("ExcelTemplate\\TempFolder\\" + FinalDestinationFileName);

            string Filepath = CreatePDF("", "", DestinationPath, FinalDestinationFileName, FinalDestinationPath);


            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");
            sb.Append("\nwindow.open('ViewReport.aspx?" + Filepath + "', 'mywindow');");
            sb.Append("</script>");
            ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
        }
        catch(Exception ex)
        {
            Label1.Visible = true;
            Label1.Text = "Error in Process " + ex.Message.ToString(); ;
        }
    }
    private string  CreatePDF(string strSourcePath, string strSourcePath1, string strDestPath,string FileName,string strFinalDest)
    {
        string filepath = "";
        try
        {
            lg.AddinLogFile(Session["Filename"].ToString(), "Inside CREATE PDF  " + DateTime.Now.ToString());
            string strSourcePath5 = string.Empty;
            ComponentInfo.SetLicense("D7OT-O3KE-PMVU-IXWZ");
            strSourcePath = "<span style='font - family: Verdana, sans - serif; font - size: 9pt;'>In total, GPESsssssss 13 managers called approximately $690,000 during the first quarter. This call will fund capital activity that occurred during the quarter and replenish cash balances in anticipation of upcoming capital calls. A description of GPES 13 activity is below:</span>< p class='MsoNormal'><span style = 'font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif' >< o:p>&nbsp;</o:p></span></p><p class='MsoNormal' style='margin-left:.25in'><b style = 'mso-bidi-font-weight:&#10;normal' >< i style='mso-bidi-font-style:normal'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > Atlas Capital Resources II(&ldquo; Atlas II&rdquo;)</span></i></b><b style = 'mso-bidi-font-weight:normal' >< span style='font-size:9.0pt;font-family:&#10;&quot;Verdana&quot;,sans-serif'> <o:p></o:p></span></b></p><p class='MsoNormal' style='margin-left:.5in'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > GPES 13 received capital calls totaling $550,000 from its $4 million commitment to Atlas II, a deep value private equity firm.The proceeds funded several new investments.Following these calls, GPES 13 will have funded 100% of its commitment.<o:p></o:p></span></p><p class='MsoNormal' style='margin-left:.25in'><b style = 'mso-bidi-font-weight:&#10;normal' >< i style='mso-bidi-font-style:normal'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > Grey Mountain Partners III(&ldquo; Grey Mountain III&rdquo;)</span></i></b><b style = 'mso-bidi-font-weight:normal' >< span style='font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif'> <o:p></o:p></span></b></p><p class='MsoNormal' style='margin-left:.5in'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > GPES 13 received capital calls totaling $50,000 from its $5 million commitment to Grey Mountain III, a U.S.lower-middle market buyout firm.<span style= 'mso-spacerun:yes' > &nbsp; </span>The proceeds funded new investments.<span style = 'mso-spacerun:yes' > &nbsp; </span>Following these calls, GPES 13 will have funded 84% of its commitment. <o:p></o:p></span></p><p class='MsoNormal' style='margin-left:.25in'><b style = 'mso-bidi-font-weight:&#10;normal' >< i style='mso-bidi-font-style:normal'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > Trilantic Capital Partners V(&ldquo; Trilantic V&rdquo;)</span></i></b><b style = 'mso-bidi-font-weight:normal' >< span style='font-size:9.0pt;font-family:&#10;&quot;Verdana&quot;,sans-serif'> <o:p></o:p></span></b></p><p class='MsoNormal' style='margin-left:.5in'><span style = 'font-size:9.0pt;&#10;font-family:&quot;Verdana&quot;,sans-serif' > GPES 13 received a capital call of $90,000 from its $5 million commitment to Trilantic V, an established U.S.middle market buyout firm.<span style= 'mso-spacerun:yes' > &nbsp; </span>The proceeds will fund new and follow-on investments.<span style= 'mso-spacerun:yes' > &nbsp; </span>Following this call, GPES 13 will have funded 100% of its commitment. <o:p></o:p></span></p>";

            DataTable dt = GetDataTable();
            strSourcePath = dt.Rows[0]["ssi_FundTxt"].ToString(); // from the database
            strSourcePath5 = txtFundDesc1.Value; // from the editor -- Working



            lg.AddinLogFile(Session["Filename"].ToString(), "ssi_FundTxt  " + DateTime.Now.ToString());
            lg.AddinLogFile(Session["Filename"].ToString(), strSourcePath + DateTime.Now.ToString());
            DocumentModel document = new DocumentModel();
            DocumentModel document1 = new DocumentModel();
            document.Content.LoadText(strSourcePath, LoadOptions.HtmlDefault);
          strDestPath  strDestPath.Replace(".pdf", ".docx");
                        document.Save(strDestPath);

            ConvertDocument(strDestPath, strFinalDest);


            //var htmlLoadOptions = new HtmlLoadOptions();
            //using (var htmlStream = new MemoryStream(htmlLoadOptions.Encoding.GetBytes(strSourcePath)))
            //{
            //    // Load input HTML text as stream.
            //    document = DocumentModel.Load(htmlStream, htmlLoadOptions);
            //    // Save output PDF file.
            //    document.Save(strDestPath);
            //}



            //// Insert HTML formatted text after the previous text.
            //var position = document.Content.End.LoadText(strSourcePath5,
            //    LoadOptions.HtmlDefault);

            //document.Save(strDestPath);






            filepath = FileName;

        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            filepath = ""; ;
        }
        return filepath;
    }
    private string ConvertDocument(string strSourcePath, string strDestPath)
    {
        try
        {

            ComponentInfo.SetLicense("D7OT-O3KE-PMVU-IXWZ");
            //ComponentInfo.FreeLimitReached += (sender1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            DocumentModel document = DocumentModel.Load(strSourcePath);

            document.Save(strDestPath.Replace(".xls", ".pdf"));

            return strDestPath.Replace(".pdf", ".xls");


        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            return "";
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
            greshamquery = "SELECT top (1) ssi_fundtxt,* FROM Ssi_templatefund ORDER BY createdon DESC";

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
        finally
        {
            Gresham_con.Close();
        }

        return ds_gresham.Tables[0];
    }
    protected void Inser_CRM_Click(object sender, EventArgs e)
    {
        lg.AddinLogFile(Session["Filename"].ToString(), "Inser_CRM_Click" + DateTime.Now.ToString());
        IOrganizationService service = null;
        bool bProceed = false;
            
        try
        {           
            
            try
            {
                service = clsGM.GetCrmService();
                bProceed = true;
               // strDescription = "Crm Service starts successfully";
            }
            catch (Exception ex)
            {
                Label1.Visible = true;
                Label1.Text =  "Error Connecting to CRM: " + ex.Message.ToString();
            }

            if (bProceed)
            {
                Entity objTemplateFund = new Entity("ssi_templatefund");

                FredCK.FCKeditorV2.FCKeditor txtFundDesc = ((FredCK.FCKeditorV2.FCKeditor)FindControl("txtFundDesc1"));            
                objTemplateFund["ssi_fundtxt"] = txtFundDesc.Value;
                lg.AddinLogFile(Session["Filename"].ToString(), "txtFundDesc.Value" + DateTime.Now.ToString());
                lg.AddinLogFile(Session["Filename"].ToString(), txtFundDesc.Value + DateTime.Now.ToString());
                //Response.Write("TEXT===" + txtFundDesc.Value);

                service.Create(objTemplateFund);
                Label1.Visible = true;
                Label1.Text = "Insert Successfull ";
            }
        }
        catch(Exception ex)
        {
            Label1.Visible = true;
            Label1.Text = Label1.Text +"Error Connecting to CRM: " + ex.Message.ToString();
        }

    }
}