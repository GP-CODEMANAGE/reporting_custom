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

public partial class Samples_Feature_Configuration_Configuration : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            // Get the document from EditorContent.htm file
            string editorFile = Page.MapPath("~/EditorContent.htm");
            StreamReader sr = File.OpenText(editorFile);
            UltimateEditor1.EditorHtml = sr.ReadToEnd();
            sr.Close();
        }
    }
    protected void rblConfiguration_SelectedIndexChanged(object sender, EventArgs e)
    {
        switch (rblConfiguration.SelectedValue)
        {
            case "Full":
                UltimateEditor1.EditorSource = "~/UltimateEditorInclude/UltimateEditorFull.xml";
                UltimateEditor1.DisplayCharCount = true;
                UltimateEditor1.MaxCharCount = 50000;
                UltimateEditor1.DisplayWordCount = true;
                UltimateEditor1.MaxWordCount = 10000;
                break;
            case "Default":
                UltimateEditor1.EditorSource = "~/UltimateEditorInclude/UltimateEditor.xml";
                UltimateEditor1.DisplayCharCount = false;
                UltimateEditor1.MaxCharCount = -1;
                UltimateEditor1.DisplayWordCount = false;
                UltimateEditor1.MaxWordCount = -1;
                break;
            case "Basic":
                UltimateEditor1.EditorSource = "~/UltimateEditorInclude/UltimateEditorBasic.xml";
                UltimateEditor1.DisplayCharCount = false;
                UltimateEditor1.MaxCharCount = -1;
                UltimateEditor1.DisplayWordCount = false;
                UltimateEditor1.MaxWordCount = -1;
                break;
        }
    }
}
