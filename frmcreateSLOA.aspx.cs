using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using ICSharpCode.SharpZipLib.Zip;
//using iTextSharp.text.pdf;
using System.Text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Data.SqlClient;
using Microsoft.IdentityModel.Claims;
using System.Threading;

public partial class frmcreateSLOA : System.Web.UI.Page
{

    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    clsReportTemplate objReportsTemplates = new clsReportTemplate();
    string zipfolderpath = string.Empty;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //  createZipFile();
            BindLegalEntity();
            BindHouseHold();
            BindFUND();

            // ZipFolder(@"C:\Users\byadav\Desktop\New folder (4)\Ziptest",)
        }

    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        // if()
        BindGridView();


    }

    protected void lstFund_SelectedIndexChanged(object sender, EventArgs e)
    {
        txtstartdate.Text = "";
        txtenddate.Text = "";
    }

    protected void txtenddate_TextChanged(object sender, EventArgs e)
    {

    }

    protected void txtstartdate_TextChanged(object sender, EventArgs e)
    {

    }

    protected void lstLegalEntity_SelectedIndexChanged(object sender, EventArgs e)
    {

    }

    protected void lstHouseHold_SelectedIndexChanged(object sender, EventArgs e)
    {

    }


    public void BindLegalEntity()
    {
        lstLegalEntity.Items.Clear();

        object HouseholdId = lstHouseHold.SelectedValue == "0" || lstHouseHold.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstHouseHold) + "'";

        //string strType = lstType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'"; 
        string sqlstr = "SP_S_LEGAL_ENTITY_LIST @HouseHoldID=" + HouseholdId + "";
        clsGM.getListForBindListBox(lstLegalEntity, sqlstr, "LegalEntityName", "LegalEntityNameId");

        lstLegalEntity.Items.Insert(0, "All");
        lstLegalEntity.Items[0].Value = "0";
        lstLegalEntity.SelectedIndex = 0;
    }

    public void BindHouseHold()
    {
        lstHouseHold.Items.Clear();

        //string strType = lstType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'"; 
        string sqlstr = "sp_s_Get_HouseHoldName";
        clsGM.getListForBindListBox(lstHouseHold, sqlstr, "name", "accountid");

        lstHouseHold.Items.Insert(0, "All");
        lstHouseHold.Items[0].Value = "0";
        lstHouseHold.SelectedIndex = 0;
    }

    public void BindFUND()
    {
        //lstLegalEntity.Items.Clear();

        ////string strType = lstType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'"; 
        //string sqlstr = "SP_S_LEGAL_ENTITY_LIST";
        //clsGM.getListForBindListBox(lstLegalEntity, sqlstr, "LegalEntityName", "LegalEntityNameId");

        lstFund.Items.Clear();

        //string strType = lstType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'"; 
        string sqlstr = "SP_S_FUND_LKUP @SLOAFundsOnlyFlg = 1";
        clsGM.getListForBindListBox(lstFund, sqlstr, "ssi_name", "ssi_FundId");


        //lstLegalEntity.Items.Insert(0, "All");
        //lstLegalEntity.Items[0].Value = "0";
        //lstLegalEntity.SelectedIndex = 0;

        #region for multi select in list box 

        //var names = new List<string>(new string[] { "ragu", "raju" });

        //foreach (var item in lstFund.Items)
        //{
        //    if (names.Contains(item.ToString()))
        //        lstFund.Items.FindByValue(item.ToString()).Selected = true;
        //}


        #endregion
    }

    public void createZipFile(string zipfolderpath)
    {
        try
        {
            // Depending on the directory this could be very large and would require more attention
            // in a commercial package.

            //Guid id = Guid.NewGuid();

            //String Todaysdate = "SLOA-" + DateTime.Now.ToString("dd-MMM-yyyy");


            //string zipfolderpath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\ZipFileTest\" + Todaysdate + id.ToString(); //+ Todaysdate + id.ToString();


            //  string zipfolderpath = @"C:\Users\byadav\Desktop\Newfolder(4)\Ziptest\"; //+ Todaysdate + id.ToString();




            // zipfolderpath = zipfolderpath + Todaysdate;
            //if (!Directory.Exists(zipfolderpath.Trim()))
            //{
            //    Directory.CreateDirectory(zipfolderpath.Trim());
            //}



            //  string zipfilepath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\";

            string[] filenames = Directory.GetFiles(zipfolderpath);

            // 'using' statements guarantee the stream is closed properly which is a big source
            // of problems otherwise.  Its exception safe as well which is great.

            if (zipfolderpath.Length > 0)
            {
                using (ZipOutputStream s = new ZipOutputStream(File.Create(zipfolderpath + "\\" + "SLOA.zip")))
                {
                    //ZipOutputStream s = new ZipOutputStream(File.Create(zipfolderpath));


                    s.SetLevel(9); // 0 - store only to 9 - means best compression

                    byte[] buffer = new byte[4096];

                    foreach (string file in filenames)
                    {

                        // Using GetFileName makes the result compatible with XP
                        // as the resulting path is not absolute.
                        ZipEntry entry = new ZipEntry(Path.GetFileName(file));

                        // Setup the entry data as required.

                        // Crc and size are handled by the library for seakable streams
                        // so no need to do them here.

                        entry.Size = System.IO.File.Open(file, System.IO.FileMode.OpenOrCreate,
                        System.IO.FileAccess.Read, System.IO.FileShare.Read).Length;


                        // Could also use the last write time or similar for the file.
                        //  entry.DateTime = DateTime.Now;
                        // s.UseZip64 = UseZip64.Off;
                        s.PutNextEntry(entry);

                        using (FileStream fs = File.OpenRead(file))
                        {

                            // Using a fixed size buffer here makes no noticeable difference for output
                            // but keeps a lid on memory usage.
                            int sourceBytes;
                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                s.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }

                    // Finish/Close arent needed strictly as the using statement does this automatically

                    // Finish is important to ensure trailing information for a Zip file is appended.  Without this
                    // the created file would be invalid.
                    s.Finish();

                    // Close is important to wrap things up and unlock the file.
                    s.Close();
                }
            }
        }
        //}
        catch (Exception ex)
        {
            Console.WriteLine("Exception during processing {0}", ex);

            // No need to rethrow the exception as for our purposes its handled.
        }
    }


    public void BindGridView()
    {
        object StartDate = txtstartdate.Text == "" ? "null" : "'" + txtstartdate.Text + "'";
        object EndDate = txtenddate.Text == "" ? "null" : "'" + txtenddate.Text + "'";


        object HouseholdId = lstHouseHold.SelectedValue == "0" || lstHouseHold.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstHouseHold) + "'";
        object FundId = lstFund.SelectedValue == "0" || lstFund.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
        object LegalEntityId = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";



        //object EndDate = txtenddate.Text == "" ? "null" : "'" + txtenddate.Text + "'";

        string sql = "SP_S_SLOA_GRIDLIST @IncEmailRecipients = 1, @StartDate = " + StartDate + ", @EndDate = " + EndDate + ", @HHID = " + HouseholdId + ", @LegalEntityNameID = " + LegalEntityId + ", @FundId = " + FundId;
        DataSet loDataset = clsDB.getDataSet(sql);


        if (loDataset.Tables[0].Rows.Count < 1)
        {
            lblError.Text = "Record not found";
            //  lblError.Visible = true;
            GridView1.DataSource = null;
            GridView1.DataBind();
            return;
        }


        GridView1.Columns[5].Visible = true;
        GridView1.Columns[6].Visible = true;

        GridView1.Columns[7].Visible = true;
        GridView1.Columns[8].Visible = true;


        GridView1.DataSource = loDataset;
        GridView1.DataBind();

        Button2.Visible = true;
        lblError.Text = "";
        GridView1.Columns[5].Visible = false;
        GridView1.Columns[6].Visible = false;
        GridView1.Columns[7].Visible = false;
        GridView1.Columns[8].Visible = false;

        ViewState["GridData"] = loDataset.Tables[0];

        string Fundid = string.Empty;
        string FundName = string.Empty;


        //string Fund_Name = null;
        DataTable dtFund = new DataTable();
        dtFund.Columns.Add("FundId");
        dtFund.Columns.Add("FundName");

        foreach (System.Web.UI.WebControls.ListItem li in lstFund.Items)
        {
            if (li.Selected)
            {
                //cnt++;
                Fundid = li.Value.ToString();
                string fundname = li.Text.ToString();
                DataRow dr = dtFund.NewRow();
                dr["FundId"] = Fundid;
                dr["FundName"] = fundname;
                dtFund.Rows.Add(dr);

            }

        }




        int cnt = 0;


        string funName = string.Empty;

        foreach (DataRow row1 in dtFund.Rows)
        {
            string Funid = row1["FundId"].ToString();
            funName = row1["FundName"].ToString();
            bool bproceed = false;
            foreach (DataRow row in loDataset.Tables[0].Rows)
            {
                //string Fund_Id = row.Cells[7].Text;
                string Fund_Id = row["FundID"].ToString();

                //  string Fund_Name = row.Cells[3].Text;

                string Fund_Name = row["Fund Name"].ToString();

                if (Fund_Id == Funid)
                {

                    bproceed = true;
                    //break;
                }
            }

            if (!bproceed)

                FundName = FundName + "|" + funName;


            //string Fund_Name = row.Cells[3].Text;


        }


        if (FundName != "")
        {
            string[] strTo = null;
            if (FundName != "")
                strTo = FundName.Split('|');


            string msg = string.Empty;
            for (int i = 0; i < strTo.Length; i++)
            {
                if (i == 0)
                    msg = strTo[i].ToString();
                else
                    msg = msg + "<br>" + strTo[i].ToString();
            }


            lblError.Text = "No Recommendation/Position found for:" + msg;
        }


        // GridView1.Columns[15].Visible = false;



    }


    protected void Button2_Click(object sender, EventArgs e)
    {
        string SourceFileName = string.Empty;
        //  string zipfolderpath = string.Empty;
        string filename = string.Empty;
        int cnt = 0;
        string[] SourceFileName1 = null;

        string strUserName = string.Empty;

        lblError.Text = "";
        DataTable dtSLOA = new DataTable();
        DataSet ds = new DataSet();

        // DataTable dtTableOfContent = new DataTable();
        int filecount = 0;
        dtSLOA.Columns.Add("LegalEntityID");
        dtSLOA.Columns.Add("ContactId");
        dtSLOA.Columns.Add("FundID");
        // dtbatch.Columns.Add("PageIndex");

        Guid id = Guid.NewGuid();
        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            strUserName = "corp\\gbhagia";
        }
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            strUserName = claimsIdentity.Name;

        }



        String Todaysdate = "SLOA-" + DateTime.Now.ToString("dd-MMM-yyyy") + id.ToString();
        zipfolderpath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + Todaysdate; //+ id.ToString();
        //   zipfolderpath = zipfolderpath + Todaysdate;
        if (!Directory.Exists(zipfolderpath.Trim()))
        {
            Directory.CreateDirectory(zipfolderpath.Trim());
        }
        try
        {


            foreach (GridViewRow row in GridView1.Rows)  // To allow or disallow action logic
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                object HHID = row.Cells[5].Text.Trim().Replace("HHID", "").Replace("&nbsp;", "");
                object LEID = row.Cells[6].Text.Trim().Replace("LEID", "").Replace("&nbsp;", "");
                // string FUNDID = row.Cells[6].Text.Trim().Replace("FUNDID", "").Replace("&nbsp;", "");

                object FUND_ID = row.Cells[7].Text.Trim().Replace("FUNDID", "").Replace("&nbsp;", "");


                object CONTACT_ID = row.Cells[8].Text.Trim().Replace("CONTACTID", "").Replace("&nbsp;", "");


                // object FUND_ID = FUNDID == "" ? "null" : "'" + FUNDID + "'";

                if (chkSelectNC.Checked)
                {

                    //  string sql = "SP_S_SLOA_REPORT  @LegalEntityID = '" + LEID + "', @FundID = '" + FUND_ID + "'";

                    //string sql = "SP_S_SLOA_REPORT  @LegalEntityID = '" + LEID + "', @FundID = '" + FUND_ID + "',@ContactID='" + CONTACT_ID + "'";
                    //DataSet loDataset = clsDB.getDataSet(sql);
                    cnt++;
                    // SourceFileName = objReportsTemplates.Get_SLOA(loDataset);
                    Guid id1 = Guid.NewGuid();
                    //  filename = "test" + id1 + ".pdf";
                    //  File.Copy(SourceFileName, zipfolderpath + "\\" + filename);

                    DataRow dr = dtSLOA.NewRow();
                    dr["LegalEntityID"] = LEID;
                    dr["ContactId"] = CONTACT_ID;
                    dr["FundID"] = FUND_ID;
                    dtSLOA.Rows.Add(dr);

                }

            }



            try
            {
                ds = InsertData("SP_S_SLOA_REPORT", dtSLOA, "@T_SLOA_REPORT");

            }

            catch (Exception ex)
            {
                lblError.Text = ex.StackTrace.ToString();
            }


            int count = ds.Tables.Count;

            SourceFileName1 = new string[count];

            for (int i = 0; i <= ds.Tables.Count - 1; i++)
            {
                Random rnd = new Random();
                int rndNum = rnd.Next(1000, 9999);
                string RandomNumStr = "_" + rndNum.ToString();


                filename = Convert.ToString(ds.Tables[i].Rows[0]["ssi_FileName"]).Replace("'","") + RandomNumStr + ".pdf";

                SourceFileName = objReportsTemplates.Get_SLOA(ds.Tables[i], Todaysdate, filename);

                //  filename = Convert.ToString(ds.Tables[i].Rows[i]["ssi_FileName"]) + ".pdf";

                //  SourceFileName1[i] = SourceFileName;


                // File.Copy(SourceFileName, zipfolderpath + "\\" + filename);
            }


            string zipfile = Server.MapPath("") + @"\ExcelTemplate\TempFolder\";

            if (ds.Tables.Count > 1)
            {

                createZipFile(zipfolderpath);
                // createZipFile(SourceFileName1);
                string filepath = zipfile + "\\" + "SLOA.zip";

                File.Copy(zipfolderpath + "\\" + "SLOA.zip", filepath, true);

                Download_File(filepath, "SLOA.zip");
                //lblMessage.Text = "Success";

                //lblMessage.Visible = true;
                lblError.Text = "SLOA generated sucessfully";
                lblError.Visible = true;

                BindGridView();


            }
            else if (ds.Tables.Count == 1)
            {
                string filepath = zipfile + "\\" + filename;

                File.Copy(SourceFileName, filepath, true);
                Download_File(filepath, filename);
                lblError.Text = "SLOA generated sucessfully";
                lblError.Visible = true;
                BindGridView();


            }

            else
            {
                if (cnt == 0)
                    lblError.Text = "Please Select a Record for generate SLOA";
            }

        }
        catch (Exception ex)
        {
            //Directory.Delete(zipfolderpath, true);

            lblError.Text = "Error Occured : " + ex.StackTrace.ToString();

            //  Response.Write(ex.ToString());
        }

        finally
        {
            //Thread.Sleep(10000);
            //Directory.Delete(zipfolderpath, true);
        }
    }


    private void Download_File(string FilePath, string FileName)
    {
        try
        {

            //Response.ContentType = ContentType;
            //Response.AppendHeader("Content-Disposition", "attachment; filename=" + FileName);
            //Response.WriteFile(FilePath);
            //Response.Flush();
            //Response.End();

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");

            ViewState["FileName"] = FileName;
            // sb.Append("\nwindow.open('ViewReport.aspx?" + FileName + "', 'mywindow');");

            sb.Append("\nwindow.open('viewreportsloa.aspx?" + FileName + "', 'mywindow');");
            sb.Append("</script>");
            ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());



            /* add beacuse Response.End() throws ThreadAborted Exception */
            // HttpContext.Current.ApplicationInstance.CompleteRequest();
            //  Directory.Delete(zipfolderpath, true);

            //Response.Clear();
            //Response.Charset = "";
            //Response.ContentType = "application/zip";
            //Response.AppendHeader("Content-Disposition", "attachment; filename=" + FileName);
            //Response.TransmitFile(FilePath);
            //Response.Flush();
            //Response.End();



        }
        catch (Exception ex)
        {


        }
    }

    public DataSet InsertData(string vSqlQuery, DataTable dt, string parameter1)
    {
        try
        {
            DataSet ds = new DataSet("TimeRanges");
            string lsConnectionstring = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);
            using (SqlConnection conn = new SqlConnection(lsConnectionstring))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = vSqlQuery;
                cmd.Parameters.AddWithValue(parameter1, dt);
                // cmd.Parameters.AddWithValue(parameter2, deleteFlg);

                //cmd.Parameters.Add("@ERROR", SqlDbType.Char, 500);
                //cmd.Parameters["@ERROR"].Direction = ParameterDirection.Output;  

                cmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = cmd;
                da.SelectCommand.CommandTimeout = 1800;
                da.Fill(ds);

                //int code = (int)cmd.Parameters["@ERROR"].Value;

                //if (code == 1)
                //{
                //    lblError.Text = "Contact Adminstartor.";

                //}

            }

            return ds;
        }
        catch (Exception ex)
        {
            lblError.Text = lblError.Text + "\n" + vSqlQuery + " Failed to execute. " + ex.Message;
            lblError.Visible = true;
            return null;
        }
    }



    private void Download_File1(string FilePath, string FileName)
    {
        // string _path = Request.PhysicalApplicationPath + "CV/" + name;
        //System.IO.FileInfo _file = new System.IO.FileInfo(FilePath);

        //Response.Clear();
        //Response.AddHeader("Content-Disposition", "attachment; filename=" + FileName);
        //Response.AddHeader("Content-Length", FileName.Length.ToString());
        //Response.ContentType = "application/octet-stream";
        //Response.WriteFile(FilePath);
        //Response.End();

        //  Response.ContentType = "application/zip";
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment; filename=" + FileName);
        Response.TransmitFile(FilePath);

        // Response.Close();

    }


    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {

    }

    protected void lstFund_SelectedIndexChanged1(object sender, EventArgs e)
    {
        txtstartdate.Text = "";
        txtenddate.Text = "";
        lblError.Text = "";
    }

    protected void lstHouseHold_SelectedIndexChanged1(object sender, EventArgs e)
    {
        BindLegalEntity();
        lblError.Text = "";
    }
}