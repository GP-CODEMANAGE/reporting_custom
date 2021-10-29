using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
//using System.IO.Compression;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Core;
using System.Data.SqlClient;
using RKLib.ExportData;
using System.Security;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

public partial class FidelityFolderRename : System.Web.UI.Page
{
    bool bproceed = false;
    string FilePath = string.Empty;
    sharepoint sp = new sharepoint();

    string DestFilepath = @"D:\Brijesh data\Development New\Gresham Partners\Fiedility folder Rename\test\";
    string ErrorMsg = string.Empty;
    string strFolderExists = string.Empty;


    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //compressDirectory(Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + "\\" + "FidilityZipFile" + "\\" + "G09014033_20190128", Server.MapPath("") + @"\ExcelTemplate\TempFolder\G09014033_20190128.zip");
            // btnSubmit.Attributes.Add("onclick", "return do_totals1();");

            LoadList();

            // RenameFile();

            //DataTable dt = GetDataTableFromCsv(@"D:\Brijesh data\Development New\Gresham Partners\Fiedility folder Rename\test\_index.csv", false);
        }
    }



    public bool UploadFile()
    {
        try
        {

            string filename = Path.GetFileName(UPLFF.FileName);
            string strextension = Path.GetExtension(UPLFF.FileName);


            if (strextension == ".zip")
            {
                //ViewState["FilePath"] = filename;
                FilePath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\FileUpload\" + filename;
                ViewState["filename"] = filename;
                if (System.IO.File.Exists(FilePath))
                    System.IO.File.Delete(FilePath);
                ViewState["FilePath"] = FilePath;
                UPLFF.SaveAs(FilePath);

                bproceed = true;
            }
            else
            {
                bproceed = false;
                string Desc = "Please Enter Correct File";
                lblError.Text = Desc;
                lblError.Visible = true;


                lblTotalFileCount.Text = "";
                lblSucessFileCount.Text = "";
                lblfailFileCount.Text = "";

            }
        }
        catch (Exception ex)
        {
            string Desc = "ERROR: Incorrect File format ";// + ex.Message;
            lblError.Text = Desc;
            lblError.Visible = true;
            bproceed = false;

            lblTotalFileCount.Text = "";
            lblSucessFileCount.Text = "";
            lblfailFileCount.Text = "";
        }

        return bproceed;
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        int failfilecount = 0;
        lblError.Text = "";

        //string ExtarctOutputFolder = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + "FidilityZipFile" + "\\";


        string FileUploadPath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\FileUpload\";
        string ExtarctOutputFolder = Server.MapPath("") + @"\ExcelTemplate\TempFolder\FileUpload\" + "FidilityZipFile" + "\\";

        string DownloadZipFile = Server.MapPath("") + @"\ExcelTemplate\TempFolder\FileUpload\" + "\\" + "DownloadZipFile" + "\\";
        // string CsvFilePath = 

        #region Clean Zip Extarct Folder
        DirectoryInfo directoryInfo = new System.IO.DirectoryInfo(ExtarctOutputFolder);
        int directoryCount = directoryInfo.GetDirectories().Length;
        DirectoryInfo[] Dir = directoryInfo.GetDirectories();


        foreach (DirectoryInfo dir in directoryInfo.GetDirectories())
        {
            dir.Delete(true);
        }

        string[] arrayFile = Directory.GetFiles(FileUploadPath, "*.zip");

        foreach (string str in arrayFile)
        {
            System.IO.File.Delete(str);
        }



        if (directoryCount > 0)
        {
            string dirName = Dir[0].ToString();
            ExtarctOutputFolder = ExtarctOutputFolder + "\\" + dirName;
            //  Directory.Delete(ExtarctOutputFolder, true);


            string[] array1 = Directory.GetFiles(DownloadZipFile, "*.zip");

            if (array1 != null && array1.Length != 0)
            {
                string DownloadPath = array1[0].ToString();
                System.IO.File.Delete(DownloadPath);
            }



            //if(Directory.GetFiles(DownloadZipFile,"*.zip"))
            //File.Delete(DownloadZipFile);
        }
        //var directoryname = directoryInfo.GetDirectories();

        //   var subfoldername = directoryname
        //if (Directory.Exists(ExtarctOutputFolder))
        //    Directory.Delete(ExtarctOutputFolder, true);
        #endregion


        #region Load File


        try
        {

            bool result = UploadFile();


            if (result)
            {


                //if (result)
                //{
                //    using (ZipArchive archive = ZipFile.OpenRead(FilePath))
                //    {
                //        foreach (ZipArchiveEntry entry in archive.Entries)
                //        {
                //            //entry.ExtractToFile(Path.Combine(DestFilepath, entry.FullName));

                //            // entry.ExtractToFile(Path.Combine(DestFilepath, entry.FullName));

                //            entry.ExtractToFile(DestFilepath + "\\" + entry.Name);
                //        }
                //    }
                //}

                ExtractZipContent(FilePath, "", ExtarctOutputFolder);

                //string IndexFilepath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + "\\" + "FidilityZipFile" + "\\" + ViewState["Foldername"] + "\\" + "_index.csv";

                string IndexFilepath = ViewState["Foldername"] + "\\" + "_index.csv";
                //  DataTable dt = new DataTable();
                // Specify the column list to export
                // int[] iColumns = { 1, 2, 3, 5, 6 };

                //int[] iColumns = { 1, 2, 3, 4, 5 };

                //// Export the details of specified columns to Excel
                //RKLib.ExportData.Export objExport = new
                //    RKLib.ExportData.Export("Web");
                //objExport.ExportDetails(dt,
                //     iColumns, Export.ExportFormat.CSV, IndexFilepath);


                DataTable dt = readcsv(IndexFilepath);

                int rowcount = dt.Rows.Count;

                //  DataTable dt = GetDataTableFromCsv(IndexFilepath, false);
                string greshamquery = "SP_S_FIDELITY_FILENAME";
                DataSet ds_gresham = InsertData(greshamquery, dt, "@FidelityAccount");

                string folderpath = Convert.ToString(ViewState["Foldername"]);

                string FinalFileName = Convert.ToString(ViewState["filename"]);


                RenameFile(folderpath, ds_gresham.Tables[0]);

                //  RenameFile1(folderpath, ds_gresham.Tables[0]);

                //   + "FidilityZipFile"


                if (!chkSaveSharepoint.Checked)
                    compressDirectory(folderpath, Server.MapPath("") + @"\ExcelTemplate\TempFolder\FileUpload\DownloadZipFile\" + FinalFileName);


                //lblError.Text = "Load Completed successfully";

                //  int totalfilecount = Convert.ToInt32(ViewState["pathFolderFileCount"]);

                int sucessfilecount = Convert.ToInt32(ViewState["sucessfilecount"]);

                failfilecount = Convert.ToInt32(ViewState["failfilecount"]);

                lblTotalFileCount.Text = "Total File Count: " + rowcount;

                lblSucessFileCount.Text = "Success File Count: " + sucessfilecount;

                if (sucessfilecount == rowcount)
                {
                    failfilecount = 0;
                    lblfailFileCount.Text = "Fail File Count: " + failfilecount;
                }

                else
                {
                    lblfailFileCount.Text = "Fail File Count: " + failfilecount;
                }

                if (!chkSaveSharepoint.Checked)
                {
                    lnkdownload.Visible = true;
                }
                else
                    lblError.Text = ErrorMsg + "<br/> " + strFolderExists;

                clearcontrols();
                // Download_File(Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FinalFileName, FinalFileName);

            }

        }
        catch (Exception ex)
        {
            lblError.Text = "Process Failed Please contact administrator.";
        }

        #endregion

    }

    public bool LoadList()
    {
        bool bProceed = false;
        try
        {
            // DataTable FolderData;     // clientPortal folderPath Datatable\
            DataSet dsTaxonomyclientPortal1;   // clientPortal taxonomy data
            DataTable dtSiteClientList;

            DataTable dtLEClientList;

            DataTable dtCorrespondenceType;
            DataTable dsDocumentTaxonomy;
            DataTable dtActiveClientList;

            #region New Client Services Shaepoint
            //dsDocumentTaxonomy = sp.getTaxonomyClientService();
            dtActiveClientList = sp.getActiveClientList();
            dtCorrespondenceType = sp.getTaxonomyCorrespondenceType();

            //  ViewState["dtDocumentTaxonomy"] = dsDocumentTaxonomy;
            ViewState["dtActiveClientList"] = dtActiveClientList;
            ViewState["dtCorrespondenceType"] = dtCorrespondenceType;



            #endregion
            //if (FolderData != null && dsTaxonomyclientPortal1 != null && dtSiteClientList != null && dsDocumentTaxonomy != null && dtActiveClientList != null && dtCorrespondenceType != null)
            if (dtActiveClientList != null && dtCorrespondenceType != null)
            {
                bProceed = true;
            }
            else
            {
                bProceed = false;
                lblError.Visible = true;
                lblError.Text = "Error Fetching Taxonomy List";
            }
        }
        catch (Exception Ex)
        {
            lblError.Visible = true;
            lblError.Text = "Error Fetching Taxonomy List";
            bProceed = false;
        }
        return bProceed;
    }

    public void clearcontrols()
    {
        chkSaveSharepoint.Checked = false;
        txtyear.Text = "";
    }


    public string FetchNewSpURL(DataTable dtActiveClients, string AccountId)
    {
        string SPURL = string.Empty;
        try
        {
            //Loop all the ActiveClients to match it with the Household or LegalEntity to get Path 
            for (int i = 0; i < dtActiveClients.Rows.Count; i++)
            {
                string ClientID = dtActiveClients.Rows[i]["ClientID"].ToString();
                if (AccountId.ToLower() == ClientID.ToLower())
                {
                    SPURL = dtActiveClients.Rows[i]["ClientSiteURL"].ToString();
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            return null;
        }
        return SPURL;
    }

    static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader)
    {
        string header = isFirstRowHeader ? "Yes" : "No";

        string pathOnly = Path.GetDirectoryName(path);
        string fileName = Path.GetFileName(path);

        string sql = @"SELECT * FROM [" + fileName + "]";

        using (OleDbConnection connection = new OleDbConnection(
                  @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                  ";Extended Properties=\"Text;HDR=" + header + "\""))
        using (OleDbCommand command = new OleDbCommand(sql, connection))
        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
        {
            DataTable dataTable = new DataTable();
            dataTable.Locale = CultureInfo.CurrentCulture;
            adapter.Fill(dataTable);
            return dataTable;
        }
    }


    public DataTable readcsv(string path)
    {
        DataTable dtCsv = new DataTable();

        dtCsv.Columns.Add("1");
        dtCsv.Columns.Add("2");
        dtCsv.Columns.Add("3");
        dtCsv.Columns.Add("4");
        dtCsv.Columns.Add("5");

        //dtCsv.Columns.Add(rowValues[j]); //add headers  
        string Fulltext;
        using (StreamReader sr = new StreamReader(path))
        {
            while (!sr.EndOfStream)
            {
                Fulltext = sr.ReadToEnd().ToString(); //read full file text  
                string[] rows = Fulltext.Split('\n'); //split full file text into rows  
                for (int i = 0; i < rows.Count() - 1; i++)
                {
                    string[] rowValues = rows[i].Split(','); //split each row with comma to get individual values  
                    {
                        //if (i == 0)
                        //{
                        //    for (int j = 0; j < rowValues.Count(); j++)
                        //    {
                        //        dtCsv.Columns.Add(rowValues[j]); //add headers  
                        //    }
                        //}
                        //else
                        //{
                        //    DataRow dr = dtCsv.NewRow();
                        //    for (int k = 0; k < rowValues.Count(); k++)
                        //    {
                        //        dr[k] = rowValues[k].ToString();
                        //    }
                        //    dtCsv.Rows.Add(dr); //add other rows  
                        //}


                        DataRow dr = dtCsv.NewRow();
                        for (int k = 0; k < rowValues.Count(); k++)
                        {
                            dr[k] = rowValues[k].ToString();
                        }
                        dtCsv.Rows.Add(dr); //add other rows  
                    }
                }
            }
        }

        return dtCsv;
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
                //cmd.Parameters.AddWithValue(parameter2, deleteFlg);

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
            //lblError.Text = vSqlQuery + " Failed to execute. " + ex.Message;
            //lblError.Visible = true;
            return null;
        }
    }


    private void compressDirectory(string DirectoryPath, string OutputFilePath, int CompressionLevel = 9)
    {
        try
        {
            // Depending on the directory this could be very large and would require more attention
            // in a commercial package.
            string[] filenames = Directory.GetFiles(DirectoryPath);

            // 'using' statements guarantee the stream is closed properly which is a big source
            // of problems otherwise.  Its exception safe as well which is great.
            using (ZipOutputStream OutputStream = new ZipOutputStream(System.IO.File.Create(OutputFilePath)))
            {

                // Define the compression level
                // 0 - store only to 9 - means best compression
                OutputStream.SetLevel(CompressionLevel);

                byte[] buffer = new byte[4096];

                foreach (string file in filenames)
                {

                    // Using GetFileName makes the result compatible with XP
                    // as the resulting path is not absolute.
                    ZipEntry entry = new ZipEntry(Path.GetFileName(file));

                    // Setup the entry data as required.

                    // Crc and size are handled by the library for seakable streams
                    // so no need to do them here.

                    // Could also use the last write time or similar for the file.
                    entry.DateTime = DateTime.Now;
                    OutputStream.PutNextEntry(entry);

                    using (FileStream fs = System.IO.File.OpenRead(file))
                    {

                        // Using a fixed size buffer here makes no noticeable difference for output
                        // but keeps a lid on memory usage.
                        int sourceBytes;

                        do
                        {
                            sourceBytes = fs.Read(buffer, 0, buffer.Length);
                            OutputStream.Write(buffer, 0, sourceBytes);
                        } while (sourceBytes > 0);
                    }
                }

                // Finish/Close arent needed strictly as the using statement does this automatically

                // Finish is important to ensure trailing information for a Zip file is appended.  Without this
                // the created file would be invalid.
                OutputStream.Finish();

                // Close is important to wrap things up and unlock the file.
                OutputStream.Close();

                Console.WriteLine("Files successfully compressed");
            }
        }
        catch (Exception ex)
        {
            // No need to rethrow the exception as for our purposes its handled.
            Console.WriteLine("Exception during processing {0}", ex);
        }
    }


    /// <summary>
    /// Method that compress all the files inside a folder (non-recursive) into a zip file.
    /// </summary>
    /// <param name="DirectoryPath"></param>
    /// <param name="OutputFilePath"></param>
    /// <param name="CompressionLevel"></param>


    /// <summary>
    /// Extracts the content from a .zip file inside an specific folder.
    /// </summary>
    /// <param name="FileZipPath"></param>
    /// <param name="password"></param>
    /// <param name="OutputFolder"></param>
    public void ExtractZipContent(string FileZipPath, string password, string OutputFolder)
    {
        ZipFile file = null;
        try
        {
            FileStream fs = System.IO.File.OpenRead(FileZipPath);
            file = new ZipFile(fs);

            if (!String.IsNullOrEmpty(password))
            {
                // AES encrypted entries are handled automatically
                file.Password = password;
            }

            foreach (ZipEntry zipEntry in file)
            {
                if (!zipEntry.IsFile)
                {
                    // Ignore directories
                    continue;
                }

                String entryFileName = zipEntry.Name;
                // to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
                // Optionally match entrynames against a selection list here to skip as desired.
                // The unpacked length is available in the zipEntry.Size property.

                // 4K is optimum
                byte[] buffer = new byte[4096];
                Stream zipStream = file.GetInputStream(zipEntry);

                // Manipulate the output filename here as desired.
                String fullZipToPath = Path.Combine(OutputFolder, entryFileName);
                string directoryName = Path.GetDirectoryName(fullZipToPath);
                string dirName = new DirectoryInfo(OutputFolder).Name;


                // ViewState["Foldername"] = dirName;

                if (directoryName.Length > 0)
                {
                    Directory.CreateDirectory(directoryName);
                }


                ViewState["Foldername"] = directoryName;
                // Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
                // of the file, but does not waste memory.
                // The "using" will close the stream even if an exception occurs.
                using (FileStream streamWriter = System.IO.File.Create(fullZipToPath))
                {
                    StreamUtils.Copy(zipStream, streamWriter, buffer);
                }
            }
        }
        finally
        {
            if (file != null)
            {
                file.IsStreamOwner = true; // Makes close also shut the underlying stream
                file.Close(); // Ensure we release resources
            }
        }
    }


    public void RenameFile(string path, DataTable dt)
    {

        try
        {
            string fExt;
            string fFromName;
            string ffrompath = string.Empty;
            string fToName = string.Empty; ;
            int i = 1;

            int sucessfilecount = 0;

            int failfilecount = 0;

            // string fPath = @"D:\Brijesh data\Development New\Gresham Partners\GP OnPremises Project\GP\New folder\adventReportTest\ExcelTemplate\TempFolder\FidilityZipFile\G09014033_20190128\";

            string fPath = path;

            //copy all files from fPath to files array
            FileInfo[] files = new DirectoryInfo(fPath).GetFiles();

            int fCount = Directory.GetFiles(fPath, "*", SearchOption.TopDirectoryOnly).Length;


            ViewState["pathFolderFileCount"] = fCount;

            //loop through all files
            foreach (var f in files)
            {
                //get the filename without the extension
                //  fFromName = Path.GetFileNameWithoutExtension(f.Name);
                //get the file extension
                fFromName = f.Name;
                fExt = Path.GetExtension(f.Name);


                // string fFromName = f.Name;


                foreach (DataRow dr in dt.Rows)
                {
                    try
                    {
                        string oldFilename = dr["FilenameTxt"].ToString();


                        if (fFromName == oldFilename)
                        {
                            string[] chars = new string[] { "/", "\"", "(", ")", "[", "]" };
                            // string spcailchar = ",./!@#$%^&*'\;_";


                            string NewFilename = dr["NewFilenameTxt"].ToString();
                            string SPFileNameTxt = dr["SPFileNameTxt"].ToString();
                            string SPCSUUID = dr["SPCSUUID"].ToString();
                            string LEFolderNameTxt = dr["LEFolderNameTxt"].ToString().Replace("&nbsp;", "").Replace("&#39;", "'").Replace("/", "_").Replace("#", "No.").Replace("*", "").Replace(":", "").Replace("<", "").Replace(">", "").Replace("?", "").Replace("\"", "").Replace("|", ""); ;

                            for (int j = 0; j < chars.Length; j++)
                            {
                                if (NewFilename.Contains(chars[j]))
                                {
                                    NewFilename = NewFilename.Replace(chars[j], "");
                                }

                                if (SPFileNameTxt.Contains(chars[j]))
                                {
                                    SPFileNameTxt = SPFileNameTxt.Replace(chars[j], "");
                                }

                            }



                            if (NewFilename != "" || SPFileNameTxt != "")
                            {
                                //set fFromName to the path + name of the existing file
                                //fFromName = string.Format(&quot;{0}{1}&quot;, fPath, f.Name);

                                ffrompath = fPath + "\\" + f.Name;

                                // string Frompath = fPath + "\\" + oldFilename;

                                //set the fToName as path + new name + _i + file extension
                                //fToName = string.Format("", fPath, "test", i.ToString(), fExt);

                                if (chkSaveSharepoint.Checked)
                                {
                                    fToName = fPath + "\\" + "zzTest_" + SPFileNameTxt + fExt;
                                    System.IO.File.Move(ffrompath, fToName);

                                    DataTable dtActiveClient = (DataTable)ViewState["dtActiveClientList"];

                                    string Hyperlink = string.Empty;

                                    string ClientSiteURL = sp.FetchNewSpURL(dtActiveClient, SPCSUUID);

                                    if (ClientSiteURL != "")
                                    {
                                        bool SiteExistsFlg = IsSiteExist(ClientSiteURL);

                                        if (SiteExistsFlg)
                                        {

                                            if (LEFolderNameTxt.Contains("James"))
                                                LEFolderNameTxt = "test123";

                                            bool ISLEFOlderExist = CheckLEgalEntityFolderExist(ClientSiteURL, "TaxDocuments", LEFolderNameTxt);

                                            if (ISLEFOlderExist)
                                                CopyFileinLegalEntityFolder(ClientSiteURL, "zzTest_" + SPFileNameTxt + fExt, fToName, "TaxDocuments/" + LEFolderNameTxt, "", "", "");
                                            else
                                            {
                                                Hyperlink = ClientSiteURL + "/" + "TaxDocuments";
                                                strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found for below File Name <br/> File Name: " + "zzTest_" + SPFileNameTxt + fExt + "<br/> The file is saved at the below path:" + "<br/>" + "<a href='" + Hyperlink + "' target=_blank >" + Hyperlink + "</a>";
                                                CopyFilenewCS(ClientSiteURL, "zzTest_" + SPFileNameTxt + fExt, fToName, "", "", "");
                                            }

                                            i++;
                                            sucessfilecount++;
                                        }
                                    }


                                    else
                                    {
                                        failfilecount++;
                                        ViewState["failfilecount"] = failfilecount;
                                        ErrorMsg = ErrorMsg + "<br/>File Not saved to Client Services <br/>Client Services Site not found for below File Name <br/> File Name: " + "zzTest_" + SPFileNameTxt + fExt;
                                        //lblError.Text = lblError.Text + "<br/>" + Errormsg;
                                    }
                                }
                                else
                                {
                                    fToName = fPath + "\\" + NewFilename + fExt;
                                    System.IO.File.Move(ffrompath, fToName);
                                    i++;
                                    sucessfilecount++;
                                }

                                //rename the file by moving to the same place and renaming
                                //File.Move(fFromName, fToName);


                                //increment i
                                // i++;
                                //  sucessfilecount++;
                                ViewState["sucessfilecount"] = sucessfilecount;
                            }


                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        failfilecount++;
                        ViewState["failfilecount"] = failfilecount;

                        // lblError.Text = "from Path:" + ffrompath + "FromTo" + fToName + ex.ToString();
                    }
                }


            }

            ViewState["failfilecount"] = failfilecount;
            ViewState["sucessfilecount"] = sucessfilecount;

        }
        catch (Exception ex)
        {

        }
    }





    public bool IsSiteExist(String URL)
    {
        bool bProceed = false;
        try
        {

            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;
            List list = context.Web.Lists.GetByTitle("Published Documents");

            context.Load(list);
            context.ExecuteQuery();

            bProceed = true;

        }
        catch (Exception Ex)
        {
            return false;
        }

        return bProceed;
    }


    public bool CheckFileExistinLegalEntity(string URL, string FolderNAme, string SubFolder, string FileName)
    {
        bool bProceed = false;

        try
        {
            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            ////foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            ////context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;


            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderNAme);//PublishedDocuments
                                                                                    // int count = currentRunFolder.Files.Count;
            context.Load(currentRunFolder);
            context.ExecuteQuery();

            Folder subRunFolder = currentRunFolder.Folders.GetByUrl(SubFolder); // LegalEntity
            context.Load(subRunFolder);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.File file = subRunFolder.Files.GetByUrl(FileName);
            context.Load(file);
            context.ExecuteQuery();
            bProceed = true;
        }
        catch (Exception ex)
        {
            return false;
        }
        return bProceed;
    }

    public bool CheckLEgalEntityFolderExist(string URL, string FolderNAme, string SubFolder)
    {
        bool bProceed = false;

        try
        {
            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl + "/");
            SecureString passWord = new SecureString();
            ////foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            ////context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;
            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderNAme);//PublishedDocuments
                                                                                    // int count = currentRunFolder.Files.Count;
            context.Load(currentRunFolder);
            context.ExecuteQuery();

            Folder subRunFolder = currentRunFolder.Folders.GetByUrl(SubFolder); // LegalEntity
            context.Load(subRunFolder);
            context.ExecuteQuery();

            bProceed = true;
        }
        catch (Exception ex)
        {
            //lg.AddinLogFile(LogFileName, "IsSharepointSiteExists Error" + ex.Message.ToString());
            return false;

        }
        return bProceed;
    }
    public string CopyFileinLegalEntityFolder(string URL, string destFilename, string vSourcrFile, string FolderNAme, string year, string BatchType, string Quarter)
    {
        string FileLink = string.Empty;
        try
        {
            DataTable dtCorrespondenceType = (DataTable)ViewState["dtCorrespondenceType"];

            //  DataTable dtLEList = (DataTable)ViewState["dtDocumentTaxonomy"];

            bool billingflag = Convert.ToBoolean(ViewState["BillingFlg"]);

            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            ////foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            ////context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
            System.IO.Stream stream = new System.IO.MemoryStream(bytes);

            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderNAme);

            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(destFilename), Overwrite = true };

            currentRunFolder.Files.Add(newFile);
            int count = currentRunFolder.Files.Count;
            currentRunFolder.Update();
            context.ExecuteQuery();
            Microsoft.SharePoint.Client.File upload = currentRunFolder.Files.GetByUrl(newFile.Url);
            context.Load(upload);
            context.Load(upload.ListItemAllFields);
            context.ExecuteQuery();
            Microsoft.SharePoint.Client.ListItem item = upload.ListItemAllFields;
            context.Load(item);
            context.ExecuteQuery();
            FileLink = item["FileRef"].ToString();
            string docType = string.Empty;
            string iID = string.Empty;
            TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            //Response.Write("biilingflag:" + billingflag);

            //foreach (DataRow rw in dtCorrespondenceType.Rows)
            //{
            //    if (docType.ToLower() == rw["DocumentType"].ToString().ToLower())
            //    {
            //        iID = rw["iID"].ToString();
            //        break;
            //    }
            //}
            //taxonomyFieldValuePath.TermGuid = iID;// Id from The Correspondence Taxonomy
            //taxonomyFieldValuePath.Label = docType;// "Correspondence"  OR "Quarterly Report";


            try
            {
                // item["ocac1b27043549bf95a6b3be20a5e5ea"] = taxonomyFieldValuePath; // Correspondence Type

                item["TaxDocumentType"] = ddldoctype.SelectedItem.Text.ToString();
                item["Year"] = txtyear.Text;
                item.Update();
                currentRunFolder.Update();
                context.ExecuteQuery();
            }
            catch (Exception Ex)
            {
                Response.Write("<br/>Error Occured while tagging the File: " + vSourcrFile +
                                     " to " + URL + "Tax Documents/" + FolderNAme + "<br/>" + Ex.Message + ", " + Ex.StackTrace);
            }


        }
        catch (Exception Exx)
        {
            FileLink = "";
            Response.Write("<br/>Error Occured when trying to copy file from: " + vSourcrFile +
                                      " to " + URL + "Tax Documents/" + FolderNAme + "<br/>" + Exx.Message + ", " + Exx.StackTrace);
            return "";
        }

        return FileLink;
    }

    public string CopyFilenewCS(string SiteURL, string destFilename, string vSourcrFile, string year, string BatchType, string Quarter)  // string vSourcefile, string vDestinationFile
    {
        string FileLink = string.Empty;
        string siteUrl = string.Empty;
        try
        {

            //DataSet dsTaxonomyclientPortal = (DataSet)ViewState["dsTaxonomyclientPortal"];
            //DataTable dtYear = dsTaxonomyclientPortal.Tables[2];
            //DataTable dtCorrespondenceType = (DataTable)ViewState["dtCorrespondenceType"];
            bool billingflag = Convert.ToBoolean(ViewState["BillingFlg"]);
            siteUrl = SiteURL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);
            Web site = context.Web;
            // vSourcrFile = @"D:\Infograte\Site\TEST Report Output\OPS REPORTS\zzTest_Cap Call - Masa TEST Legal Entity 2 2019-0731.pdf";

            byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
            System.IO.Stream stream = new System.IO.MemoryStream(bytes);
            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = destFilename, Overwrite = true };

            Microsoft.SharePoint.Client.List docs = context.Web.Lists.GetByTitle("Tax Documents");
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile); // NEW File to be created and uploaded 
            context.Load(uploadFile);
            context.Load(docs);
            context.ExecuteQuery();

            context.Load(uploadFile.ListItemAllFields);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.ListItem item = docs.GetItemById(uploadFile.ListItemAllFields.Id); // Fetch the Uploaded File to Tag
            context.Load(item);
            context.ExecuteQuery();

            TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            TaxonomyFieldValue taxonomyFieldValueLegalEntity = new TaxonomyFieldValue();

            FileLink = item["FileRef"].ToString();
            string docType = string.Empty;
            string iID = string.Empty;

            string LEName = string.Empty;
            string LEID = string.Empty;

            //if (BatchType.ToUpper() == "Q" || BatchType.ToUpper() == "M")
            //{
            //    docType = "Quarterly Report";
            //}
            //else if (BatchType.ToUpper() == "MERGE" && billingflag == false)
            //{
            //    //  docType = "Correspondence";
            //    docType = "Cap Call/Distribution";//added on 07/26/2019
            //}
            //else if (BatchType.ToUpper() == "MERGE" && billingflag == true)
            //{
            //    docType = "Correspondence";
            //}

            //foreach (DataRow rw in dtCorrespondenceType.Rows)
            //{
            //    if (docType.ToLower() == rw["DocumentType"].ToString().ToLower())
            //    {
            //        iID = rw["iID"].ToString();
            //        break;
            //    }
            //}

            //taxonomyFieldValuePath.TermGuid = iID;// Id from The Correspondence Taxonomy
            //taxonomyFieldValuePath.Label = docType;// "Correspondence"  OR "Quarterly Report";
            try
            {
                item["TaxDocumentType"] = ddldoctype.SelectedItem.Text;
                item["Year"] = txtyear.Text; // Year Field
                item.Update();
                docs.Update();
                context.ExecuteQuery();
            }
            catch (Exception Ex)
            {
                Response.Write("<br/>Error Occured while tagging the File: " + vSourcrFile +
                                     " to " + SiteURL + "TAx Documents" + "<br/>" + Ex.Message + ", " + Ex.StackTrace);
            }

        }
        catch (Exception Exx)
        {
            FileLink = "";
            Response.Write("<br/>Error Occured when trying to copy file from: " + vSourcrFile +
                                      " to " + SiteURL + "TAx Documents" + "<br/>" + Exx.Message + ", " + Exx.StackTrace);
            return "";

        }
        return FileLink;

    }

    public void RenameFile1(string path, DataTable dt)
    {

        try
        {
            string fExt;
            string fFromName;
            string fToName;
            int i = 1;

            int sucessfilecount = 0;

            int failfilecount = 0;

            // string fPath = @"D:\Brijesh data\Development New\Gresham Partners\GP OnPremises Project\GP\New folder\adventReportTest\ExcelTemplate\TempFolder\FidilityZipFile\G09014033_20190128\";

            string fPath = path;

            //copy all files from fPath to files array
            FileInfo[] files = new DirectoryInfo(fPath).GetFiles();

            int fCount = Directory.GetFiles(fPath, "*", SearchOption.TopDirectoryOnly).Length;


            ViewState["pathFolderFileCount"] = fCount;

            //loop through all files
            foreach (var f in files)
            {
                //get the filename without the extension
                fFromName = Path.GetFileNameWithoutExtension(f.Name);



                //get the file extension
                // fFromName = f.Name;
                fExt = Path.GetExtension(f.Name);


                // string fFromName = f.Name;


                foreach (DataRow dr in dt.Rows)
                {
                    try
                    {
                        string oldFilename = dr["FilenameTxt"].ToString();

                        string NewFilename = dr["NewFilenameTxt"].ToString();



                        if (NewFilename == oldFilename)
                        {
                            //string NewFilename = dr["NewFilenameTxt"].ToString();


                            if (NewFilename != "")
                            {
                                //set fFromName to the path + name of the existing file
                                //fFromName = string.Format(&quot;{0}{1}&quot;, fPath, f.Name);

                                fFromName = fPath + "\\" + oldFilename;

                                string Extension = Path.GetExtension(fFromName);
                                //set the fToName as path + new name + _i + file extension
                                //fToName = string.Format("", fPath, "test", i.ToString(), fExt);

                                fToName = fPath + "\\" + NewFilename + fExt;

                                //rename the file by moving to the same place and renaming
                                System.IO.File.Move(fFromName, fToName);
                                //increment i
                                i++;
                                sucessfilecount++;
                                ViewState["sucessfilecount"] = sucessfilecount;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        failfilecount++;
                        ViewState["failfilecount"] = failfilecount;
                    }
                }


            }
        }
        catch (Exception ex)
        {

        }
    }


    private void Download_File(string FilePath, string FileName)
    {
        // string _path = Request.PhysicalApplicationPath + "CV/" + name;
        System.IO.FileInfo _file = new System.IO.FileInfo(FilePath);

        Response.Clear();
        Response.AddHeader("Content-Disposition", "attachment; filename=" + _file.Name);
        Response.AddHeader("Content-Length", _file.Length.ToString());
        Response.ContentType = "application/octet-stream";
        Response.WriteFile(_file.FullName);
        Response.End();

    }


    public void DownloadFile()
    {

    }




    protected void lnkdownload_Click(object sender, EventArgs e)
    {
        string FinalFileName = Convert.ToString(ViewState["filename"]);
        Download_File(Server.MapPath("") + @"\ExcelTemplate\TempFolder\FileUpload\DownloadZipFile\" + FinalFileName, FinalFileName);
    }

    protected void chkSaveSharepoint_CheckedChanged(object sender, EventArgs e)
    {
        if (chkSaveSharepoint.Checked)
        {
            tryear.Visible = true;
            txtyear.Text = DateTime.Now.Year.ToString();
            //  trdoctype.Visible = true;
        }
        else
        {
            //  trdoctype.Visible = false;
            tryear.Visible = false;
            txtyear.Text = "";
        }
    }
}