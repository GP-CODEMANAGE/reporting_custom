using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security;
using System.Web;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

/// <summary>
/// Summary description for sharepoint
/// </summary>
public class sharepoint
{
    public sharepoint()
    {

    }
    public DataTable getSiteClientList()
    {
        //string vNewSharePointReportFolder = "Documents taxonomy";
        //string vSourcrFile = @"E:\Log.txt";

        try
        {
            //  string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";//"https://greshampartners.sharepoint.com";
            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            DataTable dt = new DataTable();
            dt.Columns.Add("ClientName");
            dt.Columns.Add("OnPortal");
            dt.Columns.Add("ClientPortalName");
            dt.Columns.Add("iID");

            List list = site.Lists.GetByTitle("Client");
            CamlQuery caml = new CamlQuery();
            Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(caml);
            context.Load(list);
            context.Load(items);
            context.ExecuteQuery();
            foreach (Microsoft.SharePoint.Client.ListItem item in items)
            {
                context.Load(item);
                context.ExecuteQuery();
                string ClientName = string.Empty;
                string OnPortal = string.Empty;
                string iID = string.Empty;
                string ClientPortalName = string.Empty;

                DataRow dr = dt.NewRow();
                OnPortal = item["OnPortal"].ToString();
                ClientName = item["Title"].ToString();
                iID = item["h6ed"].ToString();
                try
                {
                    ClientPortalName = item["ClientPortal"].ToString();
                    dr["ClientPortalName"] = ClientPortalName;

                }
                catch
                { }

                dr["ClientName"] = ClientName;
                dr["OnPortal"] = OnPortal;
                dr["iID"] = iID;

                dt.Rows.Add(dr);

            }

            return dt;
        }
        catch
        {
            return null;
        }

    }


    public DataTable getSPList()
    {
        try
        {
            //  string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);



            Web site = context.Web;

            List list = site.Lists.GetByTitle("CP Mapping");
            CamlQuery caml = new CamlQuery();
            Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(caml);
            context.Load(list);
            context.Load(items);
            context.ExecuteQuery();

            DataTable dt = new DataTable();
            dt.Columns.Add("FolderPath");
            dt.Columns.Add("OnPortal");
            dt.Columns.Add("Tag");

            foreach (Microsoft.SharePoint.Client.ListItem item in items) //OnPortal     ClientPortal    
            {
                context.Load(item);
                context.ExecuteQuery();
                string Folderpath = string.Empty;
                string OnPortal = string.Empty;
                string Tag = string.Empty;

                OnPortal = item["On_x0020_Portal"].ToString();
                Folderpath = item["Title"].ToString();
                Tag = item["_x0070_bi4"].ToString();

                DataRow dr = null;
                dr = dt.NewRow();
                dr["FolderPath"] = Folderpath;
                dr["OnPortal"] = OnPortal;
                dr["Tag"] = Tag;

                dt.Rows.Add(dr);

            }
            return dt;
        }
        catch
        {
            return null;
        }

    }


    public DataSet getTaxonomyClientPortal()
    {
        try
        {
            //   string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            // foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);

            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);


            Web site = context.Web;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

            // TermStore termStore = taxonomySession.TermStores.GetByName("Taxonomy_4BLQTxDzt3+F9JB2YxRRiQ=="); //commented 1_10_2019

            Guid SharepointTermStoreID = new Guid("8f0311806e7c4e72aa9d55d7cf0d8400");
            TermStore termStore = taxonomySession.TermStores.GetById(SharepointTermStoreID);

            TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);
            TermGroup termGroup1 = termStore.Groups.GetByName("Client Portal");  //GUID = {94c3c53d-2351-3b5e-bfcb-c4f1b941157c}
            TermGroup tgClientName = termStore.Groups.GetByName("Client Name");
            TermGroup tgYear = termStore.Groups.GetByName("Year");

            TermSet termsetClientName = tgClientName.TermSets.GetByName("Client Name");
            TermSet termSetDocumentType = termGroup1.TermSets.GetByName("Document Type");
            TermSet termSetYear = tgYear.TermSets.GetByName("Year");

            TermCollection tcClientName = termsetClientName.GetAllTerms();
            TermCollection tcDocType = termSetDocumentType.GetAllTerms();
            TermCollection tcyear = termSetYear.GetAllTerms();

            context.Load(taxonomySession);
            context.Load(termStore);
            context.Load(termGroup);
            context.Load(termGroup1);
            context.Load(tgClientName);


            context.Load(termsetClientName);
            context.Load(termSetDocumentType);
            context.Load(termSetYear);

            context.Load(tcClientName);
            context.Load(tcDocType);
            context.Load(tcyear);

            context.ExecuteQuery();


            DataTable dtClient = new DataTable();
            DataTable dtDocumentType = new DataTable();
            DataTable dtYear = new DataTable();

            dtClient.Columns.Add("clientName");
            dtClient.Columns.Add("iID");
            dtDocumentType.Columns.Add("DocumentType");
            dtDocumentType.Columns.Add("iID");
            dtYear.Columns.Add("Year");
            dtYear.Columns.Add("iID");

            foreach (Term ts in tcClientName)
            {
                DataRow row = dtClient.NewRow();
                row["clientName"] = ts.Name;
                row["iID"] = ts.Id.ToString();
                string id = ts.Id.ToString();
                dtClient.Rows.Add(row);
            }

            foreach (Term ts in tcDocType)
            {
                DataRow row = dtDocumentType.NewRow();
                row["DocumentType"] = ts.Name;
                row["iID"] = ts.Id.ToString();
                dtDocumentType.Rows.Add(row);

                string id = ts.Id.ToString();

            }
            //Response.Write("<br/>,");
            foreach (Term ts in tcyear)
            {
                DataRow row = dtYear.NewRow();
                row["Year"] = ts.Name;
                row["iID"] = ts.Id.ToString();
                dtYear.Rows.Add(row);

            }

            DataSet dsTaxonomy = new DataSet();
            dsTaxonomy.Tables.Add(dtClient);
            dsTaxonomy.Tables.Add(dtDocumentType);
            dsTaxonomy.Tables.Add(dtYear);

            return dsTaxonomy;
        }
        catch(Exception ex)
        {
            return null;
        }
    }

    public DataTable getTaxonomyCorrespondenceType()
    {
        try
        {
            // string siteUrl = "https://greshampartners.sharepoint.com";
            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            // foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            // foreach (var c in "51ngl3malt") passWord.AppendChar(c);
            // context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

            Guid SharepointTermStoreID = new Guid("8f0311806e7c4e72aa9d55d7cf0d8400");//Taxonomy
            TermStore termStore = taxonomySession.TermStores.GetById(SharepointTermStoreID);

            TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);

           // TermGroup termGroup1 = termStore.Groups.GetById(new Guid("1c101ae4-1c74-4156-bb13-1ce1c40a48c4"));// GetByName("Client Portal");

            TermGroup termGroup1 = termStore.Groups.GetById(new Guid("7029533f-1d78-4f49-ac9a-1233e0588bad"));// GetByName("Client Portal");
            TermSet termSetDocumentType = termGroup1.TermSets.GetById(new Guid("7d099b4b-c7b3-4a4f-a133-af4b88d764f1"));//GetByName("Correspondence Type");
           // TermSet termSetDocumentType = termGroup1.TermSets.GetById(new Guid("5a505b98-f14a-4540-8d0f-277e44ccaed1"));//GetByName("Correspondence Type");

            TermCollection tcDocType = termSetDocumentType.GetAllTerms();

            context.Load(taxonomySession);
            context.Load(termStore);
            context.Load(termGroup);
            context.Load(termGroup1);

            context.Load(termSetDocumentType);
            context.Load(tcDocType);

            context.ExecuteQuery();

            DataTable dtDocumentType = new DataTable();
            dtDocumentType.Columns.Add("DocumentType");
            dtDocumentType.Columns.Add("iID");

            foreach (Term ts in tcDocType)
            {
                if (ts.Name != "" && ts.Id != null)
                {
                    DataRow row = dtDocumentType.NewRow();
                    row["DocumentType"] = ts.Name;
                    row["iID"] = ts.Id.ToString();
                    dtDocumentType.Rows.Add(row);

                    string id = ts.Id.ToString();
                }
            }


            return dtDocumentType;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    public DataSet getTaxonomyClientService1()
    {
        try
        {
            // string siteUrl = "https://greshampartners.sharepoint.com";
            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            // foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            // foreach (var c in "51ngl3malt") passWord.AppendChar(c);
            // context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

            Guid SharepointTermStoreID = new Guid("8f0311806e7c4e72aa9d55d7cf0d8400"); // Taxonomy 
            TermStore termStore = taxonomySession.TermStores.GetById(SharepointTermStoreID);


            TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);
            TermGroup termGroup1 = termStore.Groups.GetById(new Guid("1c101ae4-1c74-4156-bb13-1ce1c40a48c4"));// GetByName("Client Portal"); 

            TermSet ClientswithLegalEntities = termGroup1.TermSets.GetById(new Guid("136c3c5b-c86e-4aad-aa58-b3ff46dfaadb"));//GetByName("Clients with Legal Entities");

            TermCollection tcClientswithLegalEntities = ClientswithLegalEntities.GetAllTerms();

            context.Load(taxonomySession);
            context.Load(termStore);
            context.Load(termGroup);
            context.Load(termGroup1);

            context.Load(ClientswithLegalEntities);

            context.Load(tcClientswithLegalEntities);

            context.ExecuteQuery();

            //Create DataTable
            DataTable dtTaxonomy = new DataTable();
            DataTable dtLegalEntity = new DataTable();

            //Add columns to Datatable
            dtTaxonomy.Columns.Add("TaxonomyName");
            dtTaxonomy.Columns.Add("TaxonomyID");
            dtTaxonomy.Columns.Add("TaxonomyKey");
            dtTaxonomy.Columns.Add("TaxonomyValue");

            dtLegalEntity.Columns.Add("TaxonomyName");
            dtLegalEntity.Columns.Add("TaxonomyID");
            dtLegalEntity.Columns.Add("TaxonomyKey");
            dtLegalEntity.Columns.Add("TaxonomyValue");

            //Loop all the terms found 
            foreach (Term ts in tcClientswithLegalEntities)
            {

                Dictionary<string, string> dicNumFilesCount = new Dictionary<string, string>();

                string val = string.Empty;
                //Loop all the custom properties and fetch just the Household and legalentity
                foreach (KeyValuePair<string, string> property in ts.CustomProperties)
                {

                    val = property.Key;

                    if (val.ToLower() == "householduuid")
                    {
                        DataRow row = dtTaxonomy.NewRow();
                        row["TaxonomyName"] = ts.Name;
                        row["TaxonomyID"] = ts.Id.ToString();
                        row["TaxonomyKey"] = property.Key;
                        row["TaxonomyValue"] = property.Value;
                        dtTaxonomy.Rows.Add(row);
                    }
                    else if (val.ToLower() == "legalentityuuid")
                    {
                        DataRow row = dtLegalEntity.NewRow();
                        row["TaxonomyName"] = ts.Name;
                        row["TaxonomyID"] = ts.Id.ToString();
                        row["TaxonomyKey"] = property.Key;
                        row["TaxonomyValue"] = property.Value;
                        dtLegalEntity.Rows.Add(row);
                    }

                }

            }

            DataSet dsTaxonomy = new DataSet();
            dsTaxonomy.Tables.Add(dtTaxonomy);
            dsTaxonomy.Tables.Add(dtLegalEntity);

            return dsTaxonomy;
        }
        catch (Exception ex)
        {
            return null;
        }
    }
    public DataTable getActiveClientList()
    {
        try
        {
            //   string siteUrl = "https://greshampartners.sharepoint.com/sites/CS/";
            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.SharepointCSURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //  foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            // List list = site.Lists.GetByTitle("Clients"); // Fetch Active client Lists

            List list = site.Lists.GetById(new Guid("0e6461ee-bb7e-49c2-9db8-c0f14e27f86a"));  // Fetch Active client Lists
            CamlQuery caml = new CamlQuery();
            Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(caml);
            context.Load(list);
            context.Load(items);
            context.ExecuteQuery();

            //Create custom table and add columns
            DataTable dt = new DataTable();
            dt.Columns.Add("Client");
            dt.Columns.Add("ClientSiteURL");
            dt.Columns.Add("ClientID");

            foreach (Microsoft.SharePoint.Client.ListItem item in items) //ActiveClients
            {
                //context.Load(item);
                //context.ExecuteQuery();
                string Client = string.Empty;
                string ClientSiteURL = string.Empty;
                string ClientID = string.Empty;
                try
                {
                    ClientSiteURL = ((Microsoft.SharePoint.Client.FieldUrlValue)(item["ClientSiteURL"])).Url;
                }
                catch (Exception ex)
                {
                    ClientSiteURL = "";
                }
                try
                {
                    Client = ((Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue)(item["Client"])).Label;

                }
                catch (Exception ex)
                {
                    Client = "";
                }
                try
                {
                    ClientID = item["UUID"].ToString();

                }
                catch (Exception ex)
                {
                    ClientID = "";
                }

                if (ClientSiteURL != "" && Client != "" && ClientID != "")
                {
                    DataRow dr = null;
                    dr = dt.NewRow();
                    dr["ClientSiteURL"] = ClientSiteURL;
                    dr["Client"] = Client;
                    dr["ClientID"] = ClientID;
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }
        catch (Exception ex)
        {
            return null;
        }

    }
    public string FetchSharepointLink(DataTable dtTaxonomy, DataTable dtActiveClients, string AccountId)
    {
        string SPLink = string.Empty;
        try
        {
            //Loop all the Taxonomyfetched and match the Household or LegalEntity. //ClientPortal--->Client with Legal Entities
            for (int i = 0; i < dtTaxonomy.Rows.Count; i++)
            {
                string TaxonomyID = dtTaxonomy.Rows[i]["TaxonomyValue"].ToString();

                if (TaxonomyID.ToLower() == AccountId.ToLower())
                {
                    string TaxonomyName = dtTaxonomy.Rows[i]["TaxonomyName"].ToString();
                    SPLink = FetchSpURL(TaxonomyName, dtActiveClients);
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            return null;
        }
        return SPLink;
    }
    public string FetchSpURL(string TaxonomyName, DataTable dtActiveClients)
    {
        string SPURL = string.Empty;
        try
        {
            //Loop all the ActiveClients to match it with the Household or LegalEntity to get Path 
            for (int i = 0; i < dtActiveClients.Rows.Count; i++)
            {
                string ClientName = dtActiveClients.Rows[i]["Client"].ToString();
                if (TaxonomyName.ToLower() == ClientName.ToLower())
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
    public DataTable getTaxonomyClientService()
    {
        try
        {
            //  string siteUrl = "https://greshampartners.sharepoint.com";
            string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL);
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            // foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //  context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
            // foreach (var c in "51ngl3malt") passWord.AppendChar(c);
            // context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

            Guid SharepointTermStoreID = new Guid("8f0311806e7c4e72aa9d55d7cf0d8400"); // Taxonomy 
            TermStore termStore = taxonomySession.TermStores.GetById(SharepointTermStoreID);



            TermGroup termGroup = termStore.GetSiteCollectionGroup(context.Site, false);
            TermGroup termGroup1 = termStore.Groups.GetById(new Guid("1c101ae4-1c74-4156-bb13-1ce1c40a48c4"));// GetByName("Client Portal"); 

            TermSet ClientswithLegalEntities = termGroup1.TermSets.GetById(new Guid("136c3c5b-c86e-4aad-aa58-b3ff46dfaadb"));//GetByName("Clients with Legal Entities");

            TermCollection tcClientswithLegalEntities = ClientswithLegalEntities.GetAllTerms();

            context.Load(taxonomySession);
            context.Load(termStore);
            context.Load(termGroup);
            context.Load(termGroup1);

            context.Load(ClientswithLegalEntities);

            context.Load(tcClientswithLegalEntities);

            context.ExecuteQuery();

            //Create DataTable
            DataTable dtTaxonomy = new DataTable();
            DataTable dtLegalEntity = new DataTable();

            //Add columns to Datatable
            dtTaxonomy.Columns.Add("TaxonomyName");
            dtTaxonomy.Columns.Add("TaxonomyID");
            dtTaxonomy.Columns.Add("TaxonomyKey");
            dtTaxonomy.Columns.Add("TaxonomyValue");

            //dtLegalEntity.Columns.Add("TaxonomyName");
            //dtLegalEntity.Columns.Add("TaxonomyID");
            //dtLegalEntity.Columns.Add("TaxonomyKey");
            //dtLegalEntity.Columns.Add("TaxonomyValue");

            //Loop all the terms found 
            foreach (Term ts in tcClientswithLegalEntities)
            {

                Dictionary<string, string> dicNumFilesCount = new Dictionary<string, string>();

                string val = string.Empty;
                //Loop all the custom properties and fetch just the Household and legalentity
                foreach (KeyValuePair<string, string> property in ts.CustomProperties)
                {

                    val = property.Key;

                    if (val.ToLower() == "householduuid" || val.ToLower() == "legalentityuuid")
                    {
                        DataRow row = dtTaxonomy.NewRow();
                        row["TaxonomyName"] = ts.Name;
                        row["TaxonomyID"] = ts.Id.ToString();
                        row["TaxonomyKey"] = property.Key;
                        row["TaxonomyValue"] = property.Value;
                        dtTaxonomy.Rows.Add(row);
                    }

                }

            }

            // DataSet dsTaxonomy = new DataSet();
            //  dsTaxonomy.Tables.Add(dtTaxonomy);
            // dsTaxonomy.Tables.Add(dtLegalEntity);

            return dtTaxonomy;
        }
        catch (Exception ex)
        {
            return null;
        }
    }

}