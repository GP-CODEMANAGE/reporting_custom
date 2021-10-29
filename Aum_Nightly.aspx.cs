using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.ServiceModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Aum_Nightly : System.Web.UI.Page
{
    string vLogFile = string.Empty;
    DB clsDB = new DB();
    GeneralMethods clsGM = new GeneralMethods();
    IOrganizationService service = null;
    bool bProceed = true;
    int CountInsert, FailedInsert = 0;
    int CountUpdate,FailedUpdate = 0;
    protected void Page_Load(object sender, EventArgs e)
    {
        bool isScheduled = false;
        if (!IsPostBack)
        {

            #region LogFile

            if (!Directory.Exists(HttpContext.Current.Server.MapPath("") + @"\Logs"))
            {

                System.IO.Directory.CreateDirectory(HttpContext.Current.Server.MapPath("") + @"\Logs");
            }
            string DateTimeValue = DateTime.Now.ToString("yyyyMMddHHmmss");
            vLogFile = HttpContext.Current.Server.MapPath("") + @"\Logs\" + @"\Log_" + DateTimeValue + ".txt";


            CreateLogFile(vLogFile);
            #endregion

            isScheduled = (Request.QueryString["id"] ?? String.Empty).Equals("Scheduled");
             isScheduled = true;
            try
            {
                if (isScheduled)
                {

                    // AddinLogFile(vLogFile, "++++++++++++++++++++++++++++++++Scheduled Process+++++++++++++++++++++++++++++++: " + DateTime.Now);

                    CRMConnection();
                    string sqlstr = "SP_U_HH_AUM";
                    DataSet AUMDataset = clsDB.getDataSet(sqlstr);

                    InsertAUM(AUMDataset);
                    AddinLogFile(vLogFile, "Insert Completed, Count: " + CountInsert);

                    UpdateAUM(AUMDataset);
                    AddinLogFile(vLogFile, "Updated Completed, Count: " + CountUpdate);

                    SendEmail("Processs Completed ,Record Updated", "AUM Updated - Nightly", "skane@infograte.com", "");

                    AddinLogFile(vLogFile, "Email SEnt " + DateTime.Now);
                    AddinLogFile(vLogFile, "Processs Completed "+DateTime.Now);
                }
            }
            catch (Exception ex)
            {
                AddinLogFile(vLogFile, "Error: " + ex.InnerException.ToString() + DateTime.Now);
            }
            //}
        }
    }
    public void SendEmail(string mailmessage, string subject, string mailTo, string Attachment1)
    {
        try
        {

            MailMessage myMessage = new MailMessage();
            // SmtpClient SMTPSERVER = new SmtpClient();

            string EmailID = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Password = AppLogic.GetParam(AppLogic.ConfigParam.Password);
            string SMTPHost = AppLogic.GetParam(AppLogic.ConfigParam.SMTPHost);
           // string ToEmailIDs = AppLogic.GetParam(AppLogic.ConfigParam.ToEmailIDs);
            int Port = Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.Port));

           // AddinLogFile(vLogFile, " EmailID" + EmailID + "Password" + Password + "SMTPHost" + SMTPHost + "ToEmailIDs" + ToEmailIDs + DateTime.Now);

            myMessage.From = new MailAddress(EmailID, "AUM Nightly Process COMPLETED");

            string ToEmailIDs = "auto-emails@infograte.com|skane@infograte.com";
            
            string[] strTo = ToEmailIDs.Split('|');


            for (int i = 0; i < strTo.Length; i++)
            {
                if (strTo[i] != "")
                {
                    myMessage.To.Add(new MailAddress(strTo[i]));
                }
            }
 
            //myMessage.Bcc.Add("auto-emails@infograte.com");
            //myMessage.Bcc.Add("skane@infograte.com");
      
            myMessage.Subject = subject;

            //if (Attachment1 != "")
            //    myMessage.Attachments.Add(new Attachment(Attachment1));
           
            mailmessage = "Nightly AUM Process Completed";

            mailmessage = mailmessage + "<br/><br/>Aum Inserted : " + CountInsert + "<br/>AUM Updated :" + CountUpdate + "<br/>";
            if (FailedInsert > 0)
            {
                mailmessage = mailmessage + "<br/>Aum Failed To Inserted : " + FailedInsert;
            }
            if (FailedUpdate > 0)
            {
                mailmessage = mailmessage + "<br/>Aum Failed To Update: " + FailedUpdate;
            }
                       
            myMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
            myMessage.Body = mailmessage;

            myMessage.IsBodyHtml = true;

            SmtpClient SMTPSERVER = new SmtpClient(SMTPHost, Port);
            SMTPSERVER.DeliveryMethod = SmtpDeliveryMethod.Network;
            //SMTPSERVER.Host = SMTPHost;
            //SMTPSERVER.Port = Port;


            SMTPSERVER.EnableSsl = true;
            // smtp.EnableSsl = true;
            SMTPSERVER.UseDefaultCredentials = true;
            System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(EmailID, Password);
            SMTPSERVER.Credentials = basicAuthenticationInfo;
            SMTPSERVER.Send(myMessage);

            myMessage.Dispose();
            myMessage = null;
            SMTPSERVER = null;
        }
        catch (Exception ex)
        {
            string strDescription = "Error sending MAil :" + ex.Message.ToString();
            AddinLogFile(vLogFile, strDescription + DateTime.Now);
        }
    }
    public void CRMConnection()
    {
        try
        {
            //service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            AddinLogFile(vLogFile, "Crm Service starts successfully " + DateTime.Now);

        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            AddinLogFile(vLogFile, "Crm Service failed to start, Error Detail: " + exc.Detail.Message + DateTime.Now);
        }
        catch (Exception exc)
        {
            bProceed = false;
            AddinLogFile(vLogFile, "Crm Service failed to start, Error Detail: " + exc.Message + DateTime.Now);

        }

    }
    public void UpdateAUM(DataSet AUMDataset)
    {
        string ssi_householdaumsId = string.Empty;
        string AumTIA = string.Empty;
        string AumGA = string.Empty;
        string AsOfDate = string.Empty;
        try
        {

            for (int j = 0; j < AUMDataset.Tables[1].Rows.Count; j++)
            {
                ssi_householdaumsId = Convert.ToString(AUMDataset.Tables[1].Rows[j]["ssi_householdaumsId"]);
               
                AumTIA = Convert.ToString(AUMDataset.Tables[1].Rows[j]["AumTIA"]);
                AumGA = Convert.ToString(AUMDataset.Tables[1].Rows[j]["AumGA"]);
                AsOfDate = Convert.ToString(AUMDataset.Tables[1].Rows[j]["AsOfDate"]);


                if (ssi_householdaumsId != "")
                {
                    Entity objAccount = new Entity("ssi_householdaums");

                    if (ssi_householdaumsId != "")
                    {
                        objAccount["ssi_householdaumsid"] = new Guid((ssi_householdaumsId));
                    }
                    if (AumTIA != "")
                    {
                        objAccount["ssi_aumtia"] = new Money(Convert.ToDecimal(AumTIA));
                    }
                    if (AumGA != "")
                    {
                        objAccount["ssi_aumga"] = new Money(Convert.ToDecimal(AumGA));
                    }
                    if (AsOfDate != "")
                    {
                        objAccount["ssi_aumdate"] = Convert.ToDateTime(Convert.ToString(AsOfDate)); ;
                    }
                    service.Update(objAccount);

                    CountUpdate++;
                }
            }
        }
        catch (Exception ex)
        {
            FailedUpdate++;
            AddinLogFile(vLogFile, "Error Updating Aum for Accountd: " + ssi_householdaumsId + DateTime.Now);
            AddinLogFile(vLogFile, "Error: " + ex.Message + DateTime.Now);
        }
    }
    public void InsertAUM(DataSet AUMDataset)
    {
        string AccountId = string.Empty;
        string AumTIA = string.Empty;
        string AumGA = string.Empty;
        string AsOfDate = string.Empty;

        try
        {

            for (int j = 0; j < AUMDataset.Tables[0].Rows.Count; j++)
            {
                AccountId = Convert.ToString(AUMDataset.Tables[0].Rows[j]["HouseHoldId"]);
                AumTIA = Convert.ToString(AUMDataset.Tables[0].Rows[j]["AumTIA"]);
                AumGA = Convert.ToString(AUMDataset.Tables[0].Rows[j]["AumGA"]);
                AsOfDate = Convert.ToString(AUMDataset.Tables[0].Rows[j]["AsOfDate"]);

                if (AccountId != "")
                {
                    Entity objAccount = new Entity("ssi_householdaums");

                    if (AccountId != "")
                    {
                        objAccount["ssi_householdid"] = new EntityReference("account", new Guid((AccountId)));
                         
                    }
                    if (AumTIA != "")
                    {
                        objAccount["ssi_aumtia"] = new Money(Convert.ToDecimal(AumTIA));
                    }
                    if (AumGA != "")
                    {
                        objAccount["ssi_aumga"] = new Money(Convert.ToDecimal(AumGA));
                    }
                    if (AsOfDate != "")
                    {
                        objAccount["ssi_aumdate"] = Convert.ToDateTime(Convert.ToString(AsOfDate)); ;
                    }
                     
                     service.Create(objAccount);
                   CountInsert++;
                }
            }
        }
        catch (Exception ex)
        {
            FailedInsert++;
            AddinLogFile(vLogFile, "Error Inserting Aum for Accountd: " + AccountId + DateTime.Now);
            AddinLogFile(vLogFile, "Error: " + ex.Message + DateTime.Now);
        }
    }

    public void CreateLogFile(string vFilePath)
    {

        string vFilePathOnly = Path.GetDirectoryName(vFilePath);
        // string FileName = Path.GetFileName(vFilePath);
        if (!Directory.Exists(vFilePathOnly))
        {
            Directory.CreateDirectory(vFilePathOnly);
        }

        File.CreateText(vFilePath).Dispose();
        string contain = "--------Application Start-" + DateTime.Now.ToString("yyyyMMddHHmmss") + "-------";

        using (StreamWriter sw = File.AppendText(vFilePath))
        {
            sw.WriteLine(contain);
            sw.Close();
        }
    }
    public void AddinLogFile(string vFilePath, string Data)
    {

        using (StreamWriter sw = File.AppendText(vFilePath))
        {
            sw.WriteLine(Data);
            sw.Close();
        }

    }
}