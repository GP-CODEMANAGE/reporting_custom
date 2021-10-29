using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections;
using System.Xml;
using System.IO;

/// <summary>
/// Summary description for AppLogic
/// </summary>
/// 
public static class AppLogic
{
    private static string strUseHashTable = "no";
    private static Hashtable __ConfigParams = null;
    private static readonly string __ParamsPath = "Params.config";

    public enum ConfigParam
    {
        
        CRMServerurl,
        OpsReports,
        BatchReports,
        OutPutReports,
        ApprovedReports,
        SharePointReports,
        ClientReports,
        DBConnectionstring,
        OpsReporting,
        CompletedMailings,
        CombinedPdfs,
        QuarteryReports,
        DBTransactions,
        FileUploadUrl,
        ImagePath,
	    PerformanceFilePath,
        SLOAPDFPath,
        AxysToolFilePath
        , DistributeRecoFilePath
        , CRMUrl
        ,CoverLetter
        ,DTSFilePath

        ,SMTPHost
        ,Port
        ,EmailId
        ,Password
        ,ToEmailIDs
        ,ToEmailIDs1
        ,CapitalCallFilePath
        , CRM2016WebAPI
        , SpireLicense
        , SummitasPath
        ,gbhaiaPassword
        , CRMPassword
        , SubscriptionLetters
            ,UserName
            , CRMPortNumber
            , CapCallBackupPath
            , DistributionBackupPath
            ,Server
            , ToEmailIDs2
        , ToEmailIDbillingReject
            , clientportalURL
             , clientservURL
            , SharepointURL
            , SharepointCSURL
           
            , httpclientservURL
            , httpSharepointURL
            , SPUserEmailID
            , SPUserPassword
            , DataBackupURL
    }

    static AppLogic()
    {
        __ConfigParams = new Hashtable();
        if (HttpContext.Current != null)
        {
            __ParamsPath = HttpContext.Current.Server.MapPath(__ParamsPath);
        }

    }

    public static void InitParams()
    {
        if (File.Exists(__ParamsPath))
        {
            XmlNodeList __paramNodes = _GetParamNodeList("Params/Param");
            try
            {
                if (__paramNodes != null && __paramNodes.Count > 0)
                {
                    foreach (XmlNode __paramNode in __paramNodes)
                    {
                        string __Key = __paramNode.Attributes["key"].Value;
                        string __Value = __paramNode.Attributes["value"].Value;
                        if (!__ConfigParams.Contains(__Key))
                        {
                            __ConfigParams.Add(__Key, __Value);
                        }
                    }
                }
            }
            finally
            {
                if (__paramNodes != null)
                {
                    __paramNodes = null;
                }
            }
        }
    }

    public static string GetParam(AppLogic.ConfigParam _param)
    {
        string _strValue = string.Empty;
        if (strUseHashTable == "yes")
        {
            string _strParam = _param.ToString();

            if (__ConfigParams.Contains(_strParam))
            {
                return __ConfigParams[_strParam].ToString();
            }
            return null;
        }
        else
        {
            if (File.Exists(__ParamsPath))
            {
                XmlNodeList __paramNodes = _GetParamNodeList("Params/Param");
                try
                {
                    if (__paramNodes != null && __paramNodes.Count > 0)
                    {
                        foreach (XmlNode __paramNode in __paramNodes)
                        {
                            string __Key = __paramNode.Attributes["key"].Value;
                            string __Value = __paramNode.Attributes["value"].Value;

                            if (_param.ToString() == __Key)
                            {
                                _strValue = __Value;
                                break;
                            }
                        }
                    }

                }
                finally
                {
                    if (__paramNodes != null)
                    {
                        __paramNodes = null;
                    }
                }

            }

            return _strValue;

        }
    }



    public static void SetParam(AppLogic.ConfigParam _param, string val)
    {
        //lock (AppLogic)
        //{
        if (strUseHashTable == "Yes")
        {
            string _strParam = _param.ToString();
            if (__ConfigParams.Contains(_param.ToString()))
            {
                __ConfigParams[_strParam] = val;
            }
            else
            {
                __ConfigParams.Add(_strParam, val);
            }
            AppLogic.__SetParam(_param, val);
        }
        else
        {
            AppLogic.__SetParam(_param, val);
        }

        //}
    }

    #region Private Methods

    /// <summary>
    /// Get Param node list based on XPath expression
    /// </summary>
    /// <param name="xPathExp">Valid XPath expression</param>
    /// <returns>returns param node list</returns>
    private static XmlNodeList _GetParamNodeList(string xPathExp)
    {
        XmlDocument paramDoc = null;
        XmlNodeList paramNodes = null;
        try
        {
            paramDoc = new XmlDocument();
            paramDoc.Load(__ParamsPath);
            paramNodes = paramDoc.SelectNodes(xPathExp);
        }
        finally
        {
            if (paramDoc != null)
            {
                paramDoc = null;
            }
        }
        return paramNodes;
    }

    private static void __SetParam(AppLogic.ConfigParam _param, string val)
    {
        XmlDocument paramDoc = null;
        XmlNodeList paramNodes = null;
        try
        {
            paramDoc = new XmlDocument();
            paramDoc.Load(__ParamsPath);
            paramNodes = paramDoc.SelectNodes("Params/Param[@key=\"" + _param.ToString() + "\"]");
            if (paramNodes != null && paramNodes.Count > 0)
            {
                paramNodes[0].Attributes["value"].Value = val;

                //If its readonly set it back to normal
                //Need to "AND" it as it can also be archive, hidden etc 
                if ((File.GetAttributes(__ParamsPath) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    File.SetAttributes(__ParamsPath, FileAttributes.Normal);
                }

                paramDoc.Save(__ParamsPath);
            }
        }
        finally
        {
            if (paramDoc != null)
            {
                paramDoc = null;
            }
        }
    }
    #endregion

}