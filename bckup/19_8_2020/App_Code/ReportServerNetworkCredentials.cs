/*****************************************************************************
 * Created By : Nizam
 * Cretaed Date : Aug 17, 2009
 * Description: This is sealed and serializable class and implementing IReportServerCredentials inteface.
 * This class Provides Network credentials to be used to connect to the report server.
 * *****************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Reporting.WebForms;
using Microsoft.ReportingServices;
using System.Security.Principal;

[Serializable]
public sealed class ReportServerNetworkCredentials : IReportServerCredentials
{
    #region IReportServerCredentials Members
    public bool GetFormsCredentials(out System.Net.Cookie authCookie, out string userName,
        out string password, out string authority)
    {
        authCookie = null;
        userName = null;
        password = null;
        authority = null;


        return false;
    }
    
    // Specifies the user to impersonate when connecting to a report server.
    //A WindowsIdentity object representing the user to impersonate.</returns>
    public WindowsIdentity ImpersonationUser
    {
        get
        {
            return null;
        }
    }
    
    // Returns network credentials to be used for authentication with the report server.
    //A NetworkCredentials object.</returns>
    public System.Net.ICredentials NetworkCredentials
    {
        get
        {
            string userName = "ztempadmin";
            string domainName = "corp";
            string password = "333Adm1n!";


            return new System.Net.NetworkCredential(userName, password, domainName);
        }
    }
    #endregion
}