using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for Logs
/// </summary>
public class Logs
{
	
    public void AddinLogFile(string vFilePath, string Data)
    {
        try
        {
            using (StreamWriter sw = File.AppendText(vFilePath))
            {
                sw.WriteLine(Data);
                sw.Close();
            }
        }
        catch
        { }

    }
}