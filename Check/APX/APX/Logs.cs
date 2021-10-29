using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APX
{
    class Logs
    {
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
}
