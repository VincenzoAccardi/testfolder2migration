/* MKS Information 
 * $Source: Program.cs $
 * $ProjectName: m:/MKS/TPDotNet/TP_Customer/Italy/Common/dev.net/Pos/Tools/ProcessInitLoyalty/project.pj $
 * $ProjectRevision: Last Checkpoint: 1.1.1.1 $
 * Last modified by $Author: Pierro, Luca (luca.pierro) $
 * $Revision: 1.1 $
 * $Log: Program.cs  $
 * Revision 1.1 2014/05/07 09:59:00CEST Pierro, Luca (luca.pierro) 
 * Initial revision
 * Member added to project m:/MKS/TPDotNet/TP_Customer/Italy/Common/dev.net/Pos/Tools/ProcessInitLoyalty/project.pj
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.IO;
using System.Runtime.InteropServices;
using TPDotnet.Base.Service;


namespace ProcessInitLoyalty
{
    class 
    Program
    {

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("Kernel32")]
        private static extern IntPtr GetConsoleWindow();

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        static void Main(string[] args)
        {
            IntPtr hwnd;
            hwnd = GetConsoleWindow();
            ShowWindow(hwnd, SW_HIDE);

            Boolean logstatus = false;
            string funcName = "ProcessInitLoyalty";
            StreamWriter filelog = null;

            HttpWebRequest request = default(HttpWebRequest);
            HttpWebResponse response = default(HttpWebResponse);
            X509Certificate cert1 = default(X509Certificate);
            X509Certificate2 cert2 = default(X509Certificate2);
            X509Store store = default(X509Store);

            string path = TPDotnet.Base.Service.TPBaseSysConfig.BaseDirectory.ToString() + "\\" + "bin" + "\\";

            string pathlog = path.Replace("bin", "Log");

            if (Directory.Exists(pathlog)) 
            {
                logstatus = true;

                pathlog += "\\" + "ProcessInitLoyalty.log";

                filelog = new System.IO.StreamWriter(pathlog);
            }

            if (logstatus == true) filelog.WriteLine(DateTime.Now.ToString() + " - Start " +  funcName);

            //Console.WriteLine("Start " +  funcName);

            if (args == null)
            {
                //Console.WriteLine("args is null"); // Check for null array
                if (logstatus == true) filelog.WriteLine(DateTime.Now.ToString() + " - args is null");
            }
            else
            {
                try
                {
                    Uri uri = null;

                    //Console.Write("args length is ");
                    //Console.WriteLine(args.Length); // Write array length
                    if (logstatus == true) filelog.WriteLine(DateTime.Now.ToString() + " - args length is: " + args.Length.ToString());

                    for (int i = 0; i < args.Length; i++) // Loop through array
                    {
                        string argument = args[i];
                        switch (i)
                        {
                            case 0: // 1st Param  url
                                uri = new Uri(args[i]);
                                //Console.WriteLine("args " + i.ToString() + " value: " + args[i].ToString());
                                if (logstatus == true) filelog.WriteLine(DateTime.Now.ToString() + " - args " + i.ToString() + " value: " + args[i].ToString());
                                break;
                        }
                    }
            
              
                    ServicePointManager.ServerCertificateValidationCallback = delegate(
                    Object obj, X509Certificate certificate, X509Chain chain,
                    SslPolicyErrors errors)
                    {
                        return (true);
                    };

                    request = (HttpWebRequest)WebRequest.Create(uri);
                    request.Method = WebRequestMethods.Http.Get;
                    response = (HttpWebResponse)request.GetResponse();
                    cert1 = request.ServicePoint.Certificate;
                    cert2 = new X509Certificate2(cert1);

                    store = new X509Store(StoreName.Root, StoreLocation.LocalMachine);
                    store.Open(OpenFlags.ReadWrite);
                    store.Add(cert2);
                    store.Close();

                    if (logstatus == true) filelog.WriteLine(DateTime.Now.ToString() + " - Certificate Installed");

                }
                catch (Exception e)
                {
                    //Console.WriteLine("Exception: " + e.Message + " " + funcName);
                    if (logstatus == true) filelog.WriteLine(DateTime.Now.ToString() + " - Exception: " + e.Message);
                }
            }
            if (logstatus == true)
            {
                filelog.WriteLine(DateTime.Now.ToString() + " - End " + funcName);
                filelog.Close();
            }
            //Console.WriteLine("End " + funcName);
        }
    }
}
