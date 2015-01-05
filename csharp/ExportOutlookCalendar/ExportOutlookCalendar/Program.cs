using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Collections.Specialized;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;

namespace ExportOutlookCalendar
{
    class Program
    {

        static void Main(string[] args)
        {

            int waitMinutes = 5;
            //int waitSeconds = 10;

            while (true)
            {

                writeOutllokToIcs("C:/Users/Public/export.ics");

                pushPost(encodeFile("C:/Users/Public/export.ics"));

                Console.WriteLine("Pushed ICS");

                Thread.Sleep(waitMinutes * 60 * 1000);
                //Thread.Sleep(waitSeconds * 1000);
            }

        }

        private static string encodeFile(string filePath)
        {
            string base64 = "";
            if (!string.IsNullOrEmpty(filePath))
            {
                FileStream fs = new FileStream(filePath,
                                               FileMode.Open,
                                               FileAccess.Read);
                byte[] filebytes = new byte[fs.Length];
                fs.Read(filebytes, 0, Convert.ToInt32(fs.Length));
                string encodedData =
                    Convert.ToBase64String(filebytes,
                                           Base64FormattingOptions.InsertLineBreaks);
                base64 = encodedData;

                fs.Close();

            }

            return base64;

        }


        private static int pushPost(string context)
        {
            int errorCode = 0;

            // hotfix
            


            WebClient wb = new WebClient();

            //Validate proxy address
            var proxyURI = new Uri("http://prx-fraint-v05.inet.cns.fra.dlh.de:8080");

            //Set credentials
            ICredentials credentials = new NetworkCredential("USERNAME", "PASSWORD");

            //Set proxy
            WebProxy wp = new WebProxy(proxyURI, false, null, credentials);

            //WebProxy wp = new WebProxy(" proxy server url here");
            wb.Proxy = wp;
            

            var data = new NameValueCollection();
            data["calendarFile"] = context;
            //data["password"] = "myPassword";
            try
            {
                System.Net.ServicePointManager.Expect100Continue = false;
                var response = wb.UploadValues("http://yourdomain/cal/calexport.php", "POST", data);
            }
            catch (System.Exception e)
            {
                Console.WriteLine("ERROR: "+e.Message);
            }
            

            //Console.WriteLine(wb.Encoding.GetString(response));

            return errorCode;

        }


        private static void writeOutllokToIcs(string calendarFileName)
        {

            Microsoft.Office.Interop.Outlook.Application oApp = null;
            Microsoft.Office.Interop.Outlook.NameSpace mapiNamespace = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder CalendarFolder = null;

            oApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = oApp.GetNamespace("MAPI"); ;
            CalendarFolder = mapiNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);

            CalendarSharing cso = CalendarFolder.GetCalendarExporter();

            cso.CalendarDetail = OlCalendarDetail.olFullDetails;
            cso.IncludeWholeCalendar = true;
            cso.IncludeAttachments = false;
            cso.IncludePrivateDetails = true;
            cso.RestrictToWorkingHours = false;

            // save to file
            cso.SaveAsICal(calendarFileName);

        }

    }

}
