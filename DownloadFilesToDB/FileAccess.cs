using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;
using System.ComponentModel;
using System.Configuration;

namespace DownloadFilesToDB
{
    internal class FileAccess
    {
        private static string pathToSave;// = config.AppSettings.Settings["path"].Value;
        public FileAccess()
        {
            System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            pathToSave = config.AppSettings.Settings["path"].Value;
        }
        public static string GetDirectoryListingRegexForUrl(string url, string year)
        {

            return @"FileName=\/YieldX\/Derivatives\/Docs_DMTM\/"+year+@"\w*_D_Daily MTM Report.xls";

        }
        /// <summary>
        /// Gets a url in format: https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM/
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static void DoDownload(String url, string year)
        {
            //int counter = 0;
            List<string> files = new List<string>();
            string siteContent;// To keep html content
            // Here we're creating our request
            // we're simply building our HTTP request to send off to a url...
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.AutomaticDecompression = DecompressionMethods.GZip;

            //Buiding http request:

            // Wrap everything that can be disposed in using blocks... 
            // They dispose of objects and prevent them from lying around in memory...
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())  // Go query 
            using (Stream responseStream = response.GetResponseStream())               // Load the response stream
            using (StreamReader streamReader = new StreamReader(responseStream))       // Load the stream reader to read the response
            {
                siteContent = streamReader.ReadToEnd(); // Read the entire response and store it in the siteContent variable
            }
            Regex regex = new Regex(GetDirectoryListingRegexForUrl(url,year));
            MatchCollection matches = regex.Matches(siteContent);

            if (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        Console.WriteLine("downloading: " + match.ToString());//to see which file is downloaded

                        try
                        {
                            string urli = "https://clientportal.jse.co.za/_layouts/15/DownloadHandler.ashx?" + match.Value;
                            Uri singleUri = new Uri(urli);
                            //Webclient to download file
                            WebClient client = new WebClient();
                            client = new WebClient();
                            // Hookup DownloadFileCompleted Event
                            client.DownloadFileCompleted += new AsyncCompletedEventHandler(client_DownloadFileCompleted);

                            // Start the download and copy the file to download path
                            int pos = match.Value.LastIndexOf("/") + 1;
                            string fileName = match.Value.Substring(pos, match.Value.Length - pos);
                            string fullFilePath = pathToSave + fileName;
                            if (!File.Exists(fullFilePath))
                            {
                                client.DownloadFileAsync(singleUri, fullFilePath);
                                DataAccess dataHelper = new DataAccess();
                                files.Add(fullFilePath);

                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Exception: " + ex.Message);
                        }


                    }

                }
                Task.WaitAll();//make sure all downloads have finished

                foreach (string file in files)
                {
                    DataAccess helper = new DataAccess();
                    //string csvFile = file.Replace(".xls", ".csv");

                    helper.ReadAndStoreFiles(file);
                }

                Console.ReadLine();

            }
        }
        static void client_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            Console.WriteLine("File downloaded");
        }
    }
}
