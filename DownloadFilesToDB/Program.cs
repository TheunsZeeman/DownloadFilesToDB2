using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace DownloadFilesToDB
{
    class Program
    {
        public static void  Main(string[] args)
        {
            string path = "https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM";
            FileAccess fileAccess = new FileAccess();
            FileAccess.DoDownload(path,"2023");// fileAccess = new FileAccess();
            
        }
    }
}
