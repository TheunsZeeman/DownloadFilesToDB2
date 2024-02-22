using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace DownloadFilesToDB
{
    class DataAccess
    {
        public async Task ReadAndStoreFiles(string file)
        {
            bool success = true;
            int ixl = file.LastIndexOf(@"\")+1;
            int wordlen = file.Length - file.LastIndexOf(@"\") ;
            string fileName = file.Substring(ixl, wordlen -1 );
            string fileDate = fileName.Substring(0,4) + '/' + fileName.Substring(4,2) + '/' + fileName.Substring(6, 2);
            string strike = "";
            string callput = "";
            string expiryDate = "";
            string contractDetails = "";
            string classification = "";
            string MTM_Yield = "";
            string markPrice = "";
            string spotPrice = "";
     
            string previousMTM = "";
            string previousPrice = "";
            string premiumOnOption = "";
            string volatility = "";
            string delta = "";
            string deltaValue = "";
            string contractsTraded = "";
            string openInterest = "";

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = 1048576;//maximum rows
            int colCount = 21;

            //Excel code
            #region excel
            Console.WriteLine("Processing files ");
            for (int i = 10; i <= rowCount; i++)
            {

                
                    for (int j = 1; j <= colCount; j++)
                    {
                        //progress..
                        if (j == 1)
                            Console.Write(".");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null) { }



                    }
                

                try
                {
                    if (xlRange.Cells[i, 1].Value2.ToString().Contains("Total for Derivatives")) break;
                    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 2] != null && xlRange.Cells[i, 3] != null && xlRange.Cells[i, 4] != null && xlRange.Cells[i, 5] != null && xlRange.Cells[i, 6] != null && xlRange.Cells[i, 7] != null && xlRange.Cells[i, 8] != null && xlRange.Cells[i, 9] != null && xlRange.Cells[i, 10] != null && xlRange.Cells[i, 11] != null && xlRange.Cells[i, 12] != null && xlRange.Cells[i, 13] != null && xlRange.Cells[i, 14] != null && xlRange.Cells[i, 15] != null && xlRange.Cells[i, 16] != null && xlRange.Cells[i, 17] != null )
                    {
                        if (!xlRange.Cells[i, 1].Value2.ToString().Contains("INTEREST RATE AND CURRENCY DERIVATIVES") && !xlRange.Cells[i, 1].Value2.ToString().Contains("DAILY MTM REPORT") && !xlRange.Cells[i, 1].Value2.ToString().Contains("DAILY SUMMARY FOR") && !xlRange.Cells[i, 1].Value2.ToString().Contains("CONTACT DETAILS"))
                        {
                            contractDetails = xlRange.Cells[i, 1].Value2.ToString();
                            expiryDate = xlRange.Cells[i, 3].Value2.ToString();
                            classification = xlRange.Cells[i, 4].Value2.ToString();
                            MTM_Yield = xlRange.Cells[i, 7].Value2.ToString();
                            markPrice = xlRange.Cells[i, 8].Value2.ToString();
                            spotPrice = xlRange.Cells[i, 9].Value2.ToString();
                            previousMTM = xlRange.Cells[i, 10].Value2.ToString();
                            previousPrice = xlRange.Cells[i, 11].Value2.ToString();


                            premiumOnOption = xlRange.Cells[i, 12].Value2.ToString();
                            volatility = xlRange.Cells[i, 13].Value2.ToString();
                            delta = xlRange.Cells[i, 14].Value2.ToString();
                            deltaValue = xlRange.Cells[i, 15].Value2.ToString();
                            contractsTraded = xlRange.Cells[i, 16].Value2.ToString();
                            openInterest = xlRange.Cells[i, 17].Value2.ToString();
                        }
                        else { break; }
                    }
                    else { break; }
                }
                catch(Exception ex)
                {
                    if (ex.Message.Contains("Cannot perform runtime binding on a null reference"))
                    {
                        //wrong format

                    }
                    else
                    {
                        Console.WriteLine("Error:" + ex.Message);
                    }
                    //break;
                }finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                #endregion
                // Database code
                #region database
                string connectionString = ConfigurationManager.ConnectionStrings["JSEConnection"].ConnectionString;
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand command = connection.CreateCommand())
                {
                    command.CommandText = $@"INSERT INTO DailyMTM (FileDate,  
                    Contract, 
                    ExpiryDate, 
                    Classification, 
                    Strike, 
                    CallPut,  
                    MTMYield,
                    MarkPrice,
                    SpotRate,
                    PreviousMTM,
                    PreviousPrice,
                    PremiumOnOption,
                    Volatility,
                    Delta,
                    DeltaValue,
                    ContractsTraded,
                    OpenInterest) 
                VALUES(@FileDate,  
                    @Contract, 
                    @ExpiryDate, 
                    @Classification, 
                    @Strike, 
                    @CallPut,  
                    @MTMYield,
                    @MarkPrice,
                    @SpotRate,
                    @PreviousMTM,
                    @PreviousPrice,
                    @PremiumOnOption,
                    @Volatility,
                    @Delta,
                    @DeltaValue,
                    @ContractsTraded,
                    @OpenInterest)";
                    try
                    {
                        command.Parameters.Add(new SqlParameter("@FileDate", System.Data.SqlDbType.Date)).Value = DateTime.ParseExact(fileDate, "yyyy/MM/dd", CultureInfo.InvariantCulture);



                        command.Parameters.Add(new SqlParameter("@Contract", System.Data.SqlDbType.NVarChar)).Value = contractDetails;
                        
                        command.Parameters.Add(new SqlParameter("@Classification", System.Data.SqlDbType.NVarChar)).Value = classification;
                        command.Parameters.Add(new SqlParameter("@Strike", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(strike) ? "0": strike);
                        command.Parameters.Add(new SqlParameter("@CallPut", System.Data.SqlDbType.NVarChar)).Value = callput;
                        command.Parameters.Add(new SqlParameter("@MTMYield", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(MTM_Yield) ? "0": MTM_Yield);
                        command.Parameters.Add(new SqlParameter("@MarkPrice", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(markPrice) ? "0": markPrice);
                        command.Parameters.Add(new SqlParameter("@SpotRate", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(spotPrice) ? "0": spotPrice);
                        command.Parameters.Add(new SqlParameter("@PreviousMTM", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(previousMTM) ? "0": previousMTM);
                        command.Parameters.Add(new SqlParameter("@PreviousPrice", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(previousPrice) ? "0": previousPrice);
                        command.Parameters.Add(new SqlParameter("@PremiumOnOption", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(premiumOnOption) ? "0": premiumOnOption);
                        volatility = volatility.Replace(".", ",");
                        command.Parameters.Add(new SqlParameter("@Volatility", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(volatility) ? "0": volatility);
                        command.Parameters.Add(new SqlParameter("@Delta", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(delta) ? "0": delta);
                        command.Parameters.Add(new SqlParameter("@DeltaValue", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(deltaValue) ? "0": deltaValue);
                        command.Parameters.Add(new SqlParameter("@ContractsTraded", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(contractsTraded) ? "0": contractsTraded);
                        command.Parameters.Add(new SqlParameter("@OpenInterest", System.Data.SqlDbType.Float)).Value = float.Parse(string.IsNullOrWhiteSpace(openInterest) ? "0": openInterest);
                        try
                        {
                            double ddate = double.Parse(expiryDate);
                               DateTime date = DateTime.FromOADate(ddate);
                            command.Parameters.Add(new SqlParameter("@ExpiryDate", System.Data.SqlDbType.Date)).Value = date;
                        }catch(Exception e)
                        {
                            string date = fileDate.Replace("2023", "2024");//incase date is incorrect
                            command.Parameters.Add(new SqlParameter("@ExpiryDate", System.Data.SqlDbType.Date)).Value = DateTime.ParseExact(date, "yyyy/MM/dd", CultureInfo.InvariantCulture);
                        }
                        }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }

                    //
                    try
                    {
                        connection.Open();
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                    catch (SqlException ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                    }

                }
                #endregion
            }

            Console.WriteLine("success: " + success);
        }
    }
}
