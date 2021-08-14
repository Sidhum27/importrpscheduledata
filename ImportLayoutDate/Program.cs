using ImportScheduleData.Utilities;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace ImportLayoutDate
{
    public class Program
    {
        private static string TICECentralConStr = string.Empty;
        private static string TICELoggingConStr = string.Empty;
        private static string jobName = "ImportLayoutDate";
        private static DateTime? jobStarted = null;
        static void Main(string[] args)
        {
            try
            {
                jobStarted = DateTime.Now;
                GetConnection();
                ProcessExcelData();
            }
            catch (Exception ex)
            {
                //send error log on mail
                string toMails = "paul.broe@cware.ie;brian.culhane@cware.ie;kanhaiyaa.lal@techbitsolution.com";
                SaveLogs("Failed", 0);
                EmailSender.SendEmail(toMails, "Failed job...Import layout date into ORegonLayout table", ex.Message);
            }
        }

        private static void GetConnection()
        {
            TICECentralConStr = ConfigurationManager.ConnectionStrings["TICECentralServices"].ConnectionString;
            TICELoggingConStr = ConfigurationManager.ConnectionStrings["TICELogging"].ConnectionString;
        }

        private static void ProcessExcelData()
        {
            try
            {
                string fileFolder = ConfigurationManager.AppSettings["folderPath"];
                if (Directory.Exists(fileFolder))
                {
                    string[] file = Directory.GetFiles(fileFolder, "Desk Reviews and Tools RTD.xlsm");
                    if (file.Length > 0)
                    {
                        Console.WriteLine($"File Path : {file[0]}");
                        Console.WriteLine("-------------------------");
                        if (File.Exists(file[0]))
                        {
                            var fileName = Path.GetFileName(file[0]);
                            Console.WriteLine($"Processing File... {fileName}");
                            if (!string.IsNullOrEmpty(file[0]) && !string.IsNullOrEmpty(TICECentralConStr) && !string.IsNullOrEmpty(TICELoggingConStr))
                            {
                                DataTable dsXls = GetDataFromExcelFile(file[0]);
                                int totalRecords = UpdateTools(dsXls);
                                SaveLogs("Completed", totalRecords);
                                Console.WriteLine($"'{file[0]}' has been processed successfully");
                            }
                        }
                    }
                    else
                        throw new Exception($"Could'nt find the excel file(s) on this path: '{fileFolder}'");
                }
                else
                    throw new Exception("Could'nt find the folder location, please check your app configuration file.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void SaveLogs(string status, int totalRecords = 0)
        {
            string query = "INSERT INTO [dbo].[JobStatus]([JobName],[Status],[CreatedDate],[UpdatedDate]) VALUES(@jobName, @status, @createdDate, @UpdatedDate);" +
                "INSERT INTO [dbo].[JobLogged]([JobName],[recordCount],[dateLogged]) VALUES(@jobName, @recordCount, @UpdatedDate)";
            using (SqlConnection con = new SqlConnection(TICELoggingConStr))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    try
                    {
                        cmd.CommandTimeout = 10000;
                        cmd.Connection = con;
                        cmd.CommandText = query;
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@jobName", jobName);
                        cmd.Parameters.AddWithValue("@status", status);
                        cmd.Parameters.AddWithValue("@createdDate", jobStarted);
                        cmd.Parameters.AddWithValue("@UpdatedDate", DateTime.Now);
                        cmd.Parameters.AddWithValue("@recordCount", totalRecords);
                        con.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }
        }

        private static DataTable GetDataFromExcelFile(string file)
        {
            var duplicateIndex = new List<int>();
            DataTable xlsData;
            string sheetName = "Desk Reviews & Tools RTD";
            string[] selectedCol = new string[] { "Entity Code Life", "Layout Target Release Date", "At Risk", "Comments for TO to bring to design" };
            try
            {
                ExcelEngine excelEngine = new ExcelEngine();
                using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
                {
                    var workbook = excelEngine.Excel.Workbooks.Open(stream);
                    var workSheet = workbook.Worksheets[sheetName];
                    var tempDT = workSheet.ExportDataTable(2, 1, workSheet.Rows.Count(), 20, ExcelExportDataTableOptions.ComputedFormulaValues);
                    for (int i = 0; i < tempDT.Columns.Count; i++)
                    {
                        if (tempDT.Rows[0][i].ToString() != "")
                            tempDT.Columns[i].ColumnName = tempDT.Rows[0][i].ToString();
                    }
                    tempDT.Rows[0].Delete();
                    xlsData = new DataView(tempDT).ToTable("LayoutData", true, selectedCol);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return xlsData;
        }

        private static int UpdateTools(DataTable dtXls)
        {
            using (SqlConnection con = new SqlConnection(TICECentralConStr))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    try
                    {
                        cmd.CommandTimeout = 10000;
                        cmd.Connection = con;
                        cmd.CommandText = "[dbo].[uspImportLayoutReport]";
                        cmd.CommandType = CommandType.StoredProcedure;
                        SqlParameter[] param = new SqlParameter[] {
                            new SqlParameter("@LayoutReport", SqlDbType.Structured),
                            new SqlParameter("@TotalRecordCount", SqlDbType.Int),
                            new SqlParameter("@ErrorLogID", SqlDbType.Int)
                        };
                        param[0].Value = dtXls;
                        param[1].Direction = ParameterDirection.Output;
                        param[2].Direction = ParameterDirection.Output;
                        cmd.Parameters.AddRange(param.ToArray());
                        con.Open();
                        cmd.ExecuteNonQuery();
                        int logID = Convert.ToInt32(cmd.Parameters["@ErrorLogID"].Value);
                        int totalRecordsAffected = Convert.ToInt32(cmd.Parameters["@TotalRecordCount"].Value);
                        if (logID > 0)
                            throw new Exception($"Data not successfully updated. DB Error LogID: {logID}");
                        return totalRecordsAffected;
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }
        }

    }
}
