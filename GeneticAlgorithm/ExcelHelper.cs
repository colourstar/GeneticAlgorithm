using Excel;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelGroupCalculater
{
    public class ExcelHelper
    {
        public static List<DataRowCollection> ReadExcel(string filePath, string strPrefix, ref List<string> strTableName)
        {
            strTableName.Clear();
            List<DataRowCollection> arrCollections = new List<DataRowCollection>();
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            for (int i = 0; i < result.Tables.Count; ++i)
            {
                DataTable table = result.Tables[i];
                if (table == null || table.TableName.StartsWith(strPrefix) == false)
                {
                    continue;
                }

                strTableName.Add(table.TableName);
                arrCollections.Add(table.Rows);
            }

            return arrCollections;
        }

        public static DataRowCollection ReadExcel(string filePath, string strTableName)
        {
            List<DataRowCollection> arrCollections = new List<DataRowCollection>();
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();

            for (int i = 0; i < result.Tables.Count; ++i)
            {
                DataTable table = result.Tables[i];
                if (table == null || table.TableName != strTableName)
                {
                    continue;
                }
                return table.Rows;
            }

            return null;
        }

        public static System.Diagnostics.Process CreateShellExProcess(string cmd, string args, string workingDir = "", bool bWaitExit = true)
        {
            var pStartInfo = new System.Diagnostics.ProcessStartInfo(cmd);
            pStartInfo.Arguments = args;
            pStartInfo.CreateNoWindow = false;
            pStartInfo.UseShellExecute = true;
            pStartInfo.RedirectStandardError = false;
            pStartInfo.RedirectStandardInput = false;
            pStartInfo.RedirectStandardOutput = false;
            if (!string.IsNullOrEmpty(workingDir))
                pStartInfo.WorkingDirectory = workingDir;

            System.Diagnostics.Process process = System.Diagnostics.Process.Start(pStartInfo);
            if (bWaitExit)
            {
                process.WaitForExit();
            }
            return process;
        }

        public static void RunBat(string batfile, string args, string workingDir = "", bool bWaitExit = true)
        {
            var p = CreateShellExProcess(batfile, args, workingDir, bWaitExit);
            p.Close();
        }
    }
}
