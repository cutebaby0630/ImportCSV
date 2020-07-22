using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SqlServerHelper.Core;
using SqlServerHelper;
using System.Diagnostics;
using ServiceStack;
using Microsoft.VisualBasic;
using NPOI.SS.Formula.Functions;
using System.Reflection;
using System.Text;
using Microsoft.VisualBasic.FileIO;

namespace ImportCSV
{
    class Program
    {

        static void Main(string[] args)
        {
            //Step1. 備份檔案
            //Step1.1. 本機檔案讀取
            var filename = "PATMPatientItem";
            DirectoryInfo readlocalfile = new DirectoryInfo($@"C:\Users\user\Downloads\{filename}.csv");
            DataTable local_dt = TxtConvertToDataTable(readlocalfile.ToString(), "localfile", "|");
            string firstColumnName = local_dt.Columns[0].ColumnName;
            DataRow[] rows = local_dt.Select();

            // Print the value one column of each DataRow.
            for (int i = 0; i < rows.Length; i++)
            {
                Console.WriteLine(rows[i][firstColumnName]);
            }
            //Step1.2. 檔案讀取→\\10.1.225.17\d$\csv  \\10.1.225.17\d$\CSV - 複製
            var host = @"10.1.225.17";
            var RDPfile = "CSV_20200721";
            var username = @"LAPTOP-ODUSIH5U\Administrator";
            var password = "p@ssw0rd";
            using (new RDPCredentials(host, username, password))
            {
                //Step1.3. 找到相對應File
                DirectoryInfo readfile = new DirectoryInfo($@"\\{host}\d$\{RDPfile}\{filename}.csv");
                //Step1.4. 將File中的資料存入var
                string LastWriteTime = File.GetLastWriteTime(readfile.ToString()).ToString("yyyyMMdd");
                //Step1.5. 修改名稱(原File_修改日期yyyyMMdd)--備份
                //readfile.MoveTo($@"\\{host}\d$\{RDPfile}\" + filename + "_" + LastWriteTime + ".csv");
                readfile.MoveTo($@"\\{host}\d$\{RDPfile}\" + filename + "_" + "1.csv");

                //Step1.6. 將下載的File 複製到 mstv
                File.Copy(readlocalfile.ToString(), $@"\\{host}\d$\{RDPfile}\" + filename + ".csv");

            }
            //Step2. SSMS import CSV
            //Step2.1. 連線SSMS
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsetting.json", optional: true, reloadOnChange: true).Build();
            //取得連線字串
            string connString = config.GetConnectionString("DefaultConnection");
            SqlServerDBHelper sqlHelper = new SqlServerDBHelper(string.Format(connString, "HISDB", "msdba", "1qaz@wsx"));
            //Step2.2. 將檔案名稱丟入SQL
            string sqlCSV = $@"--use [HISDB];
                              --use[HISBILLINGDB];
                              --, CODEPAGE = 65001
                              --已更新至225.17
                             DECLARE @TABLENAME VARCHAR(MAX) = '{filename}'; --檔案名稱去除CSV

                             IF LEFT(@TABLENAME,3) IN('CHG', 'CLA')
                             BEGIN
                             use[HISBILLINGDB];
                             END
                             ELSE
                             BEGIN
                             use[HISDB];
                             END

                             EXEC('TRUNCATE TABLE ' + @TABLENAME)
                             DECLARE @INS_CNT INT,@UPD_CNT INT
                             DECLARE @START_TIME VARCHAR(24)
                             SET @START_TIME = CONVERT(VARCHAR(24), GETDATE(), 121)
                             DECLARE @ERR_NO INT
                             DECLARE @SP_NAME VARCHAR(100) = ('mSP_INS_' + @TABLENAME + '_all');
                             --DECLARE @SP_NAME VARCHAR(100) = ('mSP_INS_' + @TABLENAME + '_fromExternal');
                             EXEC @ERR_NO = @SP_NAME @INS_CNT OUTPUT, @UPD_CNT OUTPUT
                             PRINT @ERR_NO;
                             IF @ERR_NO = 0 BEGIN
                             EXEC('SELECT * FROM ' + @TABLENAME);
                             END;";
            DataTable impotCSV_dt = sqlHelper.FillTableAsync(sqlCSV).Result;
            int rowCount = (impotCSV_dt == null) ? 0 : impotCSV_dt.Rows.Count;
            Console.WriteLine(rowCount);
            DataRow[] id_row = impotCSV_dt.Select();

            //Step3. 核對 SQL 跟 檔案中筆數及ID是否正確
            bool status = true;
            //Step3.1. Map ListA & ListB 是否相同
            for (int i = 56; i < rows.Length; i++)
            {
                bool intStatus = Int32.TryParse(rows[i][firstColumnName].ToString(), out int num);
                if (intStatus)
                {
                    for (int c = 56; c < id_row.Length; c++)
                    {
                        if (!Equals(num, id_row[c][firstColumnName]))
                        {
                            status = false;
                        }
                        else
                        {
                            status = true;
                            break;
                        }

                    }
                }
            }
            //Step3.3. 回傳比對結果
            Console.WriteLine(status);

            
            
            //Step3.3.1 如果失敗必須先將備份的File名稱rename
            //Step3.3.2 重新匯入225.17

            //Step4. 比對新跟舊的差異發送Email
            //Step4.1. 將List_(日期)和List_filenew比對差異
            //Step4.2. 差異利用List_sync儲存
            //Step4.3. 將List_sync利用Email寄發

            //Step5. 同步到各個DB
            //Step5.1 讀取相對應SyncData
            //Step5.2 執行同步到各個DB
        }

        //CSV to datatable
        public static DataTable TxtConvertToDataTable(string File, string TableName, string delimiter)
        {
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            StreamReader s = new StreamReader(File, System.Text.Encoding.Default);
            //string ss = s.ReadLine();//skip the first line
            string[] columns = s.ReadLine().Split(delimiter.ToCharArray());
            ds.Tables.Add(TableName);
            foreach (string col in columns)
            {
                bool added = false;
                string next = "";
                int i = 0;
                while (!added)
                {
                    string columnname = col + next;
                    columnname = columnname.Replace("#", "");
                    columnname = columnname.Replace("'", "");
                    columnname = columnname.Replace("&", "");

                    if (!ds.Tables[TableName].Columns.Contains(columnname))
                    {
                        ds.Tables[TableName].Columns.Add(columnname.ToUpper());
                        added = true;
                    }
                    else
                    {
                        i++;
                        next = "_" + i.ToString();
                    }
                }
            }

            string AllData = s.ReadToEnd();
            string[] rows = AllData.Split("\n".ToCharArray());

            foreach (string r in rows)
            {
                string[] items = r.Split(delimiter.ToCharArray());
                ds.Tables[TableName].Rows.Add(items);
            }

            s.Close();

            dt = ds.Tables[0];

            return dt;
        }


    }

    #region -- connect RDP --
    class RDPCredentials : IDisposable
    {
        private string Host { get; }

        public RDPCredentials(string Host, string UserName, string Password)
        {
            var cmdkey = new Process
            {
                StartInfo =
            {
                FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe"),
                Arguments = $@"/list",
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false,
                RedirectStandardOutput = true
            }
            };
            cmdkey.Start();
            cmdkey.WaitForExit();
            if (!cmdkey.StandardOutput.ReadToEnd().Contains($@"TERMSRV/{Host}"))
            {
                this.Host = Host;
                cmdkey = new Process
                {
                    StartInfo =
            {
                FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe"),
                Arguments = $@"/generic:DOMAIN/{Host} /user:{UserName} /pass:{Password}",
                WindowStyle = ProcessWindowStyle.Hidden
            }
                };
                cmdkey.Start();
            }
        }

        public void Dispose()
        {
            if (Host != null)
            {
                var cmdkey = new Process
                {
                    StartInfo =
            {
                FileName = Environment.ExpandEnvironmentVariables(@"%SystemRoot%\system32\cmdkey.exe"),
                Arguments = $@"/delete:TERMSRV/{Host}",
                WindowStyle = ProcessWindowStyle.Hidden
            }
                };
                cmdkey.Start();
            }
        }
    }
    #endregion



}
