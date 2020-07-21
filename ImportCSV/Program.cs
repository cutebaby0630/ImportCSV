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

namespace ImportCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            //Step1. 備份檔案
            //Step1.1. 本機檔案讀取
            var localfile = "PATMPatientItem";
            DirectoryInfo readlocalfile = new DirectoryInfo($@"C:\Users\user\Downloads\{localfile}.csv");
            //Step1.2. 檔案讀取→\\10.1.225.17\d$\csv  \\10.1.225.17\d$\CSV - 複製
            var host = @"10.1.225.17";
            var RDPfile = "CSV_20200721";
            var username = @"LAPTOP-ODUSIH5U\Administrator";
            var password = "p@ssw0rd";
            using (new RDPCredentials(host, username, password))
            { 
                //Step1.3. 找到相對應File
                DirectoryInfo readfile = new DirectoryInfo($@"\\{host}\d$\{RDPfile}\{localfile}.csv");
                //Step1.4. 將File中的資料存入var
                string LastWriteTime = File.GetLastWriteTime(readfile.ToString()).ToString("yyyyMMdd");
                //Step1.5. 修改名稱(原File_修改日期yyyyMMdd)--備份
                var FileName = Path.GetFileName(readfile.ToString().Replace(".csv",""));
                readfile.MoveTo($@"\\{host}\d$\{RDPfile}\" + FileName + "_" + LastWriteTime + ".csv");
                //Step1.6. 將下載的File 複製到 mstv
                File.Copy(readlocalfile.ToString(), $@"\\{host}\d$\{RDPfile}\" + FileName +".csv");
            }


            //Step2. SSMS import CSV
            //Step2.1. 連線SSMS
            //Step2.2. 將檔案名稱丟入SQL
            //Step2.3. 執行後的table存入List_dbnew
            //Step3. 核對 SQL 跟 檔案中筆數及ID是否正確
            //Step3.1. 讀取下載File將內容存入List_filenew
            //Step3.2. Map ListA & ListB 是否相同
            //Step3.3. 回傳比對結果
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
