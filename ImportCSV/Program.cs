using System;
using System.IO;

namespace ImportCSV
{
    class Program
    {
        static void Main(string[] args)
        {
			//Step1. 備份檔案
			//Step1.1. 檔案讀取→\\10.1.225.17\d$\csv
			
			//Step1.2. 找到相對應File
			//Step1.3. 將File中的資料存入List_(日期)
			//Step1.3. 修改名稱(原File_修改日期yyyyMMdd)
			//Step1.4. 將下載的File 複製到 mstv
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
}
