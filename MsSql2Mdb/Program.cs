using MsSql2Mdb.Data;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MsSql2Mdb
{
    class Program
    {
        static Logger logger;

        static string Help
        {
            get
            {
                return @"環科 MsSql2MDB 匯入公用程式 v1.0
參數用法：
  -f: 指定要產生的 Access Mdb 檔案名稱.
	例：
	MsSql2Mdb.exe -f:D:\Gelis Documents\MentorTrust\Mdb\FuJenClassRoom.mdb


注意：如果 -f: 參數內沒有指定完整路徑，會將 MDB 檔案產生在目前的執行目錄內.

Copyright © MentorTrust 2015";
            }
        }
        /// <summary>
        /// 顯示錯誤訊息
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="Message"></param>
        public static void ShowError(string Message)
        {
            logger.Error(Message);
            Console.WriteLine(Message);
        }

        static void Main(string[] args)
        {
            logger = LogManager.GetCurrentClassLogger();

            string AccessMdb = string.Empty; //來自 -f: 參數的 mdb 檔案
            string ExecPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string MdbFile = string.Empty;

            string[] Arguments = ParseArgs(args);

            if(Arguments.Length>0)
            {
                AccessMdb = Arguments[0].Substring(3, Arguments[0].Length - 3);
                if(Path.GetExtension(AccessMdb).ToUpper()!=".MDB")
                {
                    ShowError("提供的 -f:[Access 資料庫] 檔案的副檔名須為 .mdb");
                    return;
                }
                if (Path.GetDirectoryName(AccessMdb) != ExecPath && Path.GetDirectoryName(AccessMdb)!="")
                {
                    MdbFile = AccessMdb;
                }
                else
                {
                    MdbFile = Path.Combine(ExecPath, AccessMdb);
                }
                //建立 Access 資料庫 MDB 檔案
                MdbHelper.CreateNewAccessDatabase(AccessMdb);
                //將資料轉入 Access Mdb 中.
                MdbHelper.TransferData2NewAccessDatabase(AccessMdb);

                ShowError(string.Format("資料庫 '{0}' 轉換完成！", AccessMdb));
                //Console.WriteLine(string.Format("資料庫 '{0}' 轉換完成！", AccessMdb));
                //Console.WriteLine(MdbFile);
            }
            else
            {
                Console.WriteLine(Help);
            }
            //Console.ReadLine();
        }

        #region Parse 來源參數，防呆、組合路徑字串
        private static string[] ParseArgs(string[] args)
        {
            List<string> result = new List<string>();
            bool IsSource = false;

            string ff = string.Empty;
            string dd = string.Empty;
            string host = string.Empty;
            string Source = string.Empty;

            foreach (string ar in args)
            {
                if (ar.StartsWith("-H:"))
                {
                    //ff = ar;
                }
                else if (ar.StartsWith("-h:"))
                {
                    //dd = ar;
                }
                //else if(ar.StartsWith("-Host:"))
                //{
                //    host = ar;
                //}
                else if (ar.StartsWith("-f:"))
                {
                    IsSource = true;
                    Source += ar;
                }
                else
                {
                    if (IsSource)
                    {
                        Source += string.Format(" {0}", ar);
                    }
                }
            }

            if (!string.IsNullOrEmpty(ff))
                result.Add(ff);
            if (!string.IsNullOrEmpty(dd))
                result.Add(dd);
            if (!string.IsNullOrEmpty(host))
                result.Add(host);
            if (!string.IsNullOrEmpty(Source))
                result.Add(Source);

            return result.AsEnumerable().ToArray();
        }
        #endregion

    }
}
