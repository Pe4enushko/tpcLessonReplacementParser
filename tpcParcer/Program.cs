using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace tpcParcer
{
    class Program
    {
        static string Group = "Вп-21";
        static string Address = "http://www.tpcol.ru/studentam/79-raspisanie-i-zameny";
        static HttpClient Client;
        static void Main(string[] args)
        {
            Thread loaderThread = new Thread(() => 
            {
                Console.Write("Загрузка, не тыкайте кнопки пожалуйста...");

                while (true)
                {
                    Console.Write("\\");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                    Console.Write("|");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                    Console.Write("/");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                    Console.Write("-");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                    Console.Write("\\");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                    Console.Write("|");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                    Console.Write("/");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                    Console.Write("-");
                    Task.Delay(100).Wait();
                    Console.Write("\b");
                }
                
            });
            try
            {
                loaderThread.Start();
                var path = DownloadExcel().Result;
                var table = OpenExcel(path);
                DataRow[] rows = table.Select($"Группа = '{Group}'");
                loaderThread.Abort();
                Console.Clear();
                Console.WriteLine("Made by Pe4enushko");
                Console.WriteLine("Замены для группы: " + Group + " на завтра.");
                Console.WriteLine();
                foreach (DataColumn col in table.Columns)
                {
                    Console.Write(col.ColumnName + " | ");
                }
                Console.WriteLine();
                Console.WriteLine();
                if (rows.Length > 0)
                {
                    foreach (DataRow row in rows)
                    {
                        foreach (var item in row.ItemArray)
                        {
                            Console.Write(item + "\t");
                        }
                        Console.WriteLine();
                    }
                }
                else
                {
                    Console.WriteLine("Замен на завтра нет.");
                }
                Console.Read();
            }
            catch(AggregateException)
            {
                loaderThread.Abort();
                Console.Clear();
                Console.WriteLine("Подключение к интернету смэрт, либо файл с заменами на завтра не выложен");
                Console.Read();
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns>Path to downloaded file</returns>
        static async Task<string> DownloadExcel()
        {
            Client = new HttpClient();
            var doc = default(IHtmlDocument);
            using (var stream = await Client.GetStreamAsync(new Uri(Address)))
            {
                var parser = new HtmlParser();
                doc = await parser.ParseDocumentAsync(stream);
            }
            string datestr = DateTime.Now.AddDays(1).ToString("dd.MM.yyyy");
            var all_a = doc.GetElementsByTagName("a");
            var needed_a = all_a.First(a => (a.GetAttribute("href")?.Contains(datestr) ?? false));
            var exclDownloadURL = "http://www.tpcol.ru" + needed_a.GetAttribute("href");
            WebClient webcl = new WebClient();
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "tpcDocs");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            //path = Path.Combine(path, DateTime.Now.AddDays(1).ToString("dd-MM-yyyy.xlxs"));
            
            webcl.DownloadFile(exclDownloadURL, Path.Combine(path, "ZameniNaZavtra.xlsx"));
            
            
            return Path.Combine(path, "ZameniNaZavtra.xlsx");
        }
        static DataTable OpenExcel(string path)
        {
            Excel.Application app = new Excel.Application();
            try
            {
                Excel.Workbook book = app.Workbooks.Add(path);
                Excel.Worksheet sheet = book.Sheets[1];
                Excel.Range range = sheet.UsedRange;
                DataTable DT = new DataTable();
                for (int col = 2; col <= 6; col++)
                {
                    dynamic temp = range.Cells[2, col].Value;
                    DT.Columns.Add((string)(range.Cells[2,col].Value));
                }
                for (int row = 3; row <= range.Rows.Count; row++)
                {
                    List<string> cols = new List<string>(); 
                    for (int col = 2; col <= 6; col++)
                    {
                        cols.Add(Convert.ToString(range.Cells[row,col].Value));
                    }
                    DT.Rows.Add(cols.ToArray());
                    cols.Clear();
                }
                app.Quit();
                return DT;
            }
            catch(Exception exc)
            {
                app.Quit();
                throw exc;
            }
        }
    }
}
