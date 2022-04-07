using Xunit;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using AutomatedTest.Utilities;
using System.Net;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class A_QueryURL
    {
        [Fact]
        public void 分類URL並匯出到excel()
        {
            string excelExportPath = $@"{UserDataList.Upperfolderpath}xmlURLList";

            #region 撈網頁原始碼 > 存入string "sitemaps" > 字串做分割 > 存入 list<>totalURLList
            HttpWebRequest request1 = (HttpWebRequest)WebRequest.Create("https://www.esunbank.com.tw/bank/sitemap.xml");
            request1.Method = "GET";
            request1.AllowAutoRedirect = true;
            HttpWebResponse response1 = (HttpWebResponse)request1.GetResponse();

            StreamReader reader = new StreamReader(response1.GetResponseStream());
            string sitemaps = reader.ReadToEnd(); //爬出網頁原始碼並存入string "sitemaps"
            List<string> totalURLList = sitemaps.Split('\n').ToList();
            #endregion

            List<string> urlList = new List<string>();
            string pattern = @"((http|ftp|https)://)(([a-zA-Z0-9\._-]+\.[a-zA-Z]{2,6})|([0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}))(:[0-9]{1,4})*(/[a-zA-Z0-9\&%_\./-~-]*)?"; // 判斷網址URL
            foreach (var urlLink in totalURLList)
            {
                MatchCollection matchSources = Regex.Matches(urlLink, pattern); // 找出所有URL字串
                for (int i = 0; i < matchSources.Count; i++)
                {
                    if (matchSources[i].ToString().Contains("esunbank")) // 判斷字串是否包含esunbank
                    {
                        urlList.Add(matchSources[i].ToString().Replace("</loc>", "")); // URL寫入urlList
                    }
                }
            }
            Excel.Application excel_App1 = new Excel.Application(); //  new 一個excel應用程序
            Excel.Workbook excel_WB1 = excel_App1.Workbooks.Add(); // 新增excel檔案
            Excel.Worksheet excel_WS1 = new Excel.Worksheet(); //  新增工作表

            string[] structureList = new string[] { "/bank/personal", "/bank/personal/deposit", "/bank/personal/loan", "/bank/personal/credit-card", "/bank/personal/wealth",
                "/bank/personal/trust", "/bank/personal/insurance", "/bank/personal/lifefin", "/bank/personal/apply", "/bank/personal/event",
                "/bank/small-business", "/bank/corporate", "/bank/digital", "/bank/about", "/bank/marketing",
                "/bank/iframe/widget", "/bank/error", "/bank/bank-en", "/bank/preview" };

            excel_WB1.Worksheets.Add(Missing.Value, Missing.Value, structureList.Length - 1, Missing.Value); // 新增工作表數量 (total數量 = structureList個數 -1 + 預設1)

            int sheetNum = 1; //  第sheetNum個工作表
            foreach (var item in structureList)
            {
                excel_WS1 = (Excel.Worksheet)excel_WB1.Worksheets[sheetNum]; // 指到第sheetNum個工作表
                excel_WS1.Name = item.Replace('/', '_').Substring(6); //命名工作表
                excel_WS1.Activate(); // setup工作表焦點

                int index = 0; // excel 開始行數
                for (int i = 0; i < urlList.Count; i++)
                {
                    if (item == "/bank/personal" && urlList[i].EndsWith("bank/personal")) // for工作表"personal"內容
                    {
                        excel_App1.Cells[index + 1, 1] = urlList[i]; // 寫入excel
                        index++;
                        break;
                    }
                    if (urlList[i].Contains(item)) // 判斷URL是否含structureList內字串
                    {
                        excel_App1.Cells[index + 1, 1] = urlList[i]; // 寫入excel
                        index++;
                    }
                }
                sheetNum++;
            }
            ((Excel.Worksheet)excel_WB1.Worksheets[1]).Activate(); // 設定工作表focus on 工作表1

            excel_App1.DisplayAlerts = false; // 存檔時不跳是否覆蓋alert
            excel_WB1.SaveAs(excelExportPath); // 存檔

            //關閉及釋放文件
            excel_WB1.Close();
            excel_App1.Quit();

        }
    }
}
