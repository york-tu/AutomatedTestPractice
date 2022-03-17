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
    public class 檢查URL網頁
    {
        [Fact]

        //    "/bank/personal", "/bank/personal/deposit", "/bank/personal/loan", "/bank/personal/credit-card", "/bank/personal/wealth", 
        //    "/bank/personal/trust", "/bank/personal/insurance", "/bank/personal/lifefin", "/bank/personal/apply", "/bank/personal/event",
        //    "/bank/small-business", "/bank/corporate", "/bank/digital", "/bank/about", "/bank/marketing",
        //    "/bank/iframe/widget", "/bank/error", "/bank/bank-en", "/bank/preview" };

        //[Theory]
        //[InlineData (1,"")]
        //[InlineData(2, "")]
        //[InlineData(3, "")]
        // public void 讀取Excel內URL並檢查開啟的網頁(int caseNum, string expectResult)

        public void 讀取Excel內URL並檢查開啟的網頁()
        {




            Excel.Application excel_App = new Excel.Application(); //  new 一個excel應用程序
            Excel.Workbook excel_WB = excel_App.Workbooks.Open(UserDataList.sitemapsExcelPath); // pass指定路徑excel內容to "excel_WB"
            Excel.Worksheet excel_WS= (Excel.Worksheet)excel_WB.Worksheets[2]; //  讀特定sheet
            Excel.Range range = excel_WS.UsedRange; // 撈出sheet內容, pass to "range"
            for (int i = 0; i < range.Count; i++)
            {
                var aaa = ((Excel.Range) range.Cells[i + 1, 1]).Value; // 讀出第 i 行URL
            }


         
            int sheetNum = 1; //  第sheetNum個工作表



            excel_App.DisplayAlerts = false; // 存檔時不跳是否覆蓋alert
 
            //關閉及釋放文件
            excel_WB.Close();
            excel_App.Quit();

        }
    }
}
