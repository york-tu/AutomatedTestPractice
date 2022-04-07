using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using AutomatedTest.Utilities;
using System.Net;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class C_c_CheckURLResponse_Total_excel:IntegrationTestBase
    {
        public C_c_CheckURLResponse_Total_excel(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        #region 工作表對應網址分類
        /*
        工作表 1:      "/bank/personal",
        工作表 2:      "/bank/personal/deposit",
        工作表 3:      "/bank/personal/loan", 
        工作表 4:      "/bank/personal/credit-card",
        工作表 5:      "/bank/personal/wealth", 
        工作表 6:      "/bank/personal/trust", 
        工作表 7:      "/bank/personal/insurance",
        工作表 8:      "/bank/personal/lifefin", 
        工作表 9:      "/bank/personal/apply", 
        工作表 10:     "/bank/personal/event",
        工作表 11:     "/bank/small-business", 
        工作表 12:     "/bank/corporate", 
        工作表 13:     "/bank/digital", 
        工作表 14:     "/bank/about", 
        工作表 15:     "/bank/marketing",
        工作表 16:     "/bank/iframe/widget", 
        工作表 17:     "/bank/error", 
        工作表 18:     "/bank/bank-en",
        工作表 19:     "/bank/preview";
        */
        #endregion
        [Theory]
        [InlineData (1,19)] // 訪問工作表1-19頁內的URL
         public void 檢查URL是否可訪問(int startSheetIndex, int endSheetIndex)
        {
            CreateReport("URLResponseReport", "York");

            Excel.Application excel_App = new Excel.Application(); //  new 一個excel應用程序
            Excel.Workbook excel_WB = excel_App.Workbooks.Open(UserDataList.sitemapsExcelPath); // pass指定路徑excel內容to "excel_WB"

            for (int index = startSheetIndex; index <= endSheetIndex; index++)
            {
                Excel.Worksheet excel_WS = (Excel.Worksheet)excel_WB.Worksheets[index]; //  讀指定sheet
                INFO(excel_WS.Name);

                Excel.Range range = excel_WS.UsedRange; // 撈出sheet內容, pass to "range"
                for (int i = 0; i < range.Count; i++)
                {
                    string strURL = (string)((Excel.Range)range.Cells[i + 1, 1]).Value; // 讀出第 i 行URL
                    try
                    {
                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL); //打https request 到server
                        request.AllowAutoRedirect = false; // 不予許redirect
                        request.Timeout = 30000; //超時時間10秒

                        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            PASS($"URL_{i + 1}: {strURL}    ==>    {response.StatusCode}");
                            response.Close();
                        }
                        else
                        {
                            FAIL($"URL_{i + 1}: {strURL}    ==>   {response.StatusCode}");
                            response.Close();
                        }
                    }
                    catch (WebException e)
                    {
                        WARNING($"URL_{i + 1}: {strURL}    ==>    [ Exception] {e.Message}");
                    }
                }
                INFO("");
            }

            //關閉及釋放文件
            excel_WB.Close();
            excel_App.Quit();

            CloseBrowser(); 
        }
    }
}
