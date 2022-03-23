using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using AutomatedTest.Utilities;
using System.Net;
using Xunit.Abstractions;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Linq;
using System.Globalization;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class II_I_檢查網頁內容:IntegrationTestBase
    {
        public II_I_檢查網頁內容(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Theory]
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
        [InlineData(new int[] { 2 })]
        public void 檢查網頁內容是否符合預期(int[] sheet)
        {
            CreateReport("網頁內容檢查", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");
            Excel.Application excel_App = new Excel.Application(); //  新增excel應用程序

            #region step 1: 讀取sitemap excel內容到 excel_WB
            Excel.Workbook excel_WB = excel_App.Workbooks.Open(UserDataList.sitemapsExcelPath); // pass指定路徑excel全部內容to "excel_WB"
            #endregion

            #region step 2: 讀取 expect result excel (expectResult.xlsx)第一行內容 (excel_Expect_xxxx) 到 expectResultArray 陣列中
            Excel.Workbook excel_Expect_WB = excel_App.Workbooks.Open($"{UserDataList.Upperfolderpath}testdata\\ExpectResult_0323.xlsx"); // open 指定路徑excel
            Excel.Worksheet excel_Expect_WS = (Excel.Worksheet)excel_Expect_WB.Worksheets[1]; // 指定讀取excel 檔第一個工作表
            int sheetRows = 4; // 工作表內行數 
            Excel.Range expectResultRange = (Excel.Range)excel_Expect_WS.UsedRange; // export excel 內容 to Range
            string[] expectResultArray = new string[expectResultRange.Count / sheetRows];
            for (int i = 0; i < expectResultRange.Count / sheetRows; i++)
            {
                expectResultArray[i] = (string)((Excel.Range)expectResultRange.Cells[i + 1, 1]).Value; // 將excel第一行內容丟進expectResultArray陣列中
            }
            #endregion

            #region  browser 不開啟網頁設定 
            //Chrome headless 參數設定
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("--headless");
            chromeOptions.AddArguments("--disable-gpu");
            chromeOptions.AddArguments("--incognito");
            chromeOptions.AddArguments("--window-size=1920x1080");
            chromeOptions.AddArguments("--ignore-certificate-errors");
            chromeOptions.AddArguments("--allow-running-insecure-content");
            chromeOptions.AddArguments("--disable-extensions");
            chromeOptions.AddArguments("--proxy-server='direct://'");
            chromeOptions.AddArguments("--proxy-bypass-list=*");
            chromeOptions.AddArguments("--start-maximized");
            chromeOptions.AddArguments("--disable-dev-shm-usage");
            chromeOptions.AddArguments("--no-sandbox");
            chromeOptions.AddArguments("--disable-blink-features=AutomationControlled");
            chromeOptions.AddArguments("--disable-infobars");
            //建置 Chrome Driver
            var driver = new ChromeDriver(chromeOptions);

            //var firefoxOptions = new FirefoxOptions();
            ////firefoxOptions.AddArguments("--headless");
            //var driver = new FirefoxDriver(firefoxOptions);
            #endregion


            foreach (var index in sheet)
            {
                Excel.Worksheet excel_WS = (Excel.Worksheet)excel_WB.Worksheets[index]; //  讀sitemaps excel指定工作表內容
                Excel.Range range = excel_WS.UsedRange; // 撈出工作表內容, pass to "range"
                INFO(excel_WS.Name);

                for (int i = 0; i < range.Count; i++)
                {
                    string strURL = (string)((Excel.Range)range.Cells[i + 1, 1]).Value; // 讀出第 i 行URL
                    int meetExpectResultURLIndex = Array.IndexOf(expectResultArray, strURL); // 搜尋目標URL對應到expect result 陣列中的位置

                    if (meetExpectResultURLIndex == -1) // 該URL在對應表上搜尋不到
                    {
                        WARNING($"[新增/變動] URL:  {strURL}");
                        continue;
                    }
                    else
                    {
                        string cssSelector = (string)((Excel.Range)expectResultRange.Cells[meetExpectResultURLIndex + 1, 2]).Value; // 讀 expect result excel中第二行欄位值
                        string expectString = (string)((Excel.Range)expectResultRange.Cells[meetExpectResultURLIndex + 1, 3]).Value;  // 讀 expect result excel中第三行欄位值

                        ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // 另開新頁
                        driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on 新頁上 
                        driver.Navigate().GoToUrl(strURL);
                        #region 當網頁自動切到M版時, click "切換電腦版" 強制切回PC版
                        if (driver.Url.ToString().Contains("?dev=mobile")) // workaround: 當網頁自動切到M版時, 強制切回PC版
                        {
                            TestBase.ScrollPageUpOrDown(driver, 1500);
                            driver.FindElementByClassName("changeTarget").Click();
                        }
                        #endregion
                        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(600); //600秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                        if (driver.Url.ToString() != strURL) // 檢查網頁開啟當下網址是否為輸入網址 (判斷網頁是否有redirect)
                        {
                            WARNING($"[Page Redirect], {driver.Url.ToString()}  (Expect: {strURL})");
                            WARNING(TestBase.PageSnapshotToReport(driver));
                            continue;
                        }
                        else if (strURL == "https://easyfee.esunbank.com.tw/index.action")
                        {
                            WARNING($"{strURL}, Keyword: {expectString}");
                            WARNING(TestBase.PageSnapshotToReport(driver));
                            continue;
                        }
                        else if (driver.Url.ToString() == "https://www.esunbank.com.tw/bank/personal/event/calendar/events") // Workaround 1 : 活動日曆URL >>> 需抓當天日期
                        {
                            string day_of_week = DateTime.Now.ToString("dddd", new CultureInfo("en-us")); // 英文星期幾(e.g., Wednesday)
                            string day = DateTime.Now.ToString("dd"); // 日期 (e.g., 23)
                            string month = DateTime.Now.ToString("MMMM", new CultureInfo("zh-cn")); // 中文月份 (e.g., 三月)
                            expectString = $"{day}\r\n{day_of_week}\r\n{month}";
                        }
                        else if (driver.Url.ToString().EndsWith("/pages")) // Workaround 2 : 網址為 ".../page" >>> 網頁為空內容
                        {
                            expectString = "";
                        }
                        try
                        {
                            Assert.Contains(expectString, driver.FindElement(By.CssSelector(cssSelector)).Text); // 判斷element 字串是否符合預期
                            PASS($"{strURL}, Keyword: {expectString}");
                            PASS(TestBase.PageSnapshotToReport(driver));
                        }
                        catch (Exception e)
                        {
                            FAIL($"{strURL}, Exception:{e.Message}");
                            FAIL(TestBase.PageSnapshotToReport(driver));
                        }
                        driver.SwitchTo().Window(driver.WindowHandles.Last()).Close(); // 關掉新頁
                        driver.SwitchTo().Window(driver.WindowHandles.First()); // 切回原頁
                    }
                }
            }
            #region 關閉 & 釋放文件
            excel_WB.Close();
            excel_Expect_WB.Close();
            excel_App.Quit();
            #endregion
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
