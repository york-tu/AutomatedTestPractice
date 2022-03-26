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
using OpenQA.Selenium.Interactions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class II_III_存匯內文_MegaMenu檢查:IntegrationTestBase
    {
        public II_III_存匯內文_MegaMenu檢查(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Fact]
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
        public void 檢查存匯MegaMenu()
        {
            CreateReport("存匯MegaMenu內容檢查", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");
            Excel.Application excel_App = new Excel.Application(); //  新增excel應用程序



            #region step 1: 讀取 expect result excel (expectResult.xlsx) (excel_Expect_xxxx) 第四行(option名稱)到 expectResultArray 陣列中, 
            Excel.Workbook excel_Expect_WB = excel_App.Workbooks.Open($"{UserDataList.Upperfolderpath}testdata\\ExpectResult_0323.xlsx"); // open 指定路徑excel
            Excel.Worksheet excel_Expect_WS = (Excel.Worksheet)excel_Expect_WB.Worksheets[1]; // 指定讀取excel 檔第一個工作表
            int sheetRows = 4; // 工作表內行數 
            Excel.Range expectResultRange = (Excel.Range)excel_Expect_WS.UsedRange; // export excel 內容 to Range
            string[] expectResultOptionsNameArray = new string[expectResultRange.Count / sheetRows];
            for (int i = 0; i < expectResultRange.Count / sheetRows; i++)
            {
                expectResultOptionsNameArray[i] = (string)((Excel.Range)expectResultRange.Cells[i + 1, 4]).Value; // 將excel第四行內容丟進expectResultArray陣列中
            }
            #endregion

            #region  step 2: setup browser 不開啟網頁
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
            driver.Url = "https://www.esunbank.com.tw/bank/personal";
            Actions act = new Actions(driver);
            act.MoveToElement(driver.FindElement(By.CssSelector("ul.l-megaL1__list:nth-child(1) > li:nth-child(1) > div:nth-child(1)"))).Perform(); // Cursor移到 MegaMenu "存匯"選單

            for (int row= 1; row <=10; row++) // megamenu 存匯10行
            {
                for (int line = 1; line <= 10; line++) // megamenu 存匯10列
                {
                    try
                    {
                        var cssSelector = $"ul.l-megaL1__list:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child({row}) > ul:nth-child(2) > li:nth-child({line}) > a:nth-child(1)";
                        var actualOptionName = driver.FindElement(By.CssSelector(cssSelector)).Text; // 目前網頁MegaMenu上選項名稱
                        var actURL = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("href"); // 目前網頁MegaMenu上選項對應超連結

                        int meetExpectResultIndex = Array.IndexOf(expectResultOptionsNameArray, actualOptionName); // 目前網頁MegaMenu選項對應到expect result 陣列中的位置
                        if (meetExpectResultIndex == -1)
                        {
                            FAIL($"[Mega Menu][存匯] 選項\"{actualOptionName}\''不在預期選項內");
                            continue;
                        }
                        else
                        {
                            string expectURL = (string)((Excel.Range)expectResultRange.Cells[meetExpectResultIndex + 1, 1]).Value; // 讀 expect result excel中第一行欄位值

                            if (actURL == expectURL)
                            {
                                PASS($"[Mega Menu][存匯] 選項 \" {actualOptionName}\" 超連結網址  \"{actURL}\" 符合預期");
                            }
                            else if (actURL != expectURL)
                            {
                                FAIL($"[Mega Menu][存匯] 選項 \"{actualOptionName}\" 超連結網址 \"{actURL}\" (預期網址 \"{expectURL}\")");
                                // FAIL(TestBase.PageSnapshotToReport(driver));
                            }
                        }
                     }
                    catch (Exception e)
                    {
                        //WARNING(e.Message);
                    }
                }
            }

            #region 關閉 & 釋放文件
            excel_Expect_WB.Close();
            excel_App.Quit();
            #endregion
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
        public void 檢查存匯內文()
        {

        }
    }
}
