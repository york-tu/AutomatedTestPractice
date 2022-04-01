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
using System.IO;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class II_I_檢查網頁內容_json:IntegrationTestBase
    {
        public II_I_檢查網頁內容_json(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }
        [Theory]
        [InlineData(0, 200)]
        [InlineData(200, 400)]
        [InlineData(400, 600)]
        [InlineData(600, 750)]
        public void 檢查網頁內容是否符合預期(int startIndex, int endIndex)
        {
            #region step 0: kill cache driver
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");
            #endregion

            #region step 1: 讀json資料
            string path = $"{UserDataList.Upperfolderpath}Settings\\ExcelToJson.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion

            var totalURLCounts = jsonArray.Count(); // json裡data數量
            endIndex = endIndex >= 701 ? endIndex = totalURLCounts : endIndex;
            CreateReport($"網頁內容檢查_json_{startIndex}-{endIndex}", "York");

            #region step 2:  browser 不開啟網頁設定 
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

            for (int i = startIndex; i < endIndex; i++)
            {
                JObject obj = (JObject)jsonArray[i];
                string uRL = obj["TargetURL"].ToString();
                string cssSelector = obj["PageKeywordCssSelector"].ToString();
                string expectString = obj["ExpectString"].ToString();

                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // 另開新頁
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on 新頁上 
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                driver.Url = uRL;
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600秒內載完網頁內容, 否則報錯, 載完提早進下一步.

                #region 當網頁自動切到M版時, click "切換電腦版" 強制切回PC版
                if (driver.Url.ToString().Contains("?dev=mobile")) // workaround: 當網頁自動切到M版時, 強制切回PC版
                {
                    TestBase.ScrollPageUpOrDown(driver, 1500);
                    driver.FindElementByClassName("changeTarget").Click();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                }
                #endregion

                if (uRL == "https://easyfee.esunbank.com.tw/index.action" || uRL == "https://www.esunbank.com.tw/bank/marketing/loan/marketing-menu-level-2" || uRL == "https://ebank.esunbank.com.tw/index.jsp#toTaskId=FIM01007")
                {
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                    System.Threading.Thread.Sleep(3000);
                    WARNING($"{uRL }, Keyword: {expectString}");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }
                else if (driver.Url.ToString() != uRL) // 檢查網頁開啟當下網址是否為輸入網址 (判斷網頁是否有redirect)
                {
                    WARNING($"[Page Redirect], {driver.Url.ToString()}  (Expect: {uRL })");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }
                else if (driver.Url.ToString() == "https://www.esunbank.com.tw/bank/personal/event/calendar/events") // Workaround 1 : 活動日曆URL >>> 需抓當天日期
                {
                    string day_of_week = DateTime.Now.ToString("dddd", new CultureInfo("en-us")); // 英文星期幾(e.g., Wednesday)
                    string day = DateTime.Now.ToString("%d"); // 日期 (e.g., 1)
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
                    PASS($"{uRL}, Keyword: {expectString}");
                    PASS(TestBase.PageSnapshotToReport(driver));
                }
                catch (Exception e)
                {
                    FAIL($"{uRL}, Exception:{e.Message}");
                    FAIL(TestBase.PageSnapshotToReport(driver));
                }
                driver.SwitchTo().Window(driver.WindowHandles.Last()).Close(); // 關掉新頁
                driver.SwitchTo().Window(driver.WindowHandles.First()); // 切回原頁
            }

            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
