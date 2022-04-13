using Xunit;
using AutomatedTest.Utilities;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class D_b_CheckMainPage_Deposit:IntegrationTestBase
    {
        public D_b_CheckMainPage_Deposit(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }
        string testCaseName = "存匯";
        string mainPageURL = "https://www.esunbank.com.tw/bank/personal/deposit";
        string jsonPath = $@"{UserDataList.Upperfolderpath}Settings\depositMainPage.json"; // json檔路徑

        [Fact]
        public void 檢查首頁()
        {
            #region 讀 json 資料語法
            var jsonContent = File.ReadAllText(jsonPath);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion
            var totalURLCounts = jsonArray.Count(); // json裡資料數

            #region  Chrome瀏覽器不開啟網頁設定(headless)
            //Chrome headless 參數設定
            var chromeService = ChromeDriverService.CreateDefaultService();
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
            var driver = new ChromeDriver(chromeService, chromeOptions, TimeSpan.FromSeconds(120));
            #endregion
            driver.Navigate().GoToUrl(mainPageURL);

            CreateReport($"{testCaseName}_檢查首頁內容與連結", "York");

            for (int i = 0; i < totalURLCounts; i++)
            {
                JObject obj = (JObject)jsonArray[i];

                string targetURL = obj["TargetURL"].ToString();
                string elementCssSelector = obj["PageKeywordCssSelector"].ToString();
                string expectText = obj["ExpectString"].ToString();

                string actualURL = driver.FindElementByCssSelector(elementCssSelector).GetAttribute("href");
                string actualText = driver.FindElementByCssSelector(elementCssSelector).Text;

                #region 新增標題
                switch (i)
                {
                    case 0:
                        // 1. 檢查網頁 '置頂選單' 項目
                        INFO(" 檢查網頁置頂選單");
                        break;
                    case 9:
                        // 2. 檢查 '標題 & MegaMenu' 項目
                        INFO("");
                        INFO("檢查 '標題 & MegaMenu' 項目");
                        break;
                    case 17:
                        driver.FindElement(By.CssSelector(".l-hearder__dropDownList")).Click();
                        System.Threading.Thread.Sleep(1000);
                        break;
                    case 21:
                        // 3. 檢查 '常用服務' 項目
                        TestBase.ScrollPageUpOrDown(driver, 500);
                        INFO("");
                        INFO("檢查 '常用服務' 項目");
                        break;
                    case 29:
                        // 4. 檢查 '線上開立玉山數位帳戶' 項目
                        TestBase.ScrollPageUpOrDown(driver, 800);
                        INFO("");
                        INFO("檢查 '外幣匯率' 項目'");
                        break;
                    case 30:
                        // 5. 檢查 '外幣優惠' 項目
                        TestBase.ScrollPageUpOrDown(driver, 800);
                        INFO("");
                        INFO("檢查 '外幣優惠' 項目'");
                        break;
                    case 35:
                        // 6. 檢查 '熱門推薦' 項目
                        TestBase.ScrollPageUpOrDown(driver, 2100);
                        INFO("");
                        INFO("檢查 ''熱門推薦' 項目'");
                        break;
                    case 38:
                        // 7.檢查 '各項服務申請' 項目
                        TestBase.ScrollPageUpOrDown(driver, 2500);
                        INFO("");
                        INFO("檢查 '各項服務申請' 項目'");
                        break;
                    case 42:
                        TestBase.ScrollPageUpOrDown(driver, 2800);
                        INFO("");
                        INFO("檢查 '收費標準及公告說明' 項目'");
                        break;
                    case 50:
                        // 8. 檢查 '收費標準及公告說明' 項目
                        TestBase.ScrollPageUpOrDown(driver, 3200);
                        INFO("");
                        INFO("檢查 '常見問題' 項目'");
                        break;
                    case 55:
                        // 9. 檢查 '常見問題' 與 '聯繫客服' 項目
                        INFO("");
                        INFO("檢查 '聯繫客服' 項目'");
                        break;
                    case 59:
                        // 10. 檢查 '置底項目'
                        TestBase.ScrollPageUpOrDown(driver, 4000);
                        INFO("");
                        INFO("檢查 '網頁置底' 項目'");
                        break;
                    default:
                        break;
                }
                #endregion

                if (actualURL != targetURL)
                {
                    FAIL($"[連結錯誤] 預期: {targetURL}, 實際: {actualURL}");
                }
                else if (actualText != expectText)
                {
                    FAIL($"[標題錯誤] 預期: {expectText}, 實際: {actualText}");
                }
                else
                {
                    PASS($"[連結與標題正確] 連結: {actualURL}, 標題: {actualText}");
                }
            }
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }

        [Fact]
        public void 檢查內文連結導引頁()
        {
            #region 讀json資料語法
            var jsonContent = File.ReadAllText(jsonPath);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion
            int totalURLCounts = jsonArray.Count(); // json裡資料數

            #region Chrome瀏覽器不開啟網頁設定(headless) 

            var chromeService = ChromeDriverService.CreateDefaultService();
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("--headless");
            chromeOptions.AddArguments("--disable-gpu");
            chromeOptions.AddArguments("--incognito");
            chromeOptions.AddArguments("--window-size=1440x900");
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

            var driver = new ChromeDriver(chromeService, chromeOptions, TimeSpan.FromSeconds(300));
            #endregion

            CreateReport($"{testCaseName}_內文連結導引頁", "York");
           
            for (int i = 0; i < totalURLCounts; i++)
            {
                JObject obj = (JObject)jsonArray[i];
                string uRL = obj["TargetURL"].ToString();
                string cssSelector = obj["DirectPageElementCssSelector"].ToString();
                string expectString = obj["DirectPageKeyword"].ToString();

                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // 瀏覽器另開新頁
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on 新頁上 
                driver.Navigate().GoToUrl(uRL);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600);

                #region 當網頁自動切到M版時, click "切換電腦版" 強制切回PC版
                if (driver.Url.ToString().Contains("?dev=mobile"))
                {
                    TestBase.ScrollPageUpOrDown(driver, 5000);
                    if (driver.Url.ToString().Contains("www.esunfhc.com"))
                    {
                        driver.FindElementById("fhc_layout_m_0_fhc_maincontent_m_2_HlkToWeb").Click();
                    }
                    else
                    {
                        driver.FindElementByClassName("changeTarget").Click();
                    }
                }
                #endregion

                if (driver.Url.ToString() != uRL) // 網頁開啟當下網址與預期不符 >>> 網頁redirect
                {
                    WARNING($"[Page Redirect], {driver.Url.ToString()}  (Expect: {uRL })");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }
                // Wordaround 當網頁為以下網址 >>> 人工 check
                else if (uRL == "https://gib.esunbank.com/" || uRL == "https://netbank.esunbank.com.tw/webatm/#/login" || uRL == " https://www.esunbank.com.tw/bank/-/media/esunbank/files/credit-card/2017card_ex.pdf?la=en")
                {
                    WARNING($"[NeedManualCheck], {driver.Url.ToString()}  (Expect: {uRL })");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }

                string actualText = "";
                if (uRL == "https://www.esunbank.com.tw/event/credit/1040408web/index.htm" || uRL == "https://www.esunbank.com.tw/event/credit/1100412home_al/index.html" || uRL == "https://accessible.esunbank.com.tw/Accessibility/Index")
                {
                    actualText = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("alt");
                }
                else if (uRL == "https://ebank.esunbank.com.tw/index.jsp")
                {
                    driver.SwitchTo().Frame("iframe1");
                    actualText = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("title");
                }
                else
                {
                    actualText = driver.FindElement(By.CssSelector(cssSelector)).Text;
                }

                try
                {
                    Assert.Contains(expectString, actualText); // 判斷element 字串是否符合預期
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
