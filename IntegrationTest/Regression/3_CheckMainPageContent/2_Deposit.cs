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

        [Fact]
        public void 檢查存匯_首頁()
        {
            CreateReport($"存匯_首頁_內容檢查", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");

            #region  Browser 不開啟網頁設定
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

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal/deposit");

            #region 讀取 json
            string path = $"{UserDataList.Upperfolderpath}Settings\\DepositMainPage.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion

            INFO("檢查 [存匯] 首頁內容 選項名稱 與 連結");
            INFO("");
            var totalURLCounts = jsonArray.Count(); // json裡data數量
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
     }
}
