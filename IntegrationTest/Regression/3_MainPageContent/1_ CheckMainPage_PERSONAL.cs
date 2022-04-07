using Xunit;
using AutomatedTest.Utilities;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class D_a_CheckMainPage_PERSONAL:IntegrationTestBase
    {
        public D_a_CheckMainPage_PERSONAL(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Fact]
        public void 檢查個人服務_首頁()
        {
            CreateReport($"個人服務_首頁_內容檢查", "York");

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
            var driver = new ChromeDriver(chromeService,chromeOptions, TimeSpan.FromSeconds(120));

            //var firefoxOptions = new FirefoxOptions();
            ////firefoxOptions.AddArguments("--headless");
            //var driver = new FirefoxDriver(firefoxOptions);
            #endregion

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal");

            #region 讀取 json
            string path = $"{UserDataList.Upperfolderpath}Settings\\PersonalMainPage.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion

            INFO("檢查 [個人服務] 首頁內容 選項名稱 與 連結");
            var totalURLCounts = jsonArray.Count(); // json裡data數量
            for (int i = 0; i < totalURLCounts; i++)
            {
                JObject obj = (JObject)jsonArray[i];

                string targetURL = obj["TargetURL"].ToString();
                string elementCssSelector = obj["PageKeywordCssSelector"].ToString();
                string expectText = obj["ExpectString"].ToString();

                string actualURL = driver.FindElementByCssSelector(elementCssSelector).GetAttribute("href");
                string actualText = driver.FindElementByCssSelector(elementCssSelector).Text;

                #region 1. 檢查網頁 '置頂選單' 項目
                if (i <= 8 )
                {
                    if (i == 0)
                    {
                        INFO(" 檢查網頁置頂選單");
                    }

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
                #endregion

                #region 2. 檢查 '標題 & MegaMenu' 項目
                else if (i >=9 && i <= 20)
                {
                    if (i == 9)
                    {
                        INFO("");
                        INFO("檢查 '標題 & MegaMenu' 項目");
                    }
                    else if (i == 17)
                    {
                        driver.FindElement(By.CssSelector(".l-hearder__dropDownList")).Click();
                        System.Threading.Thread.Sleep(1000);
                    }

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
                #endregion

                #region 3. 檢查 '產品與服務' 項目
                else if (i >= 21 && i <= 28)
                {
                    TestBase.ScrollPageUpOrDown(driver, 500);
                    if (i==21)
                    {
                        INFO("");
                        INFO("檢查 '產品與服務' 項目");
                    }

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
                #endregion

                #region 4. 檢查 '外幣匯率' 項目
                else if (i >= 29 && i <= 34)
                {
                    TestBase.ScrollPageUpOrDown(driver, 1200);
                    if (i == 29)
                    {
                        INFO("");
                        INFO("檢查 '外幣匯率' 項目'");
                    }
                    else if (i==33)
                    {
                        driver.FindElement(By.CssSelector(".color-primary-underline")).Click();
                    }
                    
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
                #endregion

                #region 5. 檢查 '你的生活金融' 項目
                else if (i >= 35 && i <= 38)
                {
                    TestBase.ScrollPageUpOrDown(driver, 1900);
                    if (i == 35)
                    {
                        INFO("");
                        INFO("檢查 '你的生活金融' 項目'");
                    }

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
                #endregion

                #region 6. 檢查 '探索數位服務' 項目
                else if (i >= 39 && i <= 42)
                {
                    TestBase.ScrollPageUpOrDown(driver, 2500);
                    if (i == 39)
                    {
                        INFO("");
                        INFO("檢查 '探索數位服務' 項目'");
                    }

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
                #endregion

                #region 7. 檢查 '最新消息' 項目
                else if (i >= 43 && i <= 44 )
                {
                    TestBase.ScrollPageUpOrDown(driver, 3000);
                    if (i == 43)
                    {
                        INFO("");
                        INFO("檢查 '最新消息' 項目'");
                    }

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
                #endregion

                #region 8. 檢查 '更多連結 與 '置底項目'
                else if (i >= 45)
                {
                    TestBase.ScrollPageUpOrDown(driver, 4000);
                    if (i == 45)
                    {
                        INFO("");
                        INFO("檢查 '更多連結' 項目'");
                    }
                    else if (i == 49 )
                    {
                        INFO("檢查 '關於玉山' 項目'");
                    }
                    else if (i == 53 )
                    {
                        INFO("檢查 '玉山服務網' 項目'");
                    }
                    else if (i == 57 )
                    {
                        INFO("檢查 '玉山社群' 項目'");
                    }
                    else if (i == 60 )
                    {
                        INFO("檢查 '存款保險' 項目'");
                    }
                    else if (i == 61)
                    {
                        INFO("檢查 '網頁置底' 項目'");
                    }

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
                #endregion
            }
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
