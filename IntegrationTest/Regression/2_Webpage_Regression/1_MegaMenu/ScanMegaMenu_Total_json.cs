using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using AutomatedTest.Utilities;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Linq;
using OpenQA.Selenium.Interactions;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class B_a_CheckMegaMenu_Total:IntegrationTestBase
    {
        public B_a_CheckMegaMenu_Total(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Theory]
        /* MegaMenu選單
         * 1: 存匯
         * 2: 貸款
         * 3: 信用卡/支付
         * 4: 財富管理
         * 5: 信託
         * 6: 保險
         * 7: 生活金融
         */
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(7)]
        public void 檢查TotalMegaMenu_excel(int megaMenuIndex)
        {
            #region step 1. switch測項
            string testCaseName = "";
            int row = 0, line = 0; // 設定 megamenu 選單 "行 & 列" 數
            switch(megaMenuIndex)
            {
                case 1:
                    testCaseName = "存匯";
                    row = 6; line = 10;
                    break;
                case 2:
                    testCaseName = "貸款";
                    row = 5; line = 10;
                    break;
                case 3:
                    testCaseName = "信用卡_支付";
                    row = 6; line = 10;
                    break;
                case 4:
                    testCaseName = "財富管理";
                    row = 6; line = 10;
                    break;
                case 5:
                    testCaseName = "信託";
                    row = 5; line = 8;
                    break;
                case 6:
                    testCaseName = "保險";
                    row = 4; line = 8;
                    break;
                case 7:
                    testCaseName = "生活金融";
                    row = 1; line = 8;
                    break;
                default:
                    break;
            }
            #endregion

            CreateReport($"[{testCaseName}] MegaMenu 內容檢查", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");

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

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal");

            #region step 3: 讀json資料
            string path = $"{UserDataList.Upperfolderpath}Settings\\URL_Css_MegaMenu_ExpectString.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JArray.Parse(jsonContent);
            #endregion

            Actions act = new Actions(driver);
            if (megaMenuIndex == 7 ) // 當測項為 "7:生活金融" 時, 切換不同架構CssSelector
            {
                act.MoveToElement(driver.FindElement(By.CssSelector(".l-megaL1__include > a:nth-child(1)"))).Perform();
            }
            else
            {
                // Cursor移到 MegaMenu [testCaseName] 選單
                act.MoveToElement(driver.FindElement(By.CssSelector($"ul.l-megaL1__list:nth-child(1) > li:nth-child({megaMenuIndex}) > div:nth-child(1)"))).Perform(); 
            }

            // 循環掃megamenu選項 (i=行, j=列)
            for (int i = 1; i<=row ; i++) 
            {
                for (int j = 0; j<=line ; j++) 
                {
                    var cssSelector = "";

                    if (megaMenuIndex == 7) // 當選項為 "7:生活金融" 時, 切換不同架構CssSelector
                    {
                        cssSelector = $".l-megaL2__noInclude > div:nth-child(1) > ul:nth-child(1) > li:nth-child({j}) > a:nth-child(1)";
                    }
                    else if (megaMenuIndex < 7 && j == 0) // 選項不為 "7:生活金融" & 針對 megamenu 次層粗黑體字大標 (j=0)
                    {
                        cssSelector = $"ul.l-megaL1__list:nth-child(1) > li:nth-child({megaMenuIndex}) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child({i}) > a:nth-child(1)";
                    }
                    else
                    {
                        cssSelector = $"ul.l-megaL1__list:nth-child(1) > li:nth-child({megaMenuIndex}) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child({i}) > ul:nth-child(2) > li:nth-child({j}) > a:nth-child(1)";
                    }

                    try
                    {
                        string actualOptionName = driver.FindElement(By.CssSelector(cssSelector)).Text; // 網頁MegaMenu上選項名稱
                        string actURL = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("href"); // 網頁MegaMenu上選項對應超連結

                        IEnumerable<JToken> tokens = jsonArray.SelectTokens($"[?(@.MegaMenuOptions == '{actualOptionName}')]",true); // 搜尋選項對應 json 資料位置
                        if (tokens.Any() == false) // falas >>> json 內搜尋不到
                        {
                            FAIL($"[Mega Menu][{testCaseName}] 選項\"{actualOptionName}\''不在預期選項內");
                            continue;
                        }
                        else
                        {
                            JToken targetToken = i <=3 ? targetToken = tokens.First() : targetToken = tokens.Last(); // for同選項名稱, 但不同內容
                            var targetArrayIndex = Convert.ToInt32((targetToken.Path).Substring(1, (targetToken.Path).Length - 2));
                            var expectURL = targetToken["TargetURL"].ToString();

                            if (actURL == expectURL)
                            {
                                PASS($"[Mega Menu][{testCaseName}] 選項 \" {actualOptionName}\" 超連結網址  \"{actURL}\" 符合預期");
                            }
                            else if (actURL != expectURL)
                            {
                                FAIL($"[Mega Menu][{testCaseName}] 選項 \"{actualOptionName}\" 超連結網址 \"{actURL}\" (預期網址 \"{expectURL}\")");
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
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
