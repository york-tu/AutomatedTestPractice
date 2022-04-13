using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using AutomatedTest.Utilities;

using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Linq;
using System.Globalization;
using OpenQA.Selenium.Interactions;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class test:IntegrationTestBase
    {
        public test(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Fact]

        public void test_test()
        {
   
            //CreateReport($"個人服務_首頁_內容檢查", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");

            #region  setup browser 不開啟網頁
            //Chrome headless 參數設定
            var chromeOptions = new ChromeOptions();
            //chromeOptions.AddArguments("--headless");
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

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal/wealth/offshore-bond/intro");
            //driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(600);
            TestBase.ScrollPageUpOrDown(driver, 500); // 常用服務
            //TestBase.ScrollPageUpOrDown(driver, 800); // 基金e指選

            //TestBase.ScrollPageUpOrDown(driver, 1500); // 即時查詢

            //TestBase.ScrollPageUpOrDown(driver, 2000); // 熱門推薦

            //TestBase.ScrollPageUpOrDown(driver, 2800); // 理財電子週報

            //TestBase.ScrollPageUpOrDown(driver, 3500); // 市場資訊 常見問題 聯繫客服

            //TestBase.ScrollPageUpOrDown(driver, 4000); // 置底


            //driver.SwitchTo().Frame("iframe1");
            var aaa= driver.FindElement(By.CssSelector("#layout_0_rightcontent_0_h1Container")).Text;
            var bbb= driver.FindElement(By.CssSelector(".line > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > img:nth-child(2)")).GetAttribute("title");
            var ccc = driver.FindElement(By.CssSelector(".line > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > img:nth-child(2)")).GetAttribute("alt");

            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
