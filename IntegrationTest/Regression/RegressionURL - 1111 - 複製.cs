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

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal/deposit");
            TestBase.ScrollPageUpOrDown(driver, 300);
            TestBase.ScrollPageUpOrDown(driver, 800);
            TestBase.ScrollPageUpOrDown(driver, 1000);
            TestBase.ScrollPageUpOrDown(driver, 1200);
            TestBase.ScrollPageUpOrDown(driver, 1500);
            TestBase.ScrollPageUpOrDown(driver, 1800);
            TestBase.ScrollPageUpOrDown(driver, 2100);
            TestBase.ScrollPageUpOrDown(driver, 2300);
            TestBase.ScrollPageUpOrDown(driver, 2500);
            TestBase.ScrollPageUpOrDown(driver, 2800);
            TestBase.ScrollPageUpOrDown(driver, 3000);
            TestBase.ScrollPageUpOrDown(driver, 3200);
            TestBase.ScrollPageUpOrDown(driver, 3500);
            TestBase.ScrollPageUpOrDown(driver, 3800);
            TestBase.ScrollPageUpOrDown(driver, 4000);
            TestBase.ScrollPageUpOrDown(driver, 4200);


            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
