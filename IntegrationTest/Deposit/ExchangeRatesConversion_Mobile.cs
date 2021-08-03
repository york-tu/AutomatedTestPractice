using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;

namespace AutomatedTest.IntegrationTest.Deposit
{
    public class 即期匯率_匯率換算_M版:IntegrationTestBase
    {
        public 即期匯率_匯率換算_M版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/deposit/rate/forex/currency-converter?dev=mobile";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void 即期匯率_所有幣值匯率換算_欄位檢查(string browser)
        {
            StartTestCase(browser, "即期匯率_所有幣值匯率換算_欄位檢查", "York");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\ExchangeRate";
            Tools.CreateSnapshotFolder(snapshotpath);
            Thread.Sleep(100);
            Tools.ScrollPageUpOrDown(driver, 100);

            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/input")).SendKeys("10000"); //輸入起始幣值
            int currencyAmount = driver.FindElements(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li")).Count; // 獲取起始幣別下拉選單選項數量
            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/div[2]/a[2]")).Click(); // 點 開始試算

            for (int baseCurrencyIndex = 2; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> 第一個幣別
            {
                Tools.ScrollPageUpOrDown(driver, 100);
                Thread.Sleep(100);

                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li"));
                BaseCurrencyDropDownList.Click(); // 點擊起始幣別下拉選單
                string baseCurrencyXPath = $"//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li[{baseCurrencyIndex}]/span"; // 起始幣別XPath
                driver.FindElement(By.XPath(baseCurrencyXPath)).Click(); // 選擇一個起始幣別
                string BaseCurrencyName = BaseCurrencyDropDownList.Text; // 起始幣別選項名稱

                for (int targetCurrencyIndex = 2; targetCurrencyIndex <= currencyAmount; targetCurrencyIndex++)
                {
                    Tools.ScrollPageUpOrDown(driver, 200); 
                    Thread.Sleep(100);

                    if (baseCurrencyIndex != targetCurrencyIndex) 
                    {
                        IWebElement TargetCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[2]/div/ul/li"));
                        TargetCurrencyDropDownList.Click(); // 點開目標幣別下拉選單
                        string targetCurrencyXPath = $"//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[2]/div/ul/li/ul/li[{targetCurrencyIndex}]"; // 目標幣別XPath
                        driver.FindElement(By.XPath(targetCurrencyXPath)).Click(); // 選擇一個目標幣別
                        string targetCurrencyName = TargetCurrencyDropDownList.Text; // 目標幣別選項名稱
                        Tools.ScrollPageUpOrDown(driver, 300);
                        Thread.Sleep(100);
                        Tools.PageSnapshot($@"{snapshotpath}\{BaseCurrencyName}_兌換_{targetCurrencyName}_M.png", driver);
                        Thread.Sleep(100);
                    }
                    else //起始幣別與目標幣別相同時, skip, 跳過
                    {
                        continue;
                    }
                }
            }
            driver.Quit();
        }
    }
}


