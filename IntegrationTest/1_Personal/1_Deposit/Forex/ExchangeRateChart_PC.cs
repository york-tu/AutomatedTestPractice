using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;

namespace AutomatedTest.IntegrationTest.Personal.Deposit.Forex
{
    public class 外匯匯率_歷史匯率走勢_PC版:IntegrationTestBase
    {
        public 外匯匯率_歷史匯率走勢_PC版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/deposit/rate/forex/exchange-rate-chart"; // 登入PC版 "歷史匯率走勢" 網頁
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void 幣別查詢(string browser)
        {
            StartTestCase(browser, "歷史匯率走勢_幣別查詢", "York");
            INFO("");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\Forex";
            TestBase.CreateFolder(snapshotpath);
            TestBase.CleanUPFolder(snapshotpath);

            int dropDownListCount = driver.FindElements(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li/ul/li")).Count; // 下拉選單選項數量
            for (int index = 2; index <= dropDownListCount; index++) // inxde = 2 = 第一筆可選幣別
            {
                TestBase.ScrollPageUpOrDown(driver, 450);
                Thread.Sleep(300);
                driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li")).Click(); // 點開幣別下拉選單
                string currencyXPath = $"//*[@id='tab-01']/div[1]/div[1]/ul/li/ul/li[{index}]";
                driver.FindElement(By.XPath(currencyXPath)).Click(); // 選擇一個幣別

                IWebElement StartDate = driver.FindElement(By.XPath("//*[@id='datepicker1']"));
                StartDate.Clear();
                StartDate.SendKeys("2021-07-01");
                Thread.Sleep(300);
                StartDate.Click(); // (避免輸入日期時彈出的日曆小視窗影響底下選項點擊, 點欄位確保小日曆視窗關閉)
                
                for (int i = 1; i <= 2 ; i++) // i=1=現金radiobutton, i=2=即期radiobutton
                {
                    TestBase.ScrollPageUpOrDown(driver, 450);
                    Thread.Sleep(300);

                    string CurrencyName = driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li")).Text;
                    
                    // 因以下幣別不提供現金選項 >>> 只查詢 "即期"
                    if (CurrencyName == "南非幣(ZAR)" || CurrencyName == "紐西蘭幣(NZD)" || CurrencyName == "墨西哥披索(MXN)" || CurrencyName == "瑞士法郎(CHF)" || CurrencyName == "瑞典幣(SEK)" || CurrencyName == "新加坡幣(SGD)" || CurrencyName == "泰銖(THB)")
                    {
                        i = 2;
                    }

                    driver.FindElement(By.XPath($"//*[@id='tab-01']/div[1]/div[1]/div/label[{i}]")).Click(); // 點 "現金/即期"

                    for (int j = 1; j <= 11; j++) // 連續切換走勢圖/數據表 (為了避免點擊選項後沒有正確切換, 切換多次以確保確實切換至數據表)
                    {
                        driver.FindElement(By.XPath($"//*[@id='tab-01']/div[1]/div[3]/div/label[{j % 2 + 1}]")).Click(); // 點 "走勢圖/數據表"
                        Thread.Sleep(100);
                        driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[4]/a")).Click(); // 點 "開始查詢"
                    }
                    //*[@id="tab-01"]/div[1]/div[1]/ul/li/ul/li[5]/span
                    int pagelist = driver.FindElements(By.XPath("//*[@id='pagelist']/ul/li")).Count; // 獲取分頁數量
                    for (int list = 1; list <= pagelist; list++)
                    {
                        TestBase.ScrollPageUpOrDown(driver, 900);
                        Thread.Sleep(300);
                        driver.FindElement(By.XPath($"//*[@id='pagelist']/ul/li[{list}]/a")).Click(); // 點分頁
                        Thread.Sleep(100);

                        IWebElement CurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li")); // 幣別下拉選單
                        IWebElement CashSort = driver.FindElement(By.XPath($"//*[@id='tab-01']/div[1]/div[1]/div/label[{i}]")); // 現金/即期
                        TestBase.PageSnapshot(driver,$@"{snapshotpath}\{CurrencyDropDownList.Text}_{CashSort.Text}_page{list}.png");
                        
                        Thread.Sleep(100);
                    }
                }
            }
            driver.Quit();
        }
    }
}
