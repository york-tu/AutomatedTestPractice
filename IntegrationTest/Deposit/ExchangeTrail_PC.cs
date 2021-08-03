using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;
using System.Windows.Forms;
using WindowsInput;
using System.Drawing;

namespace AutomatedTest.IntegrationTest.Deposit
{
    public class 官網首頁_換匯試算_PC版:IntegrationTestBase
    {
        public 官網首頁_換匯試算_PC版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void 換匯試算_所有幣值匯率換算_欄位檢查(string browser)
        {
            StartTestCase(browser, " 換匯試算_所有幣值匯率換算_欄位檢查", "York");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\ExchangeRate\ExchangeTrail";
            Tools.CreateSnapshotFolder(snapshotpath);
            Thread.Sleep(100);
            Tools.ScrollPageUpOrDown(driver, 100);

            driver.Navigate().Refresh();
            driver.FindElement(By.XPath("//*[@id='btnAntiFraud']")).Click(); // 關掉"提醒您"彈跳視窗

            if (driver.FindElement(By.XPath("//*[@id='mainform']/div[5]/div/ul")).GetAttribute("style").ToString() == "display: block;")
            {
                driver.FindElement(By.XPath("//*[@id='mainform']/div[5]/div/a")).Click();
            }

            Tools.ScrollPageUpOrDown(driver, 1000);

            driver.FindElement(By.XPath("//*[@id='amount']")).SendKeys("10000"); // 輸入起始幣值

            int currencyAmount = driver.FindElements(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[1]/ul/li/ul/li")).Count; // 獲取起始幣別下拉選單選項數量

            for (int baseCurrencyIndex = 2; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> 第一個幣別
            {
                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[1]/ul"));
                BaseCurrencyDropDownList.Click(); // 點擊起始幣別下拉選單
                string baseCurrencyXPath = $"//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[1]/ul/li/ul/li[{baseCurrencyIndex}]"; // 起始幣別XPath
 
                if (baseCurrencyIndex == 9 || baseCurrencyIndex == 13 || baseCurrencyIndex == 16) // for 滾動選單判斷式
                {
                    Point TargetPosition = new Point(900, 280); //起始幣別選單在畫面上X,Y座標
                    Cursor.Position = TargetPosition; // 滑鼠指標移動到起始幣別下拉選單裡
                    Thread.Sleep(300); 

                    var sim = new InputSimulator();
                    sim.Mouse.VerticalScroll(-1); // 將滑鼠滾輪向下滾一格 >>> 選單slider向下拉一格
                    Thread.Sleep(300);
                }

                driver.FindElement(By.XPath(baseCurrencyXPath)).Click(); // 選擇一個起始幣別
                string BaseCurrencyName = BaseCurrencyDropDownList.Text; // 起始幣別選項名稱

                for (int targetCurrencyIndex = 2; targetCurrencyIndex <= currencyAmount; targetCurrencyIndex++) 
                {
                    IWebElement TargetCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[3]/ul/li"));
                    TargetCurrencyDropDownList.Click(); // 點開目標幣別下拉選單
                    string targetCurrencyXPath = $"//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[3]/ul/li/ul/li[{targetCurrencyIndex}]"; // 目標幣別XPath

                    if (targetCurrencyIndex == 9 || targetCurrencyIndex == 13 || targetCurrencyIndex == 16) // for 滾動選單判斷式
                    {
                        Point TargetPosition = new Point(900, 400); //目標幣別選單在畫面上X,Y座標
                        Cursor.Position = TargetPosition;  // 滑鼠指標移動到目標幣別下拉選單裡
                        Thread.Sleep(300);

                        var sim = new InputSimulator();
                        sim.Mouse.VerticalScroll(-1); // 將滑鼠滾輪向下滾一格 >>> 選單slider向下拉一格
                        Thread.Sleep(300);
                    }

                    driver.FindElement(By.XPath(targetCurrencyXPath)).Click(); // 選擇一個目標幣別
                    string targetCurrencyName = TargetCurrencyDropDownList.Text; // 目標幣別選項名稱

                    driver.FindElement(By.XPath("//*[@id='calculate']")).Click(); // 點 "開始試算"

                    //Tools.PageSnapshot($@"{snapshotpath}\{BaseCurrencyName}_兌換_{targetCurrencyName}_換匯試算.png", driver); // 截圖
                    Tools.ElementSnapshotshot(driver.FindElement(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div")), $@"{snapshotpath}\{BaseCurrencyName}_兌換_{targetCurrencyName}_換匯試算.png");
                    Thread.Sleep(100);
                }
            }
            driver.Quit();
        }
    }
}


