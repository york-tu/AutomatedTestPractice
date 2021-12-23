using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using WindowsInput;
using System.Drawing;

namespace AutomatedTest.IntegrationTest.Personal
{
    public class 首頁右側常駐快捷選單__試算工具_PC版:IntegrationTestBase
    {
        public 首頁右側常駐快捷選單__試算工具_PC版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void 常駐快捷選單_所有幣值匯率換算_欄位檢查(string browser)
        {
            StartTestCase(browser, " 常駐快捷選單_所有幣值匯率換算_欄位檢查", "York");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\ExchangeRate";
            TestBase.CreateFolder(snapshotpath);
            Thread.Sleep(100);
            TestBase.ScrollPageUpOrDown(driver, 100);

            driver.Navigate().Refresh(); // refresh網頁 >>> 讓首頁右側快捷選單顯示
            driver.FindElement(By.XPath("//*[@id='btnAntiFraud']")).Click(); // 關掉"提醒您"彈跳視窗

            Actions act = new Actions(driver);
            act.MoveToElement(driver.FindElement(By.XPath("//*[@id='mainform']/div[5]/div/ul/li[2]"))).Perform(); // 滑鼠"移到" 右側常駐 "試算工具"選單

            driver.SwitchTo().Frame(1); // 切到 "試算工具"選單 frame
            driver.FindElement(By.XPath("//*[@id='amount']")).SendKeys("10000"); // 輸入起始幣值

            Point CursorHover = new Point(1550, 500);
            Cursor.Position = CursorHover; // 滑鼠指標停在試算選單頁面上 (避免當沒focus在選單時選單自動hide)

            int currencyAmount = driver.FindElements(By.XPath("//*[@id='tabs2-01']/div/div[1]/div[1]/ul/li/ul/li")).Count; // 獲取起始幣別下拉選單選項數量

            for (int baseCurrencyIndex = 2; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> 第一個幣別
            {
                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='tabs2-01']/div/div[1]/div[1]/ul/li"));
                BaseCurrencyDropDownList.Click(); // 點擊起始幣別下拉選單
                string baseCurrencyXPath = $"//*[@id='tabs2-01']/div/div[1]/div[1]/ul/li/ul/li[{baseCurrencyIndex}]/span"; // 起始幣別XPath

                if (baseCurrencyIndex == 5 || baseCurrencyIndex == 8 || baseCurrencyIndex == 12 || baseCurrencyIndex == 15) // for 滾動選單判斷式
                {
                    Point TargetPosition = new Point(1375, 550); //起始幣別選單在畫面上X,Y座標
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
                    IWebElement TargetCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='tabs2-01']/div/div[1]/div[3]/ul/li"));
                    TargetCurrencyDropDownList.Click(); // 點開目標幣別下拉選單
                    string targetCurrencyXPath = $" //*[@id='tabs2-01']/div/div[1]/div[3]/ul/li/ul/li[{targetCurrencyIndex}]/span"; // 目標幣別XPath

                    if (targetCurrencyIndex == 5 || targetCurrencyIndex == 8 || targetCurrencyIndex == 12 || targetCurrencyIndex == 15) // for 滾動選單判斷式
                    {
                        Point TargetPosition = new Point(1375, 650); //目標幣別選單在畫面上X,Y座標
                        Cursor.Position = TargetPosition;  // 滑鼠指標移動到目標幣別選單裡
                        Thread.Sleep(300);

                        var sim = new InputSimulator();
                        sim.Mouse.VerticalScroll(-1); // 將滑鼠滾輪向下滾一格 >>> 選單slider向下拉一格
                        Thread.Sleep(300);
                    }

                    driver.FindElement(By.XPath(targetCurrencyXPath)).Click(); // 選擇一個目標幣別
                    string targetCurrencyName = TargetCurrencyDropDownList.Text; // 目標幣別選項名稱

                    driver.FindElement(By.XPath("//*[@id='calculate']")).Click(); // 點 "開始試算"

                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\{BaseCurrencyName}_兌換_{targetCurrencyName}_PC.png"); // 截圖
                    Thread.Sleep(100);
                }
            }
            driver.Quit();
        }
    }
}


