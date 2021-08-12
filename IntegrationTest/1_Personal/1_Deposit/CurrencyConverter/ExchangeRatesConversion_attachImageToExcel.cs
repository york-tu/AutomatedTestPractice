using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;



namespace AutomatedTest.IntegrationTest.Personal.Deposit.CurrencyConverter
{
    public class 練習ToExcel_即期匯率_匯率換算_M版:IntegrationTestBase
    {
        public 練習ToExcel_即期匯率_匯率換算_M版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
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
            Tools.CleanUPFolder(snapshotpath);
            Thread.Sleep(100);
            Tools.ScrollPageUpOrDown(driver, 100);
            string time = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "即期匯率_所有幣值匯率換算_欄位檢查";
            int cell = 2;

            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/input")).SendKeys("10000"); //輸入起始幣值
            int currencyAmount = driver.FindElements(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li")).Count; // 獲取起始幣別下拉選單選項數量
            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/div[2]/a[2]")).Click(); // 點 開始試算

            for (int baseCurrencyIndex = 10; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> 第一個幣別, (default baseCurrencyIndex = 2)
            {
                Tools.ScrollPageUpOrDown(driver, 100);
                Thread.Sleep(100);

                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li"));
                BaseCurrencyDropDownList.Click(); // 點擊起始幣別下拉選單
                string baseCurrencyXPath = $"//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li[{baseCurrencyIndex}]/span"; // 起始幣別XPath
                driver.FindElement(By.XPath(baseCurrencyXPath)).Click(); // 選擇一個起始幣別
                string BaseCurrencyName = BaseCurrencyDropDownList.Text; // 起始幣別選項名稱

                for (int targetCurrencyIndex = 10; targetCurrencyIndex <= currencyAmount; targetCurrencyIndex++) //default targetCurrencyIndex = 2
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


                        Excel.Range oRange = (Excel.Range)xlWorkSheet.Cells[cell+1, 2];
                        float Left = (float)((double)oRange.Left);
                        float Top = (float)((double)oRange.Top);

                        xlWorkSheet.Cells[cell, 1] = $"{BaseCurrencyName}_兌換_{targetCurrencyName}";
                        xlWorkSheet.Shapes.AddPicture($@"{snapshotpath}\{BaseCurrencyName}_兌換_{targetCurrencyName}_M.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left,Top, 750, 350); 
                        cell +=23;


                    }
                    else //起始幣別與目標幣別相同時, skip, 跳過
                    {
                        continue;
                    }
                }
            }
            xlWorkBook.SaveAs($@"d:\即期匯率_所有幣值匯率換算_{time}.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            driver.Quit();
        }
    }
}


