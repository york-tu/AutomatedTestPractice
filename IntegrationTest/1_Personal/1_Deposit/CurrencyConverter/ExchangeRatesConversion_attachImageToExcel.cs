using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;



namespace AutomatedTest.IntegrationTest.Personal.Deposit.CurrencyConverter
{
    public class �m��ToExcel_�Y���ײv_�ײv����_M��:IntegrationTestBase
    {
        public �m��ToExcel_�Y���ײv_�ײv����_M��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/deposit/rate/forex/currency-converter?dev=mobile";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void �Y���ײv_�Ҧ����ȶײv����_����ˬd(string browser)
        {
            StartTestCase(browser, "�Y���ײv_�Ҧ����ȶײv����_����ˬd", "York");

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
            xlWorkSheet.Cells[1, 1] = "�Y���ײv_�Ҧ����ȶײv����_����ˬd";
            int cell = 2;

            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/input")).SendKeys("10000"); //��J�_�l����
            int currencyAmount = driver.FindElements(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li")).Count; // ����_�l���O�U�Կ��ﶵ�ƶq
            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/div[2]/a[2]")).Click(); // �I �}�l�պ�

            for (int baseCurrencyIndex = 10; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> �Ĥ@�ӹ��O, (default baseCurrencyIndex = 2)
            {
                Tools.ScrollPageUpOrDown(driver, 100);
                Thread.Sleep(100);

                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li"));
                BaseCurrencyDropDownList.Click(); // �I���_�l���O�U�Կ��
                string baseCurrencyXPath = $"//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li[{baseCurrencyIndex}]/span"; // �_�l���OXPath
                driver.FindElement(By.XPath(baseCurrencyXPath)).Click(); // ��ܤ@�Ӱ_�l���O
                string BaseCurrencyName = BaseCurrencyDropDownList.Text; // �_�l���O�ﶵ�W��

                for (int targetCurrencyIndex = 10; targetCurrencyIndex <= currencyAmount; targetCurrencyIndex++) //default targetCurrencyIndex = 2
                {
                    Tools.ScrollPageUpOrDown(driver, 200); 
                    Thread.Sleep(100);

                    if (baseCurrencyIndex != targetCurrencyIndex) 
                    {
                        IWebElement TargetCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[2]/div/ul/li"));
                        TargetCurrencyDropDownList.Click(); // �I�}�ؼй��O�U�Կ��
                        string targetCurrencyXPath = $"//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[2]/div/ul/li/ul/li[{targetCurrencyIndex}]"; // �ؼй��OXPath
                        driver.FindElement(By.XPath(targetCurrencyXPath)).Click(); // ��ܤ@�ӥؼй��O
                        string targetCurrencyName = TargetCurrencyDropDownList.Text; // �ؼй��O�ﶵ�W��
                        Tools.ScrollPageUpOrDown(driver, 300);
                        Thread.Sleep(100);
                        Tools.PageSnapshot($@"{snapshotpath}\{BaseCurrencyName}_�I��_{targetCurrencyName}_M.png", driver);
                        Thread.Sleep(100);


                        Excel.Range oRange = (Excel.Range)xlWorkSheet.Cells[cell+1, 2];
                        float Left = (float)((double)oRange.Left);
                        float Top = (float)((double)oRange.Top);

                        xlWorkSheet.Cells[cell, 1] = $"{BaseCurrencyName}_�I��_{targetCurrencyName}";
                        xlWorkSheet.Shapes.AddPicture($@"{snapshotpath}\{BaseCurrencyName}_�I��_{targetCurrencyName}_M.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left,Top, 750, 350); 
                        cell +=23;


                    }
                    else //�_�l���O�P�ؼй��O�ۦP��, skip, ���L
                    {
                        continue;
                    }
                }
            }
            xlWorkBook.SaveAs($@"d:\�Y���ײv_�Ҧ����ȶײv����_{time}.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            driver.Quit();
        }
    }
}


