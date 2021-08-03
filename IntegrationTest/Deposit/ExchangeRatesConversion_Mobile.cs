using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;

namespace AutomatedTest.IntegrationTest.Deposit
{
    public class �Y���ײv_�ײv����_M��:IntegrationTestBase
    {
        public �Y���ײv_�ײv����_M��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
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
            Thread.Sleep(100);
            Tools.ScrollPageUpOrDown(driver, 100);

            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/input")).SendKeys("10000"); //��J�_�l����
            int currencyAmount = driver.FindElements(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li")).Count; // ����_�l���O�U�Կ��ﶵ�ƶq
            driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/div[2]/a[2]")).Click(); // �I �}�l�պ�

            for (int baseCurrencyIndex = 2; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> �Ĥ@�ӹ��O
            {
                Tools.ScrollPageUpOrDown(driver, 100);
                Thread.Sleep(100);

                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li"));
                BaseCurrencyDropDownList.Click(); // �I���_�l���O�U�Կ��
                string baseCurrencyXPath = $"//*[@id='mainContent']/div/div[4]/div/div/div[1]/section[1]/div/ul/li/ul/li[{baseCurrencyIndex}]/span"; // �_�l���OXPath
                driver.FindElement(By.XPath(baseCurrencyXPath)).Click(); // ��ܤ@�Ӱ_�l���O
                string BaseCurrencyName = BaseCurrencyDropDownList.Text; // �_�l���O�ﶵ�W��

                for (int targetCurrencyIndex = 2; targetCurrencyIndex <= currencyAmount; targetCurrencyIndex++)
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
                    }
                    else //�_�l���O�P�ؼй��O�ۦP��, skip, ���L
                    {
                        continue;
                    }
                }
            }
            driver.Quit();
        }
    }
}


