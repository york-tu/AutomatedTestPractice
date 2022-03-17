using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;

namespace AutomatedTest.IntegrationTest.Personal.Deposit.Forex
{
    public class �~�׶ײv_���v�ײv����_PC��:IntegrationTestBase
    {
        public �~�׶ײv_���v�ײv����_PC��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/deposit/rate/forex/exchange-rate-chart"; // �n�JPC�� "���v�ײv����" ����
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void ���O�d��(string browser)
        {
            StartTestCase(browser, "���v�ײv����_���O�d��", "York");
            INFO("");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\Forex";
            TestBase.CreateFolder(snapshotpath);
            TestBase.CleanUPFolder(snapshotpath);

            int dropDownListCount = driver.FindElements(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li/ul/li")).Count; // �U�Կ��ﶵ�ƶq
            for (int index = 2; index <= dropDownListCount; index++) // inxde = 2 = �Ĥ@���i����O
            {
                TestBase.ScrollPageUpOrDown(driver, 450);
                Thread.Sleep(300);
                driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li")).Click(); // �I�}���O�U�Կ��
                string currencyXPath = $"//*[@id='tab-01']/div[1]/div[1]/ul/li/ul/li[{index}]";
                driver.FindElement(By.XPath(currencyXPath)).Click(); // ��ܤ@�ӹ��O

                IWebElement StartDate = driver.FindElement(By.XPath("//*[@id='datepicker1']"));
                StartDate.Clear();
                StartDate.SendKeys("2021-07-01");
                Thread.Sleep(300);
                StartDate.Click(); // (�קK��J����ɼu�X�����p�����v�T���U�ﶵ�I��, �I���T�O�p����������)
                
                for (int i = 1; i <= 2 ; i++) // i=1=�{��radiobutton, i=2=�Y��radiobutton
                {
                    TestBase.ScrollPageUpOrDown(driver, 450);
                    Thread.Sleep(300);

                    string CurrencyName = driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li")).Text;
                    
                    // �]�H�U���O�����Ѳ{���ﶵ >>> �u�d�� "�Y��"
                    if (CurrencyName == "�n�D��(ZAR)" || CurrencyName == "�æ�����(NZD)" || CurrencyName == "������ܯ�(MXN)" || CurrencyName == "��h�k��(CHF)" || CurrencyName == "����(SEK)" || CurrencyName == "�s�[�Y��(SGD)" || CurrencyName == "����(THB)")
                    {
                        i = 2;
                    }

                    driver.FindElement(By.XPath($"//*[@id='tab-01']/div[1]/div[1]/div/label[{i}]")).Click(); // �I "�{��/�Y��"

                    for (int j = 1; j <= 11; j++) // �s��������չ�/�ƾڪ� (���F�קK�I���ﶵ��S�����T����, �����h���H�T�O�T������ܼƾڪ�)
                    {
                        driver.FindElement(By.XPath($"//*[@id='tab-01']/div[1]/div[3]/div/label[{j % 2 + 1}]")).Click(); // �I "���չ�/�ƾڪ�"
                        Thread.Sleep(100);
                        driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[4]/a")).Click(); // �I "�}�l�d��"
                    }
                    //*[@id="tab-01"]/div[1]/div[1]/ul/li/ul/li[5]/span
                    int pagelist = driver.FindElements(By.XPath("//*[@id='pagelist']/ul/li")).Count; // ��������ƶq
                    for (int list = 1; list <= pagelist; list++)
                    {
                        TestBase.ScrollPageUpOrDown(driver, 900);
                        Thread.Sleep(300);
                        driver.FindElement(By.XPath($"//*[@id='pagelist']/ul/li[{list}]/a")).Click(); // �I����
                        Thread.Sleep(100);

                        IWebElement CurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[1]/ul/li")); // ���O�U�Կ��
                        IWebElement CashSort = driver.FindElement(By.XPath($"//*[@id='tab-01']/div[1]/div[1]/div/label[{i}]")); // �{��/�Y��
                        TestBase.PageSnapshot(driver,$@"{snapshotpath}\{CurrencyDropDownList.Text}_{CashSort.Text}_page{list}.png");
                        
                        Thread.Sleep(100);
                    }
                }
            }
            driver.Quit();
        }
    }
}
