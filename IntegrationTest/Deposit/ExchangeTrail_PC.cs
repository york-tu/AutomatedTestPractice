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
    public class �x������_���׸պ�_PC��:IntegrationTestBase
    {
        public �x������_���׸պ�_PC��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void ���׸պ�_�Ҧ����ȶײv����_����ˬd(string browser)
        {
            StartTestCase(browser, " ���׸պ�_�Ҧ����ȶײv����_����ˬd", "York");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\ExchangeRate\ExchangeTrail";
            Tools.CreateSnapshotFolder(snapshotpath);
            Thread.Sleep(100);
            Tools.ScrollPageUpOrDown(driver, 100);

            driver.Navigate().Refresh();
            driver.FindElement(By.XPath("//*[@id='btnAntiFraud']")).Click(); // ����"�����z"�u������

            if (driver.FindElement(By.XPath("//*[@id='mainform']/div[5]/div/ul")).GetAttribute("style").ToString() == "display: block;")
            {
                driver.FindElement(By.XPath("//*[@id='mainform']/div[5]/div/a")).Click();
            }

            Tools.ScrollPageUpOrDown(driver, 1000);

            driver.FindElement(By.XPath("//*[@id='amount']")).SendKeys("10000"); // ��J�_�l����

            int currencyAmount = driver.FindElements(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[1]/ul/li/ul/li")).Count; // ����_�l���O�U�Կ��ﶵ�ƶq

            for (int baseCurrencyIndex = 2; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> �Ĥ@�ӹ��O
            {
                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[1]/ul"));
                BaseCurrencyDropDownList.Click(); // �I���_�l���O�U�Կ��
                string baseCurrencyXPath = $"//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[1]/ul/li/ul/li[{baseCurrencyIndex}]"; // �_�l���OXPath
 
                if (baseCurrencyIndex == 9 || baseCurrencyIndex == 13 || baseCurrencyIndex == 16) // for �u�ʿ��P�_��
                {
                    Point TargetPosition = new Point(900, 280); //�_�l���O���b�e���WX,Y�y��
                    Cursor.Position = TargetPosition; // �ƹ����в��ʨ�_�l���O�U�Կ���
                    Thread.Sleep(300); 

                    var sim = new InputSimulator();
                    sim.Mouse.VerticalScroll(-1); // �N�ƹ��u���V�U�u�@�� >>> ���slider�V�U�Ԥ@��
                    Thread.Sleep(300);
                }

                driver.FindElement(By.XPath(baseCurrencyXPath)).Click(); // ��ܤ@�Ӱ_�l���O
                string BaseCurrencyName = BaseCurrencyDropDownList.Text; // �_�l���O�ﶵ�W��

                for (int targetCurrencyIndex = 2; targetCurrencyIndex <= currencyAmount; targetCurrencyIndex++) 
                {
                    IWebElement TargetCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[3]/ul/li"));
                    TargetCurrencyDropDownList.Click(); // �I�}�ؼй��O�U�Կ��
                    string targetCurrencyXPath = $"//*[@id='mainform']/div[13]/div[2]/div[2]/div/div[2]/div[1]/div[3]/ul/li/ul/li[{targetCurrencyIndex}]"; // �ؼй��OXPath

                    if (targetCurrencyIndex == 9 || targetCurrencyIndex == 13 || targetCurrencyIndex == 16) // for �u�ʿ��P�_��
                    {
                        Point TargetPosition = new Point(900, 400); //�ؼй��O���b�e���WX,Y�y��
                        Cursor.Position = TargetPosition;  // �ƹ����в��ʨ�ؼй��O�U�Կ���
                        Thread.Sleep(300);

                        var sim = new InputSimulator();
                        sim.Mouse.VerticalScroll(-1); // �N�ƹ��u���V�U�u�@�� >>> ���slider�V�U�Ԥ@��
                        Thread.Sleep(300);
                    }

                    driver.FindElement(By.XPath(targetCurrencyXPath)).Click(); // ��ܤ@�ӥؼй��O
                    string targetCurrencyName = TargetCurrencyDropDownList.Text; // �ؼй��O�ﶵ�W��

                    driver.FindElement(By.XPath("//*[@id='calculate']")).Click(); // �I "�}�l�պ�"

                    //Tools.PageSnapshot($@"{snapshotpath}\{BaseCurrencyName}_�I��_{targetCurrencyName}_���׸պ�.png", driver); // �I��
                    Tools.ElementSnapshotshot(driver.FindElement(By.XPath("//*[@id='mainform']/div[13]/div[2]/div[2]/div")), $@"{snapshotpath}\{BaseCurrencyName}_�I��_{targetCurrencyName}_���׸պ�.png");
                    Thread.Sleep(100);
                }
            }
            driver.Quit();
        }
    }
}


