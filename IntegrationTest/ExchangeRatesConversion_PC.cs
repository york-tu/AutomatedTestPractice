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
    public class �����k���`�n�ֱ����__�պ�u��_PC��:IntegrationTestBase
    {
        public �����k���`�n�ֱ����__�պ�u��_PC��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void �`�n�ֱ����_�Ҧ����ȶײv����_����ˬd(string browser)
        {
            StartTestCase(browser, " �`�n�ֱ����_�Ҧ����ȶײv����_����ˬd", "York");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\ExchangeRate";
            TestBase.CreateFolder(snapshotpath);
            Thread.Sleep(100);
            TestBase.ScrollPageUpOrDown(driver, 100);

            driver.Navigate().Refresh(); // refresh���� >>> �������k���ֱ�������
            driver.FindElement(By.XPath("//*[@id='btnAntiFraud']")).Click(); // ����"�����z"�u������

            Actions act = new Actions(driver);
            act.MoveToElement(driver.FindElement(By.XPath("//*[@id='mainform']/div[5]/div/ul/li[2]"))).Perform(); // �ƹ�"����" �k���`�n "�պ�u��"���

            driver.SwitchTo().Frame(1); // ���� "�պ�u��"��� frame
            driver.FindElement(By.XPath("//*[@id='amount']")).SendKeys("10000"); // ��J�_�l����

            Point CursorHover = new Point(1550, 500);
            Cursor.Position = CursorHover; // �ƹ����а��b�պ��歶���W (�קK��Sfocus�b���ɿ��۰�hide)

            int currencyAmount = driver.FindElements(By.XPath("//*[@id='tabs2-01']/div/div[1]/div[1]/ul/li/ul/li")).Count; // ����_�l���O�U�Կ��ﶵ�ƶq

            for (int baseCurrencyIndex = 2; baseCurrencyIndex <= currencyAmount; baseCurrencyIndex++) // currencyindex =2 >>> �Ĥ@�ӹ��O
            {
                IWebElement BaseCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='tabs2-01']/div/div[1]/div[1]/ul/li"));
                BaseCurrencyDropDownList.Click(); // �I���_�l���O�U�Կ��
                string baseCurrencyXPath = $"//*[@id='tabs2-01']/div/div[1]/div[1]/ul/li/ul/li[{baseCurrencyIndex}]/span"; // �_�l���OXPath

                if (baseCurrencyIndex == 5 || baseCurrencyIndex == 8 || baseCurrencyIndex == 12 || baseCurrencyIndex == 15) // for �u�ʿ��P�_��
                {
                    Point TargetPosition = new Point(1375, 550); //�_�l���O���b�e���WX,Y�y��
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
                    IWebElement TargetCurrencyDropDownList = driver.FindElement(By.XPath("//*[@id='tabs2-01']/div/div[1]/div[3]/ul/li"));
                    TargetCurrencyDropDownList.Click(); // �I�}�ؼй��O�U�Կ��
                    string targetCurrencyXPath = $" //*[@id='tabs2-01']/div/div[1]/div[3]/ul/li/ul/li[{targetCurrencyIndex}]/span"; // �ؼй��OXPath

                    if (targetCurrencyIndex == 5 || targetCurrencyIndex == 8 || targetCurrencyIndex == 12 || targetCurrencyIndex == 15) // for �u�ʿ��P�_��
                    {
                        Point TargetPosition = new Point(1375, 650); //�ؼй��O���b�e���WX,Y�y��
                        Cursor.Position = TargetPosition;  // �ƹ����в��ʨ�ؼй��O����
                        Thread.Sleep(300);

                        var sim = new InputSimulator();
                        sim.Mouse.VerticalScroll(-1); // �N�ƹ��u���V�U�u�@�� >>> ���slider�V�U�Ԥ@��
                        Thread.Sleep(300);
                    }

                    driver.FindElement(By.XPath(targetCurrencyXPath)).Click(); // ��ܤ@�ӥؼй��O
                    string targetCurrencyName = TargetCurrencyDropDownList.Text; // �ؼй��O�ﶵ�W��

                    driver.FindElement(By.XPath("//*[@id='calculate']")).Click(); // �I "�}�l�պ�"

                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\{BaseCurrencyName}_�I��_{targetCurrencyName}_PC.png"); // �I��
                    Thread.Sleep(100);
                }
            }
            driver.Quit();
        }
    }
}


