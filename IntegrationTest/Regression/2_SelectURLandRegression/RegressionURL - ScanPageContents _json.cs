using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using AutomatedTest.Utilities;
using System.Net;
using Xunit.Abstractions;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Linq;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class II_I_�ˬd�������e_json:IntegrationTestBase
    {
        public II_I_�ˬd�������e_json(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }
        [Theory]
        [InlineData(0, 200)]
        [InlineData(200, 400)]
        [InlineData(400, 600)]
        [InlineData(600, 750)]
        public void �ˬd�������e�O�_�ŦX�w��(int startIndex, int endIndex)
        {
            #region step 0: kill cache driver
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");
            #endregion

            #region step 1: Ūjson���
            string path = $"{UserDataList.Upperfolderpath}Settings\\ExcelToJson.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion

            var totalURLCounts = jsonArray.Count(); // json��data�ƶq
            endIndex = endIndex >= 701 ? endIndex = totalURLCounts : endIndex;
            CreateReport($"�������e�ˬd_json_{startIndex}-{endIndex}", "York");

            #region step 2:  browser ���}�Һ����]�w 
            //Chrome headless �ѼƳ]�w
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("--headless");
            chromeOptions.AddArguments("--disable-gpu");
            chromeOptions.AddArguments("--incognito");
            chromeOptions.AddArguments("--window-size=1920x1080");
            chromeOptions.AddArguments("--ignore-certificate-errors");
            chromeOptions.AddArguments("--allow-running-insecure-content");
            chromeOptions.AddArguments("--disable-extensions");
            chromeOptions.AddArguments("--proxy-server='direct://'");
            chromeOptions.AddArguments("--proxy-bypass-list=*");
            chromeOptions.AddArguments("--start-maximized");
            chromeOptions.AddArguments("--disable-dev-shm-usage");
            chromeOptions.AddArguments("--no-sandbox");
            chromeOptions.AddArguments("--disable-blink-features=AutomationControlled");
            chromeOptions.AddArguments("--disable-infobars");
            //�ظm Chrome Driver
            var driver = new ChromeDriver(chromeOptions);

            //var firefoxOptions = new FirefoxOptions();
            ////firefoxOptions.AddArguments("--headless");
            //var driver = new FirefoxDriver(firefoxOptions);
            #endregion

            for (int i = startIndex; i < endIndex; i++)
            {
                JObject obj = (JObject)jsonArray[i];
                string uRL = obj["TargetURL"].ToString();
                string cssSelector = obj["PageKeywordCssSelector"].ToString();
                string expectString = obj["ExpectString"].ToString();

                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // �t�}�s��
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on �s���W 
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600�������������e, �_�h����, ���������i�U�@�B.
                driver.Url = uRL;
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600�������������e, �_�h����, ���������i�U�@�B.

                #region ������۰ʤ���M����, click "�����q����" �j����^PC��
                if (driver.Url.ToString().Contains("?dev=mobile")) // workaround: ������۰ʤ���M����, �j����^PC��
                {
                    TestBase.ScrollPageUpOrDown(driver, 1500);
                    driver.FindElementByClassName("changeTarget").Click();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600�������������e, �_�h����, ���������i�U�@�B.
                }
                #endregion

                if (uRL == "https://easyfee.esunbank.com.tw/index.action" || uRL == "https://www.esunbank.com.tw/bank/marketing/loan/marketing-menu-level-2" || uRL == "https://ebank.esunbank.com.tw/index.jsp#toTaskId=FIM01007")
                {
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600); //600�������������e, �_�h����, ���������i�U�@�B.
                    System.Threading.Thread.Sleep(3000);
                    WARNING($"{uRL }, Keyword: {expectString}");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }
                else if (driver.Url.ToString() != uRL) // �ˬd�����}�ҷ�U���}�O�_����J���} (�P�_�����O�_��redirect)
                {
                    WARNING($"[Page Redirect], {driver.Url.ToString()}  (Expect: {uRL })");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }
                else if (driver.Url.ToString() == "https://www.esunbank.com.tw/bank/personal/event/calendar/events") // Workaround 1 : ���ʤ��URL >>> �ݧ��Ѥ��
                {
                    string day_of_week = DateTime.Now.ToString("dddd", new CultureInfo("en-us")); // �^��P���X(e.g., Wednesday)
                    string day = DateTime.Now.ToString("%d"); // ��� (e.g., 1)
                    string month = DateTime.Now.ToString("MMMM", new CultureInfo("zh-cn")); // ������ (e.g., �T��)
                    expectString = $"{day}\r\n{day_of_week}\r\n{month}";
                }
                else if (driver.Url.ToString().EndsWith("/pages")) // Workaround 2 : ���}�� ".../page" >>> �������Ť��e
                {
                    expectString = "";
                }

                try
                {
                    Assert.Contains(expectString, driver.FindElement(By.CssSelector(cssSelector)).Text); // �P�_element �r��O�_�ŦX�w��
                    PASS($"{uRL}, Keyword: {expectString}");
                    PASS(TestBase.PageSnapshotToReport(driver));
                }
                catch (Exception e)
                {
                    FAIL($"{uRL}, Exception:{e.Message}");
                    FAIL(TestBase.PageSnapshotToReport(driver));
                }
                driver.SwitchTo().Window(driver.WindowHandles.Last()).Close(); // �����s��
                driver.SwitchTo().Window(driver.WindowHandles.First()); // ���^�쭶
            }

            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
