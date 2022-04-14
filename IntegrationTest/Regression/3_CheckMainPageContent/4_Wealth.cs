using Xunit;
using AutomatedTest.Utilities;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class D_b_CheckMainPage_Wealth:IntegrationTestBase
    {
        public D_b_CheckMainPage_Wealth(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }
        string testCaseName = "�]�I�޲z";
        string mainPageURL = "https://www.esunbank.com.tw/bank/personal/wealth";
        string jsonPath = $@"{UserDataList.Upperfolderpath}Settings\wealthMainPage.json"; // json�ɸ��|

        [Fact]
        public void �ˬd����()
        {
            #region Ū json ��ƻy�k
            var jsonContent = File.ReadAllText(jsonPath);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion
            var totalURLCounts = jsonArray.Count(); // json�̸�Ƽ�

            #region  Chrome�s�������}�Һ����]�w(headless)
            //Chrome headless �ѼƳ]�w
            var chromeService = ChromeDriverService.CreateDefaultService();
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
            var driver = new ChromeDriver(chromeService, chromeOptions, TimeSpan.FromSeconds(120));
            #endregion
            driver.Navigate().GoToUrl(mainPageURL);

            CreateReport($"{testCaseName}_�ˬd�������e�P�s��", "York");

            for (int i = 0; i < totalURLCounts; i++)
            {
                JObject obj = (JObject)jsonArray[i];

                string targetURL = obj["TargetURL"].ToString();
                string elementCssSelector = obj["PageKeywordCssSelector"].ToString();
                string expectText = obj["ExpectString"].ToString();

                string actualURL = driver.FindElementByCssSelector(elementCssSelector).GetAttribute("href");
                string actualText = driver.FindElementByCssSelector(elementCssSelector).Text;

                #region �s�W���D
                switch (i)
                {
                    case 0:
                        // 1. �ˬd���� '�m�����' ����
                        INFO(" �ˬd�����m�����");
                        break;
                    case 9:
                        // 2. �ˬd '���D & MegaMenu' ����
                        INFO("");
                        INFO("�ˬd '���D & MegaMenu' ����");
                        break;
                    case 17:
                        driver.FindElement(By.CssSelector(".l-hearder__dropDownList")).Click();
                        System.Threading.Thread.Sleep(1000);
                        break;
                    case 21:
                        // 3. �ˬd '�`�ΪA��' ����
                        TestBase.ScrollPageUpOrDown(driver, 500);
                        INFO("");
                        INFO("�ˬd '�`�ΪA��' ����");
                        break;
                    case 29:
                        // 4. �ˬd '���e����' ����
                        TestBase.ScrollPageUpOrDown(driver, 800);
                        INFO("");
                        INFO("�ˬd '���e����' ����'");
                        break;
                    case 30:
                        // 5. �ˬd '�Y�ɬd��' ����
                        TestBase.ScrollPageUpOrDown(driver, 1500);
                        INFO("");
                        INFO("�ˬd '�Y�ɬd��' ����'");
                        break;
                    case 35:
                        // 6. �ˬd '��������' ����
                        TestBase.ScrollPageUpOrDown(driver, 2000);
                        INFO("");
                        INFO("�ˬd ''��������' ����'");
                        break;
                    case 38:
                        // 7.�ˬd '�z�]�q�l�g��' ����
                        TestBase.ScrollPageUpOrDown(driver, 2800);
                        INFO("");
                        INFO("�ˬd '�z�]�q�l�g��' ����'");
                        break;
                    case 39:
                        // 8. �ˬd '������T' ����
                        TestBase.ScrollPageUpOrDown(driver, 3500);
                        INFO("");
                        INFO("�ˬd '������T' ����'");
                        break;
                    case 43:
                        // 9. �ˬd '�`�����D' �P '�pô�ȪA' ����
                        INFO("");
                        INFO("�ˬd '�`�����D & �pô�ȪA' ����'");
                        break;
                    case 52:
                        // 10. �ˬd '�m������'
                        TestBase.ScrollPageUpOrDown(driver, 4000);
                        INFO("");
                        INFO("�ˬd '�����m��' ����'");
                        break;
                    default:
                        break;
                }
                #endregion

                if (actualURL != targetURL)
                {
                    FAIL($"[�s�����~] �w��: {targetURL}, ���: {actualURL}");
                }
                else if (actualText != expectText)
                {
                    FAIL($"[���D���~] �w��: {expectText}, ���: {actualText}");
                }
                else
                {
                    PASS($"[�s���P���D���T] �s��: {actualURL}, ���D: {actualText}");
                }
            }
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }

        [Fact]
        public void �ˬd����s���ɤޭ�()
        {
            #region Ūjson��ƻy�k
            var jsonContent = File.ReadAllText(jsonPath);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion
            int totalURLCounts = jsonArray.Count(); // json�̸�Ƽ�

            #region Chrome�s�������}�Һ����]�w(headless) 

            var chromeService = ChromeDriverService.CreateDefaultService();
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments("--headless");
            chromeOptions.AddArguments("--disable-gpu");
            chromeOptions.AddArguments("--incognito");
            chromeOptions.AddArguments("--window-size=1440x900");
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

            var driver = new ChromeDriver(chromeService, chromeOptions, TimeSpan.FromSeconds(300));
            #endregion

            CreateReport($"{testCaseName}_�ˬd����s���ɤޭ�", "York");

            for (int i = 28; i < totalURLCounts; i++)
            {
                JObject obj = (JObject)jsonArray[i];
                string uRL = obj["TargetURL"].ToString();
                string cssSelector = obj["DirectPageElementCssSelector"].ToString();
                string expectString = obj["DirectPageKeyword"].ToString();

                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // �s�����t�}�s��
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on �s���W 
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600);
                driver.Navigate().GoToUrl(uRL);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600);

                var bb = driver.Url.ToString();
                #region ������۰ʤ���M����, click "�����q����" �j����^PC��
                if (driver.Url.ToString().Contains("?dev=mobile"))
                {
                    TestBase.ScrollPageUpOrDown(driver, 5000);
                    if (driver.Url.ToString().Contains("www.esunfhc.com"))
                    {
                        driver.FindElementById("fhc_layout_m_0_fhc_maincontent_m_2_HlkToWeb").Click();
                    }
                    else
                    {
                        driver.FindElementByClassName("changeTarget").Click();
                    }
                }
                #endregion

                if (driver.Url.ToString() != uRL) // �����}�ҷ�U���}�P�w������ >>> ����redirect
                {
                    WARNING($"[Page Redirect], {driver.Url.ToString()}  (Expect: {uRL })");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }
                // Wordaround ��������H�U���} >>> �H�u check
                else if (uRL == "https://gib.esunbank.com/" || uRL == "https://netbank.esunbank.com.tw/webatm/#/login" || uRL == " https://www.esunbank.com.tw/bank/-/media/esunbank/files/credit-card/2017card_ex.pdf?la=en")
                {
                    WARNING($"[NeedManualCheck], {driver.Url.ToString()}  (Expect: {uRL })");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }

                string actualText = "";
                if (uRL == "https://accessible.esunbank.com.tw/Accessibility/Index")
                {
                    actualText = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("alt");
                }
                else if (uRL == "https://ebank.esunbank.com.tw/index.jsp")
                {
                    driver.SwitchTo().Frame("iframe1");
                    actualText = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("title");
                }
                else
                {
                    actualText = driver.FindElement(By.CssSelector(cssSelector)).Text;
                }

                try
                {
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600);
                    Assert.Contains(expectString, actualText); // �P�_element �r��O�_�ŦX�w��
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
