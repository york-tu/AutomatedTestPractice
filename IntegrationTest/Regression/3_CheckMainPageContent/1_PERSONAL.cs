using Xunit;
using AutomatedTest.Utilities;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Linq;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class D_a_CheckMainPage_PERSONAL:IntegrationTestBase
    {
        public D_a_CheckMainPage_PERSONAL(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Fact]
        public void �ˬd�ӤH�A��_����()
        {
            string path = $"{UserDataList.Upperfolderpath}Settings\\PersonalMainPage.json"; // json�ɸ��|
            #region Ū json ��ƻy�k
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion
            var totalURLCounts = jsonArray.Count(); // json��data�ƶq

            #region  Chrome�s�������}�Һ����]�w(headless)

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

            var driver = new ChromeDriver(chromeService,chromeOptions, TimeSpan.FromSeconds(120));
            #endregion
            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal");

            CreateReport($"�ӤH�A��_����_���e�ˬd", "York");
            INFO("�ˬd [�ӤH�A��] �������e �ﶵ�W�� �P �s��");
            
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
                        INFO("�ˬd '���D & MegaMenu' ����");
                        break;
                    case 17:
                        driver.FindElement(By.CssSelector(".l-hearder__dropDownList")).Click();
                        System.Threading.Thread.Sleep(1000);
                        break;
                    case 21:
                        // 3. �ˬd '���~�P�A��' ����
                        TestBase.ScrollPageUpOrDown(driver, 500);
                        INFO("�ˬd '���~�P�A��' ����");
                        break;
                    case 29:
                        // 4. �ˬd '�~���ײv' ����
                        TestBase.ScrollPageUpOrDown(driver, 1200);
                        INFO("�ˬd '�~���ײv' ����'");
                        break;
                    case 33:
                        driver.FindElement(By.CssSelector(".color-primary-underline")).Click();
                        break;
                    case 35:
                        // 5. �ˬd '�A���ͬ�����' ����
                        TestBase.ScrollPageUpOrDown(driver, 1900);
                        INFO("�ˬd '�A���ͬ�����' ����'");
                        break;
                    case 39:
                        // 6. �ˬd '�����Ʀ�A��' ����
                        TestBase.ScrollPageUpOrDown(driver, 2500);
                        INFO("�ˬd '�����Ʀ�A��' ����'");
                        break;
                    case 43:
                        // 7. �ˬd '�̷s����' ����
                        TestBase.ScrollPageUpOrDown(driver, 3000);
                        INFO("�ˬd '�̷s����' ����'");
                        break;
                    case 45:
                        // 8. �ˬd '��h�s��' ����
                        TestBase.ScrollPageUpOrDown(driver, 4000);
                        INFO("�ˬd '��h�s��' ����");
                        break;
                    case 49:
                        // 9. �ˬd '�����m��' ����
                        INFO("�ˬd '�����m��' ����");
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
        public void �ˬd�ӤH�A��_����s���ɤޭ�()
        {
            string path = $"{UserDataList.Upperfolderpath}Settings\\PersonalMainPage.json"; // json�ɸ��|
            #region Ūjson��ƻy�k
            var jsonContent = File.ReadAllText(path);
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

            CreateReport($"�ӤH�A��_����s���ɤޭ�", "York");

            for (int i = 0; i < totalURLCounts; i++)
            {
                JObject obj = (JObject)jsonArray[i];
                string uRL = obj["TargetURL"].ToString();
                string cssSelector = obj["DirectPageElementCssSelector"].ToString();
                string expectString = obj["DirectPageKeyword"].ToString();

                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // �s�����t�}�s��
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on �s���W 

                driver.Navigate().GoToUrl(uRL);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600);

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
                if (uRL == "https://www.esunbank.com.tw/event/credit/1040408web/index.htm" || uRL == "https://www.esunbank.com.tw/event/credit/1100412home_al/index.html" || uRL == "https://accessible.esunbank.com.tw/Accessibility/Index")
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
