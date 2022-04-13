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
using System.Threading;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class D_c_CheckMainPage_CreditCard:IntegrationTestBase
    {
        public D_c_CheckMainPage_CreditCard(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Fact]
        public void �ˬd�H�Υd��I_�������e()
        {
            string path = $"{UserDataList.Upperfolderpath}Settings\\CreditCardMainPage.json"; // json�ɸ��|
            #region Ū json ��ƻy�k
            var jsonContent = File.ReadAllText(path);
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
            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal/credit-card");

            CreateReport($"�H�Υd��I_����_���e�ˬd", "York");
            INFO("�ˬd [�H�Υd/��I] �������e �ﶵ�W�� �P �s��");
            INFO("");
            
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
                        // 3. �ˬd '�`�ΪA��' ����
                        TestBase.ScrollPageUpOrDown(driver, 500);
                        INFO("�ˬd '�`�ΪA��' ����");
                        break;
                    case 26:
                        // 4. �ˬd 'ú�|�u�ٿ����D' ����
                        TestBase.ScrollPageUpOrDown(driver, 800);
                        INFO("�ˬd 'ú�|�u�ٿ����D' ����'");
                        break;
                    case 27:
                        // 5. �ˬd '�Ʀ�A��' ����
                        INFO("�ˬd '�Ʀ�A��' ����'");
                        break;
                    case 35:
                        // 6. �ˬd '��������' ����
                        TestBase.ScrollPageUpOrDown(driver, 2100);
                        INFO("�ˬd ''��������' ����'");
                        break;
                    case 38:
                        // 7. �ˬd '�ӤH�A��' ����
                        TestBase.ScrollPageUpOrDown(driver, 2500);
                        INFO("�ˬd '�ӤH�A��' ����'");
                        break;
                    case 45:
                        // 8. �ˬd '�����s��' ����
                        TestBase.ScrollPageUpOrDown(driver, 2800);
                        INFO("�ˬd '�����s��' ����'");
                        INFO("�͵��M��");
                        break;
                    case 48:
                        INFO("���i�P��������");
                        break;
                    case 52:
                        INFO("���a�A��");
                        break;
                    case 53:
                        // 9. �ˬd '�`�����D���pô�ȪA' ����
                        TestBase.ScrollPageUpOrDown(driver, 3200);
                        INFO("�`�����D");
                        break;
                    case 58:
                        INFO("�pô�ȪA");
                        break;
                    case 62:
                        // 10. �ˬd '�m������'
                        TestBase.ScrollPageUpOrDown(driver, 4000);
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
        public void �ˬd�H�Υd��I_����s���ɤޭ�()
        {
            #region KillProcess
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");
            #endregion
            #region step 1: Ūjson���
            string path = $"{UserDataList.Upperfolderpath}Settings\\CreditCardMainPage.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion
            var totalURLCounts = jsonArray.Count(); // json��data�ƶq
            CreateReport($"�H�Υd��I_����s���ɤޭ�", "York");

            #region step 2:  browser ���}�Һ����]�w 
            //Chrome headless �ѼƳ]�w
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
            //�ظm Chrome Driver
            var driver = new ChromeDriver(chromeService, chromeOptions, TimeSpan.FromSeconds(300));

            //var firefoxOptions = new FirefoxOptions();
            ////firefoxOptions.AddArguments("--headless");
            //var driver = new FirefoxDriver(firefoxOptions);
            #endregion

            for (int i = 0; i < totalURLCounts; i++) 
            {
                JObject obj = (JObject)jsonArray[i];
                string uRL = obj["TargetURL"].ToString();
                string cssSelector = obj["DirectPageElementCssSelector"].ToString();
                string expectString = obj["DirectPageKeyword"].ToString();

                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // �t�}�s��
                driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on �s���W 
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600);
                driver.Navigate().GoToUrl(uRL);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(600);
                
                #region ������۰ʤ���M����, click "�����q����" �j����^PC��
                if (driver.Url.ToString().Contains("?dev=mobile")) // workaround: ������۰ʤ���M����, �j����^PC��
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

                if (driver.Url.ToString() != uRL) 
                {
                    WARNING($"[Page Redirect], {driver.Url.ToString()}  (Expect: {uRL })");
                    WARNING(TestBase.PageSnapshotToReport(driver));
                    continue;
                }
                else if ( uRL == "https://gib.esunbank.com/" || uRL == "https://netbank.esunbank.com.tw/webatm/#/login" || uRL == " https://www.esunbank.com.tw/bank/-/media/esunbank/files/credit-card/2017card_ex.pdf?la=en")
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
