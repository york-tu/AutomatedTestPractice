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
    public class D_b_CheckMainPage_Deposit:IntegrationTestBase
    {
        public D_b_CheckMainPage_Deposit(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Fact]
        public void �ˬd�s��_����()
        {
            CreateReport($"�s��_����_���e�ˬd", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");

            #region  Browser ���}�Һ����]�w
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

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal/deposit");

            #region Ū�� json
            string path = $"{UserDataList.Upperfolderpath}Settings\\DepositMainPage.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion

            INFO("�ˬd [�s��] �������e �ﶵ�W�� �P �s��");
            INFO("");
            var totalURLCounts = jsonArray.Count(); // json��data�ƶq
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
                        // 4. �ˬd '�u�W�}�ߥɤs�Ʀ�b��' ����
                        TestBase.ScrollPageUpOrDown(driver, 800);
                        INFO("");
                        INFO("�ˬd '�~���ײv' ����'");
                        break;
                    case 30:
                        // 5. �ˬd '�~���u�f' ����
                        TestBase.ScrollPageUpOrDown(driver, 800);
                        INFO("");
                        INFO("�ˬd '�~���u�f' ����'");
                        break;
                    case 35:
                        // 6. �ˬd '��������' ����
                        TestBase.ScrollPageUpOrDown(driver, 2100);
                        INFO("");
                        INFO("�ˬd ''��������' ����'");
                        break;
                    case 38:
                        // 7.�ˬd '�U���A�ȥӽ�' ����
                        TestBase.ScrollPageUpOrDown(driver, 2500);
                        INFO("");
                        INFO("�ˬd '�U���A�ȥӽ�' ����'");
                        break;
                    case 42:
                        TestBase.ScrollPageUpOrDown(driver, 2800);
                        INFO("");
                        INFO("�ˬd '���O�зǤΤ��i����' ����'");
                        break;
                    case 50:
                        // 8. �ˬd '���O�зǤΤ��i����' ����
                        TestBase.ScrollPageUpOrDown(driver, 3200);
                        INFO("");
                        INFO("�ˬd '�`�����D' ����'");
                        break;
                    case 55:
                        // 9. �ˬd '�`�����D' �P '�pô�ȪA' ����
                        INFO("");
                        INFO("�ˬd '�pô�ȪA' ����'");
                        break;
                    case 59:
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
     }
}
