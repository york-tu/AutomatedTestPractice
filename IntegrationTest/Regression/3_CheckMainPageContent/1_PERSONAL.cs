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
            CreateReport($"�ӤH�A��_����_���e�ˬd", "York");

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
            var driver = new ChromeDriver(chromeService,chromeOptions, TimeSpan.FromSeconds(120));

            //var firefoxOptions = new FirefoxOptions();
            ////firefoxOptions.AddArguments("--headless");
            //var driver = new FirefoxDriver(firefoxOptions);
            #endregion

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal");

            #region Ū�� json
            string path = $"{UserDataList.Upperfolderpath}Settings\\PersonalMainPage.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);
            #endregion

            INFO("�ˬd [�ӤH�A��] �������e �ﶵ�W�� �P �s��");
            var totalURLCounts = jsonArray.Count(); // json��data�ƶq
            for (int i = 0; i < totalURLCounts; i++)
            {
                JObject obj = (JObject)jsonArray[i];

                string targetURL = obj["TargetURL"].ToString();
                string elementCssSelector = obj["PageKeywordCssSelector"].ToString();
                string expectText = obj["ExpectString"].ToString();

                string actualURL = driver.FindElementByCssSelector(elementCssSelector).GetAttribute("href");
                string actualText = driver.FindElementByCssSelector(elementCssSelector).Text;

                #region 1. �ˬd���� '�m�����' ����
                if (i <= 8 )
                {
                    if (i == 0)
                    {
                        INFO(" �ˬd�����m�����");
                    }

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
                #endregion

                #region 2. �ˬd '���D & MegaMenu' ����
                else if (i >=9 && i <= 20)
                {
                    if (i == 9)
                    {
                        INFO("");
                        INFO("�ˬd '���D & MegaMenu' ����");
                    }
                    else if (i == 17)
                    {
                        driver.FindElement(By.CssSelector(".l-hearder__dropDownList")).Click();
                        System.Threading.Thread.Sleep(1000);
                    }

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
                #endregion

                #region 3. �ˬd '���~�P�A��' ����
                else if (i >= 21 && i <= 28)
                {
                    TestBase.ScrollPageUpOrDown(driver, 500);
                    if (i==21)
                    {
                        INFO("");
                        INFO("�ˬd '���~�P�A��' ����");
                    }

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
                #endregion

                #region 4. �ˬd '�~���ײv' ����
                else if (i >= 29 && i <= 34)
                {
                    TestBase.ScrollPageUpOrDown(driver, 1200);
                    if (i == 29)
                    {
                        INFO("");
                        INFO("�ˬd '�~���ײv' ����'");
                    }
                    else if (i==33)
                    {
                        driver.FindElement(By.CssSelector(".color-primary-underline")).Click();
                    }
                    
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
                #endregion

                #region 5. �ˬd '�A���ͬ�����' ����
                else if (i >= 35 && i <= 38)
                {
                    TestBase.ScrollPageUpOrDown(driver, 1900);
                    if (i == 35)
                    {
                        INFO("");
                        INFO("�ˬd '�A���ͬ�����' ����'");
                    }

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
                #endregion

                #region 6. �ˬd '�����Ʀ�A��' ����
                else if (i >= 39 && i <= 42)
                {
                    TestBase.ScrollPageUpOrDown(driver, 2500);
                    if (i == 39)
                    {
                        INFO("");
                        INFO("�ˬd '�����Ʀ�A��' ����'");
                    }

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
                #endregion

                #region 7. �ˬd '�̷s����' ����
                else if (i >= 43 && i <= 44 )
                {
                    TestBase.ScrollPageUpOrDown(driver, 3000);
                    if (i == 43)
                    {
                        INFO("");
                        INFO("�ˬd '�̷s����' ����'");
                    }

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
                #endregion

                #region 8. �ˬd '��h�s�� �P '�m������'
                else if (i >= 45)
                {
                    TestBase.ScrollPageUpOrDown(driver, 4000);
                    if (i == 45)
                    {
                        INFO("");
                        INFO("�ˬd '��h�s��' ����'");
                    }
                    else if (i == 49 )
                    {
                        INFO("�ˬd '����ɤs' ����'");
                    }
                    else if (i == 53 )
                    {
                        INFO("�ˬd '�ɤs�A�Ⱥ�' ����'");
                    }
                    else if (i == 57 )
                    {
                        INFO("�ˬd '�ɤs���s' ����'");
                    }
                    else if (i == 60 )
                    {
                        INFO("�ˬd '�s�ګO�I' ����'");
                    }
                    else if (i == 61)
                    {
                        INFO("�ˬd '�����m��' ����'");
                    }

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
                #endregion
            }
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
