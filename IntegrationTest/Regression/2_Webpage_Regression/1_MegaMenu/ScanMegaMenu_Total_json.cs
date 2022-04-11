using Xunit;
using Excel = Microsoft.Office.Interop.Excel;
using AutomatedTest.Utilities;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System.Linq;
using OpenQA.Selenium.Interactions;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class B_a_CheckMegaMenu_Total:IntegrationTestBase
    {
        public B_a_CheckMegaMenu_Total(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Theory]
        /* MegaMenu���
         * 1: �s��
         * 2: �U��
         * 3: �H�Υd/��I
         * 4: �]�I�޲z
         * 5: �H�U
         * 6: �O�I
         * 7: �ͬ�����
         */
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        [InlineData(5)]
        [InlineData(6)]
        [InlineData(7)]
        public void �ˬdTotalMegaMenu_excel(int megaMenuIndex)
        {
            #region step 1. switch����
            string testCaseName = "";
            int row = 0, line = 0; // �]�w megamenu ��� "�� & �C" ��
            switch(megaMenuIndex)
            {
                case 1:
                    testCaseName = "�s��";
                    row = 6; line = 10;
                    break;
                case 2:
                    testCaseName = "�U��";
                    row = 5; line = 10;
                    break;
                case 3:
                    testCaseName = "�H�Υd_��I";
                    row = 6; line = 10;
                    break;
                case 4:
                    testCaseName = "�]�I�޲z";
                    row = 6; line = 10;
                    break;
                case 5:
                    testCaseName = "�H�U";
                    row = 5; line = 8;
                    break;
                case 6:
                    testCaseName = "�O�I";
                    row = 4; line = 8;
                    break;
                case 7:
                    testCaseName = "�ͬ�����";
                    row = 1; line = 8;
                    break;
                default:
                    break;
            }
            #endregion

            CreateReport($"[{testCaseName}] MegaMenu ���e�ˬd", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");

            #region  step 2: setup browser ���}�Һ���
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

            driver.Navigate().GoToUrl("https://www.esunbank.com.tw/bank/personal");

            #region step 3: Ūjson���
            string path = $"{UserDataList.Upperfolderpath}Settings\\URL_Css_MegaMenu_ExpectString.json";
            var jsonContent = File.ReadAllText(path);
            JArray jsonArray = JArray.Parse(jsonContent);
            #endregion

            Actions act = new Actions(driver);
            if (megaMenuIndex == 7 ) // ������� "7:�ͬ�����" ��, �������P�[�cCssSelector
            {
                act.MoveToElement(driver.FindElement(By.CssSelector(".l-megaL1__include > a:nth-child(1)"))).Perform();
            }
            else
            {
                // Cursor���� MegaMenu [testCaseName] ���
                act.MoveToElement(driver.FindElement(By.CssSelector($"ul.l-megaL1__list:nth-child(1) > li:nth-child({megaMenuIndex}) > div:nth-child(1)"))).Perform(); 
            }

            // �`����megamenu�ﶵ (i=��, j=�C)
            for (int i = 1; i<=row ; i++) 
            {
                for (int j = 0; j<=line ; j++) 
                {
                    var cssSelector = "";

                    if (megaMenuIndex == 7) // ��ﶵ�� "7:�ͬ�����" ��, �������P�[�cCssSelector
                    {
                        cssSelector = $".l-megaL2__noInclude > div:nth-child(1) > ul:nth-child(1) > li:nth-child({j}) > a:nth-child(1)";
                    }
                    else if (megaMenuIndex < 7 && j == 0) // �ﶵ���� "7:�ͬ�����" & �w�� megamenu ���h�ʶ���r�j�� (j=0)
                    {
                        cssSelector = $"ul.l-megaL1__list:nth-child(1) > li:nth-child({megaMenuIndex}) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child({i}) > a:nth-child(1)";
                    }
                    else
                    {
                        cssSelector = $"ul.l-megaL1__list:nth-child(1) > li:nth-child({megaMenuIndex}) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child({i}) > ul:nth-child(2) > li:nth-child({j}) > a:nth-child(1)";
                    }

                    try
                    {
                        string actualOptionName = driver.FindElement(By.CssSelector(cssSelector)).Text; // ����MegaMenu�W�ﶵ�W��
                        string actURL = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("href"); // ����MegaMenu�W�ﶵ�����W�s��

                        IEnumerable<JToken> tokens = jsonArray.SelectTokens($"[?(@.MegaMenuOptions == '{actualOptionName}')]",true); // �j�M�ﶵ���� json ��Ʀ�m
                        if (tokens.Any() == false) // falas >>> json ���j�M����
                        {
                            FAIL($"[Mega Menu][{testCaseName}] �ﶵ\"{actualOptionName}\''���b�w���ﶵ��");
                            continue;
                        }
                        else
                        {
                            JToken targetToken = i <=3 ? targetToken = tokens.First() : targetToken = tokens.Last(); // for�P�ﶵ�W��, �����P���e
                            var targetArrayIndex = Convert.ToInt32((targetToken.Path).Substring(1, (targetToken.Path).Length - 2));
                            var expectURL = targetToken["TargetURL"].ToString();

                            if (actURL == expectURL)
                            {
                                PASS($"[Mega Menu][{testCaseName}] �ﶵ \" {actualOptionName}\" �W�s�����}  \"{actURL}\" �ŦX�w��");
                            }
                            else if (actURL != expectURL)
                            {
                                FAIL($"[Mega Menu][{testCaseName}] �ﶵ \"{actualOptionName}\" �W�s�����} \"{actURL}\" (�w�����} \"{expectURL}\")");
                                // FAIL(TestBase.PageSnapshotToReport(driver));
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        //WARNING(e.Message);
                    }
                }
            }
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
