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
using OpenQA.Selenium.Interactions;

namespace AutomatedTest.IntegrationTest.Regression
{
    public class II_III_�s�פ���_MegaMenu�ˬd:IntegrationTestBase
    {
        public II_III_�s�פ���_MegaMenu�ˬd(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Fact]
        #region �u�@��������}����
        /*
        �u�@�� 1:      "/bank/personal",
        �u�@�� 2:      "/bank/personal/deposit",
        �u�@�� 3:      "/bank/personal/loan", 
        �u�@�� 4:      "/bank/personal/credit-card",
        �u�@�� 5:      "/bank/personal/wealth", 
        �u�@�� 6:      "/bank/personal/trust", 
        �u�@�� 7:      "/bank/personal/insurance",
        �u�@�� 8:      "/bank/personal/lifefin", 
        �u�@�� 9:      "/bank/personal/apply", 
        �u�@�� 10:     "/bank/personal/event",
        �u�@�� 11:     "/bank/small-business", 
        �u�@�� 12:     "/bank/corporate", 
        �u�@�� 13:     "/bank/digital", 
        �u�@�� 14:     "/bank/about", 
        �u�@�� 15:     "/bank/marketing",
        �u�@�� 16:     "/bank/iframe/widget", 
        �u�@�� 17:     "/bank/error", 
        �u�@�� 18:     "/bank/bank-en",
        �u�@�� 19:     "/bank/preview";
        */
        #endregion
        public void �ˬd�s��MegaMenu()
        {
            CreateReport("�s��MegaMenu���e�ˬd", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");
            Excel.Application excel_App = new Excel.Application(); //  �s�Wexcel���ε{��



            #region step 1: Ū�� expect result excel (expectResult.xlsx) (excel_Expect_xxxx) �ĥ|��(option�W��)�� expectResultArray �}�C��, 
            Excel.Workbook excel_Expect_WB = excel_App.Workbooks.Open($"{UserDataList.Upperfolderpath}testdata\\ExpectResult_0323.xlsx"); // open ���w���|excel
            Excel.Worksheet excel_Expect_WS = (Excel.Worksheet)excel_Expect_WB.Worksheets[1]; // ���wŪ��excel �ɲĤ@�Ӥu�@��
            int sheetRows = 4; // �u�@����� 
            Excel.Range expectResultRange = (Excel.Range)excel_Expect_WS.UsedRange; // export excel ���e to Range
            string[] expectResultOptionsNameArray = new string[expectResultRange.Count / sheetRows];
            for (int i = 0; i < expectResultRange.Count / sheetRows; i++)
            {
                expectResultOptionsNameArray[i] = (string)((Excel.Range)expectResultRange.Cells[i + 1, 4]).Value; // �Nexcel�ĥ|�椺�e��iexpectResultArray�}�C��
            }
            #endregion

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
            driver.Url = "https://www.esunbank.com.tw/bank/personal";
            Actions act = new Actions(driver);
            act.MoveToElement(driver.FindElement(By.CssSelector("ul.l-megaL1__list:nth-child(1) > li:nth-child(1) > div:nth-child(1)"))).Perform(); // Cursor���� MegaMenu "�s��"���

            for (int row= 1; row <=10; row++) // megamenu �s��10��
            {
                for (int line = 1; line <= 10; line++) // megamenu �s��10�C
                {
                    try
                    {
                        var cssSelector = $"ul.l-megaL1__list:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > ul:nth-child(1) > li:nth-child(1) > div:nth-child(1) > div:nth-child({row}) > ul:nth-child(2) > li:nth-child({line}) > a:nth-child(1)";
                        var actualOptionName = driver.FindElement(By.CssSelector(cssSelector)).Text; // �ثe����MegaMenu�W�ﶵ�W��
                        var actURL = driver.FindElement(By.CssSelector(cssSelector)).GetAttribute("href"); // �ثe����MegaMenu�W�ﶵ�����W�s��

                        int meetExpectResultIndex = Array.IndexOf(expectResultOptionsNameArray, actualOptionName); // �ثe����MegaMenu�ﶵ������expect result �}�C������m
                        if (meetExpectResultIndex == -1)
                        {
                            FAIL($"[Mega Menu][�s��] �ﶵ\"{actualOptionName}\''���b�w���ﶵ��");
                            continue;
                        }
                        else
                        {
                            string expectURL = (string)((Excel.Range)expectResultRange.Cells[meetExpectResultIndex + 1, 1]).Value; // Ū expect result excel���Ĥ@������

                            if (actURL == expectURL)
                            {
                                PASS($"[Mega Menu][�s��] �ﶵ \" {actualOptionName}\" �W�s�����}  \"{actURL}\" �ŦX�w��");
                            }
                            else if (actURL != expectURL)
                            {
                                FAIL($"[Mega Menu][�s��] �ﶵ \"{actualOptionName}\" �W�s�����} \"{actURL}\" (�w�����} \"{expectURL}\")");
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

            #region ���� & ������
            excel_Expect_WB.Close();
            excel_App.Quit();
            #endregion
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
        public void �ˬd�s�פ���()
        {

        }
    }
}
