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

namespace AutomatedTest.IntegrationTest.Regression
{
    public class II_I_�ˬd�������e:IntegrationTestBase
    {
        public II_I_�ˬd�������e(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
        }

        [Theory]
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
        [InlineData(new int[] { 2 })]
        public void �ˬd�������e�O�_�ŦX�w��(int[] sheet)
        {
            CreateReport("�������e�ˬd", "York");
            //TestBase.KillProcess("chromedriver.exe");
            //TestBase.KillProcess("chrome.exe");
            //TestBase.KillProcess("EXCEL.EXE");
            //TestBase.KillProcess("conhost.exe");
            Excel.Application excel_App = new Excel.Application(); //  �s�Wexcel���ε{��

            #region step 1: Ū��sitemap excel���e�� excel_WB
            Excel.Workbook excel_WB = excel_App.Workbooks.Open(UserDataList.sitemapsExcelPath); // pass���w���|excel�������eto "excel_WB"
            #endregion

            #region step 2: Ū�� expect result excel (expectResult.xlsx)�Ĥ@�椺�e (excel_Expect_xxxx) �� expectResultArray �}�C��
            Excel.Workbook excel_Expect_WB = excel_App.Workbooks.Open($"{UserDataList.Upperfolderpath}testdata\\ExpectResult_0323.xlsx"); // open ���w���|excel
            Excel.Worksheet excel_Expect_WS = (Excel.Worksheet)excel_Expect_WB.Worksheets[1]; // ���wŪ��excel �ɲĤ@�Ӥu�@��
            int sheetRows = 4; // �u�@����� 
            Excel.Range expectResultRange = (Excel.Range)excel_Expect_WS.UsedRange; // export excel ���e to Range
            string[] expectResultArray = new string[expectResultRange.Count / sheetRows];
            for (int i = 0; i < expectResultRange.Count / sheetRows; i++)
            {
                expectResultArray[i] = (string)((Excel.Range)expectResultRange.Cells[i + 1, 1]).Value; // �Nexcel�Ĥ@�椺�e��iexpectResultArray�}�C��
            }
            #endregion

            #region  browser ���}�Һ����]�w 
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


            foreach (var index in sheet)
            {
                Excel.Worksheet excel_WS = (Excel.Worksheet)excel_WB.Worksheets[index]; //  Ūsitemaps excel���w�u�@���e
                Excel.Range range = excel_WS.UsedRange; // ���X�u�@���e, pass to "range"
                INFO(excel_WS.Name);

                for (int i = 0; i < range.Count; i++)
                {
                    string strURL = (string)((Excel.Range)range.Cells[i + 1, 1]).Value; // Ū�X�� i ��URL
                    int meetExpectResultURLIndex = Array.IndexOf(expectResultArray, strURL); // �j�M�ؼ�URL������expect result �}�C������m

                    if (meetExpectResultURLIndex == -1) // ��URL�b������W�j�M����
                    {
                        WARNING($"[�s�W/�ܰ�] URL:  {strURL}");
                        continue;
                    }
                    else
                    {
                        string cssSelector = (string)((Excel.Range)expectResultRange.Cells[meetExpectResultURLIndex + 1, 2]).Value; // Ū expect result excel���ĤG������
                        string expectString = (string)((Excel.Range)expectResultRange.Cells[meetExpectResultURLIndex + 1, 3]).Value;  // Ū expect result excel���ĤT������

                        ((IJavaScriptExecutor)driver).ExecuteScript("window.open();"); // �t�}�s��
                        driver.SwitchTo().Window(driver.WindowHandles.Last()); // focus on �s���W 
                        driver.Navigate().GoToUrl(strURL);
                        #region ������۰ʤ���M����, click "�����q����" �j����^PC��
                        if (driver.Url.ToString().Contains("?dev=mobile")) // workaround: ������۰ʤ���M����, �j����^PC��
                        {
                            TestBase.ScrollPageUpOrDown(driver, 1500);
                            driver.FindElementByClassName("changeTarget").Click();
                        }
                        #endregion
                        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(600); //600�������������e, �_�h����, ���������i�U�@�B.
                        if (driver.Url.ToString() != strURL) // �ˬd�����}�ҷ�U���}�O�_����J���} (�P�_�����O�_��redirect)
                        {
                            WARNING($"[Page Redirect], {driver.Url.ToString()}  (Expect: {strURL})");
                            WARNING(TestBase.PageSnapshotToReport(driver));
                            continue;
                        }
                        else if (strURL == "https://easyfee.esunbank.com.tw/index.action")
                        {
                            WARNING($"{strURL}, Keyword: {expectString}");
                            WARNING(TestBase.PageSnapshotToReport(driver));
                            continue;
                        }
                        else if (driver.Url.ToString() == "https://www.esunbank.com.tw/bank/personal/event/calendar/events") // Workaround 1 : ���ʤ��URL >>> �ݧ��Ѥ��
                        {
                            string day_of_week = DateTime.Now.ToString("dddd", new CultureInfo("en-us")); // �^��P���X(e.g., Wednesday)
                            string day = DateTime.Now.ToString("dd"); // ��� (e.g., 23)
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
                            PASS($"{strURL}, Keyword: {expectString}");
                            PASS(TestBase.PageSnapshotToReport(driver));
                        }
                        catch (Exception e)
                        {
                            FAIL($"{strURL}, Exception:{e.Message}");
                            FAIL(TestBase.PageSnapshotToReport(driver));
                        }
                        driver.SwitchTo().Window(driver.WindowHandles.Last()).Close(); // �����s��
                        driver.SwitchTo().Window(driver.WindowHandles.First()); // ���^�쭶
                    }
                }
            }
            #region ���� & ������
            excel_WB.Close();
            excel_Expect_WB.Close();
            excel_App.Quit();
            #endregion
            CloseBrowser();
            driver.Close();
            driver.Quit();
        }
    }
}
