using System;
using Xunit;
using Xunit.Abstractions;
using AventStack.ExtentReports;
using AutomatedTest.Utilities;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OpenQA.Selenium;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using AventStack.ExtentReports.MarkupUtils;

namespace AutomatedTest.IntegrationTest
{

    public class IntegrationTestBase:XunitContextBase, IClassFixture<Setup>
    {
        protected ExtentReportHelper _extentReport;
        protected ExtentTest _extentTObj;
        protected string testurl; // 測試用的網址
        protected string testFullName; // 測試案例全名
        protected string browserName; // 測試用的瀏覽器
        protected string runtimeEnvironment; // 執行環境
        protected bool result = false; // 初始測試結果
        protected string domain;
        protected bool getreport;
        private IConfiguration _config;
        protected ServiceProvider _serviceProvider;
        protected ITestOutputHelper _output;
        protected IWebDriver driver;
        BrowserHelper browser;

        public IntegrationTestBase(ITestOutputHelper output, Setup testSetup) : base(output)
        {
            _serviceProvider = testSetup.ServiceProvider;

            _config = _serviceProvider.GetService<IConfiguration>();
            runtimeEnvironment = _config.GetValue<string>("RuntimeEnvironment");
            string getjson = "Domain: " + runtimeEnvironment;

            //組成url
            domain = _config.GetValue<string>(getjson);

            //確認是否產生report
            getreport = Convert.ToBoolean(_config.GetValue<string>("GetReport"));
            testFullName = Context.UniqueTestName;

        }


        /// <summary>
        /// 開始測試 & 初始設定
        /// </summary>
        /// <param name="browserArg"></param>
        /// <param name="reportName"></param>
        /// <param name="testURL"></param>
        public void StartTestCase(string browserArg, string reportName, string tester)
        {
            browser = new BrowserHelper(browserArg);
           this.driver = browser.driver;
            driver.Navigate().GoToUrl(testurl);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
            driver.Manage().Window.Maximize();

            if (getreport)
            {
                CreateReport(reportName, tester);
                browserName = browserArg;
            }
        }
        public void CreateReport(string reportName, string tester)
        {
            _extentReport = new ExtentReportHelper(reportName, tester);
            _extentReport.Init();
            _extentReport.CreateTestCase(testFullName, "");
            _extentReport.GetTestEnvironment(runtimeEnvironment);
            if (browser !=null)
            {
                _extentReport.TestCaseDriverVersion(browser.GetBrowserName() + " v" + browser.GetBrowserVersion());
            }
            else
            {
                _extentReport.TestCaseDriverVersion("N/A");
            }
            _extentTObj = _extentReport.ExtentTestObjects;
        }
        public void CloseBrowser()
        {
            if (this.driver != null)
            {
                this.driver.Quit();
            }
            if (getreport)
            {
                _extentReport.ExportReport();
            }
        }
        public bool CloseWindows(string window)
        {
            try
            {
                this.driver.SwitchTo().Window(window);
                this.driver.Quit();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public void StartTestCaseForCustomizedSize(string browserArg, string reportName, string tester, int position_X, int position_Y, int width, int height) // For 客製化開啟的視窗位置 + 大小
        {
            browser = new BrowserHelper(browserArg);
            //this.driver = browser.driver;
            driver.Navigate().GoToUrl(testurl);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
            driver.Manage().Window.Position = new Point(position_X, position_Y); //設定網頁開啟在畫面什麼位置
            driver.Manage().Window.Size = new Size(width, height); // 設定開啟的網頁大小

            if (getreport)
            {
                CreateReport(reportName, tester);
                browserName = browserArg;
            }
        }


        ///<summary>
        ///Extent Report Logs
        ///</summary> 
        public void FAIL(string content)
        {
            if (getreport)
            {
                _extentTObj.Fail(MarkupHelper.CreateLabel(content, ExtentColor.Red));
            }
        }
        public void PASS(string content)
        {
            if (getreport)
            {
                _extentTObj.Pass(MarkupHelper.CreateLabel(content, ExtentColor.Green));
            }
        }
        public void ERROR(string content)
        {
            if (getreport)
            {
                _extentTObj.Error(MarkupHelper.CreateLabel(content, ExtentColor.Red));
            }
        }
        public void INFO(string content)
        {
            if (getreport)
            {
                //_extentTObj.Info(MarkupHelper.CreateLabel(content, ExtentColor.Yellow));
                _extentTObj.Info(content);

            }
        }
        public void WARNING(string content)
        {
            if (getreport)
            {
                _extentTObj.Warning(MarkupHelper.CreateLabel(content, ExtentColor.Orange));

            }
        }
        public void INFOfromCodeFormat (string content)
        {
            if (getreport)
            {
                _extentTObj.Info(MarkupHelper.CreateCodeBlock(content));
            }
        }
    }
}

