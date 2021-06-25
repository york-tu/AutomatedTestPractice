using System;
using Xunit;
using Xunit.Abstractions;
using AventStack.ExtentReports;
using AutomatedTest.Utilities;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OpenQA.Selenium;


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

            getreport = Convert.ToBoolean(_config.GetValue<string>("GetReport"));
            testFullName = Context.UniqueTestName;

        }

        /// <summary>
        /// 開始測試 & 初始設定
        /// </summary>
        /// <param name="browserArg"></param>
        /// <param name="reportName"></param>
        /// <param name="testURL"></param>
        public void StartTestCase(string browserArg, string reportName)
        {
            browser = new BrowserHelper(browserArg);
            this.driver = browser.driver;
            driver.Navigate().GoToUrl(testurl);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
            driver.Manage().Window.Maximize();

            if (getreport)
            {
                CreateReport(reportName);
                browserName = browserArg;
            }
        }

        public void CreateReport(string reportName)
        {

            _extentReport = new ExtentReportHelper(reportName);
            _extentReport.Init();
            _extentReport.CreateTestCase(testFullName, "");
            _extentReport.GetTestEnvironment(runtimeEnvironment);
            _extentReport.TestCaseDriverVersion(browserName + browser.GetBrowserVersion());
            _extentTObj = _extentReport.ExtentTestObjects;

        }
        
        public void CloseBrowser ()
        {
            this.driver.Quit();
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



        ///<summary>
        ///Extent Report Logs
        ///</summary> 
        public void FAIL(string content)
        {
            if (getreport)
            {
                _extentTObj.Fail(content);
            }
        }
        public void PASS(string content)
        {
            if (getreport)
            {
                _extentTObj.Pass(content);
            }
        }
        public void ERROR(string content)
        {
            if (getreport)
            {
                _extentTObj.Error(content);
            }
        }
        public void INFO(string content)
        {
            if (getreport)
            {
                _extentTObj.Info(content);
                
            }
        }

    }
}
