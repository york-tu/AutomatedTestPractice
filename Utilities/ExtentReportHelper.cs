using System;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;

namespace AutomatedTest.Utilities
{
    public class ExtentReportHelper
    {
        private ExtentReports _extentReportsObject;
        private string _testName;
        private string _environment;
        private string _driverVersion;

        public ExtentTest ExtentTestObjects;

        List<TestResult> testReport = new List<TestResult>();

        public ExtentReportHelper(string testName)
        {
            _testName = testName;
        }

        public void Init()
        {
            _extentReportsObject = new ExtentReports();

           // var dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location + "\\");
            var filename = String.Format("{0}_{1}.html", _testName, DateTime.Now.ToString("yyyyMMddHHmm"));
            //var _htmlReporter = new ExtentV3HtmlReporter(dir + filename);
            var _htmlReporter = new ExtentV3HtmlReporter($@"{UserDataList.folderpath}\{filename}");

            _extentReportsObject.AttachReporter(_htmlReporter);
            _extentReportsObject.AddSystemInfo("測試案例", _testName);

        }

        public void CreateTestCase(string caseName, string typeName = "")
        {
            if (_extentReportsObject == null)
            {
                Init();
            }

            string setName = string.IsNullOrEmpty(typeName) ? caseName : string.Format("{0}.{1}", typeName, caseName);
            ExtentTestObjects = _extentReportsObject.CreateTest(setName);
        }

        public void GetTestEnvironment(string environment)
        {
            _environment = environment;
            _extentReportsObject.AddSystemInfo("測試環境", _environment);
        }

        public void TestCaseDriverVersion(string driverVersion)
        {
            _driverVersion = driverVersion;
            _extentReportsObject.AddSystemInfo("瀏覽器版本", _driverVersion);
        }

        public void ExportReport()
        {
            _extentReportsObject.Flush();
        }

        public void AddTestResult (string testcasename)
        {
            List<ResultLog> resultLogs = new List<ResultLog>();
            testReport.Add(new TestResult { TestcaseName = testcasename, Result = resultLogs });
        }

        public class TestResult
        {
            public string TestcaseName { get; set; }
            public List<ResultLog> Result { get; set; }
        }

        public class ResultLog
        {
            public string Status { get; set; }
            public string Information { get; set; }
        }

    }
   
}
