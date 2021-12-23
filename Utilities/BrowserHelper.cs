using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;

namespace AutomatedTest.Utilities
{

    public class BrowserHelper
    {
        private IWebDriver driver;
        public IDictionary<string, object> vars { get; private set; }
        private IJavaScriptExecutor js;
        private string browserVersion;
        private string browserName;
        ICapabilities capabilities;


        /// <summary>
        /// 初始化參數
        /// </summary>
        public BrowserHelper()
        {
            js = (IJavaScriptExecutor)driver;
            vars = new Dictionary<string, object>();
        }


        /// <summary>
        /// 選擇初始化之瀏覽器
        /// </summary>
        /// <param name="browser"></param>
        public BrowserHelper(string browser)
        {
            js = (IJavaScriptExecutor)driver;
            vars = new Dictionary<string, object>();
            switch (browser)
            {
                case "Chrome":
                    ChromeOptions chromeoptions = new ChromeOptions();
                    //chromeoptions.setExperimentalOption();
                    chromeoptions.AddAdditionalCapability("useAutomationExtension", false);
                    chromeoptions.AddExcludedArgument("enable-automation");
                   // chromeoptions.AddExcludedArgument("--disable-infobars");
                   // chromeoptions.AddAdditionalCapability("useAutomationExtension", false);
                    driver = new ChromeDriver();
                    capabilities = ((RemoteWebDriver)driver).Capabilities;
                    browserVersion = capabilities.GetCapability("browserVersion").ToString();
                    browserName = "Chrome";
                    break;

                case "Firefox":
                    driver = new FirefoxDriver();
                    capabilities = ((RemoteWebDriver)driver).Capabilities;
                    browserVersion = capabilities.GetCapability("browserVersion").ToString();
                    browserName = "Firefox";
                    break;

                case "IE":
                    driver = new InternetExplorerDriver();
                    browserName = "IE";
                    break;
            }
        }


        /// <summary>
        /// 瀏覽器list
        /// </summary>
        public static IEnumerable<object[]> BrowserList =>
            new List<object[]>
            {
                new object[]{"Chrome"},
               // new object[]{"Firefox"},
               // new object[]{"IE"}
            };


        public string GetBrowserVersion()
        {
            return browserVersion;
        }
        public string GetBrowserName()
        {
            return browserName;
        }

    }
}
