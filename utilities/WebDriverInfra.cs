using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;

namespace Utilities
{

    public enum BrowserType
    {
        Chrome,
        Firefox,
        IE,
    }

    public class WebDriverInfra
    {
        public static IWebDriver Create_Browser(BrowserType browserType)
        {
            switch (browserType)
            {
                case BrowserType.IE:

                    return new InternetExplorerDriver();

                case BrowserType.Firefox:

                    return new FirefoxDriver();

                case BrowserType.Chrome:

                    return new ChromeDriver();

                default:
                    return new ChromeDriver();
            }
        }
    }

    
}
