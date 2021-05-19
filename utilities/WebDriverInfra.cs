using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;

namespace OpenQA.Selenium
{
    public class WebDriverInfra
    {
        public static IWebDriver Create_Browser(BrowserType browserType)
        {
            return browserType switch
            {
                BrowserType.Chrome => new ChromeDriver(),
                BrowserType.Firefox => new FirefoxDriver(),
                BrowserType.IE => new EdgeDriver(),
                _ => throw new ArgumentOutOfRangeException(nameof(browserType), browserType, null),
            };
        }
    }

}
