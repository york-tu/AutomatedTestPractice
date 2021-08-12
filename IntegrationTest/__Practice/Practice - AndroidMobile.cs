using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.IO;
using Xunit.Abstractions;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System;

namespace AutomatedTest.IntegrationTest.__Practice
{
    public class 練習_連點1:IntegrationTestBase
    {
        public 練習_連點1(ITestOutputHelper output, Setup testsetup) : base(output, testsetup)
        {
            testurl = domain + "https://popcat.click/";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]


        public void 連點(string browser)
        {
            StartTestCaseForCustomizedSize(browser, "練習_連點", "York", 100, 100, 200, 400);

            for (int i = 0; i < 1000000000; i++)
            {

                driver.FindElement(By.XPath("//*[@id='app']/div")).Click();
                //Thread.Sleep(1);
            }

            driver.Quit();
        }
    }
}


