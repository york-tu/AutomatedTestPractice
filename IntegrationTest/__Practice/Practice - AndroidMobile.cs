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
    public class �m��_�s�I1:IntegrationTestBase
    {
        public �m��_�s�I1(ITestOutputHelper output, Setup testsetup) : base(output, testsetup)
        {
            testurl = domain + "https://popcat.click/";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]


        public void �s�I(string browser)
        {
            StartTestCaseForCustomizedSize(browser, "�m��_�s�I", "York", 100, 100, 200, 400);

            for (int i = 0; i < 1000000000; i++)
            {

                driver.FindElement(By.XPath("//*[@id='app']/div")).Click();
                //Thread.Sleep(1);
            }

            driver.Quit();
        }
    }
}


