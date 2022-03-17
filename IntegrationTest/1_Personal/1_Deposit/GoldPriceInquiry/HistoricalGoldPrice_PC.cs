using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;
using System;

namespace AutomatedTest.IntegrationTest.Personal.Deposit.GoldPriceInquiry
{
    public class ���v�����s�P�P���d��_PC��:IntegrationTestBase
    {
        public ���v�����s�P�P���d��_PC��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/wealth/gold/price/historical-price"; // �n�JPC�� "���v�����s�P�P���d��" ����
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void ��J�_�l����d��(string browser)
        {
            StartTestCase(browser, "���v�����s�P_��J�_�l����d��_PC", "York");
            INFO("");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\GoldPriceInquiry";
            TestBase.CreateFolder(snapshotpath);
            TestBase.CleanUPFolder(snapshotpath);

            int currentYear = Convert.ToInt32(System.DateTime.Now.ToString("yyyy"));
            int currentMonth = Convert.ToInt32(System.DateTime.Now.ToString("MM"));
            int currentDate = Convert.ToInt32(System.DateTime.Now.ToString("dd"));

            #region [�ۭq][�_�l] �q�@�~�e�����Ѷ}�l "����J���" �d�ߨ줵��

            IWebElement StartDate = driver.FindElement(By.XPath("//*[@id='datepicker1']"));
            IWebElement StartSearch = driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[4]/a"));
            IWebElement Chart = driver.FindElement(By.XPath("//*[@id='tab-01']/div[1]/div[3]/div/label[1]"));

            TestBase.ScrollPageUpOrDown(driver, 300);
            int maxFebruaryDays;

            for (int year = currentYear; year <= currentYear;) // 2020 - 2021 year
            {
            switchYear:
                if (year == currentYear - 1) // 2020
                {
                    for (int mon = currentMonth; mon <= 12; mon++)
                    {
                        if (mon == currentMonth)
                        {
                            if (mon == 1 || mon == 3 || mon == 5 || mon == 7 || mon == 8 || mon == 10 || mon == 12)
                            {
                                for (int day = currentDate + 1; day <= 31; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                    if (mon == 12 && day == 31)
                                    {
                                        year++;
                                        goto switchYear;
                                    }
                                }
                            }
                            else if (mon == 4 || mon == 6 || mon == 9 || mon == 11)
                            {
                                for (int day = currentDate; day <= 30; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                }
                            }
                            else if (mon == 2)
                            {

                                if (currentYear % 4 == 0)
                                {
                                    maxFebruaryDays = 29;
                                }
                                else
                                {
                                    maxFebruaryDays = 28;
                                }

                                for (int day = currentDate; day <= maxFebruaryDays; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                }
                            }
                        }
                        else
                        {
                            if (mon == 1 || mon == 3 || mon == 5 || mon == 7 || mon == 8 || mon == 10 || mon == 12)
                            {
                                for (int day = 1; day <= 31; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                    if (mon == 12 && day == 31)
                                    {
                                        year++;
                                        goto switchYear;
                                    }
                                }
                            }
                            else if (mon == 4 || mon == 6 || mon == 9 || mon == 11)
                            {
                                for (int day = 1; day <= 30; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                }
                            }
                            else if (mon == 2)
                            {
                                if (currentYear % 4 == 0)
                                {
                                    maxFebruaryDays = 29;
                                }
                                else
                                {
                                    maxFebruaryDays = 28;
                                }

                                for (int day = 1; day <= maxFebruaryDays; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                }
                            }
                        }
                    }
                }
                else if (year == currentYear) // 2021
                {
                    for (int _mon = 8; _mon <= currentMonth; _mon++)
                    {
                        if (_mon == currentMonth)
                        {
                            for (int day = 1; day <= currentDate; day++)
                            {
                                StartDate.Clear();
                                StartDate.SendKeys($"{year}-{_mon}-{day}");
                                Chart.Click();
                                StartSearch.Click();
                                string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                Thread.Sleep(300);
                            }
                        }
                        else if (_mon != currentMonth)
                        {
                            if (_mon == 1 || _mon == 3 || _mon == 5 || _mon == 7 || _mon == 8 || _mon == 10 || _mon == 12)
                            {
                                for (int day = 1; day <= 31; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{_mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                }
                            }
                            else if (_mon == 4 || _mon == 6 || _mon == 9 || _mon == 11)
                            {
                                for (int day = 1; day <= 30; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{_mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                }
                            }
                            else if (_mon == 2)
                            {
                                if (currentYear % 4 == 0)
                                {
                                    maxFebruaryDays = 29;
                                }
                                else
                                {
                                    maxFebruaryDays = 28;
                                }
                                for (int day = 1; day <= maxFebruaryDays; day++)
                                {
                                    StartDate.Clear();
                                    StartDate.SendKeys($"{year}-{_mon}-{day}");
                                    Chart.Click();
                                    StartSearch.Click();
                                    string startColumnName = driver.FindElement(By.XPath("//*[@id='datepicker1']")).GetAttribute("value");
                                    TestBase.PageSnapshot(driver,$@"{snapshotpath}\�_�l��_{startColumnName}_PC.png");
                                    Thread.Sleep(300);
                                }
                            }
                        }
                    }
                }
                #endregion

                driver.Quit();
            }
        }
    }
}

