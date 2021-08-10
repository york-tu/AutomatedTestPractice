using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;
using System;

namespace AutomatedTest.IntegrationTest.Personal.Deposit.GoldPriceInquiry
{
    public class 歷史黃金存摺牌價查詢_M版:IntegrationTestBase
    {
        public 歷史黃金存摺牌價查詢_M版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/wealth/gold/price/historical-price?dev=mobile"; // 登入M版 "歷史黃金存摺牌價查詢" 網頁
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void 歷史黃金存摺牌價查詢(string browser)
        {
            StartTestCase(browser, "歷史黃金存摺_牌價查詢", "York");
            INFO("");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\GoldPriceInquiry";
            Tools.CreateSnapshotFolder(snapshotpath);
            Tools.CleanUPFolder(snapshotpath);

            #region [自訂][起始] 從一年前的今天開始 "點擊小日曆日期" 查詢到今日
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[1]/label[1]")).Click(); // 點 "新台幣"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[2]/label[5]")).Click(); // 點 "自訂"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[4]/label[1]")).Click(); // 點 "走勢圖"

            Tools.ScrollPageUpOrDown(driver, 200);
            Thread.Sleep(100);

            for (int monthindex = 1; monthindex <= 13;) // 第 monthindex 個月份
            {
                if (monthindex == 13) // 判斷當月份為當下月份時
                {
                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // 點 "開始查詢" 鍵
                    string times = System.DateTime.Now.ToString("yyyy-MM-dd");
                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear(); // 清空起始日期欄位
                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys(times); // 欄位輸入當下當天日期
                }
                for (int weekindex = 2; weekindex <= 7; weekindex++) //  第月第 weekindex - 1 周
                {
                    for (int dayindex = 1; dayindex <= 7; dayindex++) // 當周第 dayindex 天
                    {
                        driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[3]/div/div[1]/span/button")).Click(); // 點 "起始小日曆" icon
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[2]/table[2]/tr[{weekindex}]/td[{dayindex}]"));

                        string date_value = date.GetAttribute("class");
                        if (date_value == "dp_disabled" || date_value == "dp_weekend dp_disabled" || date_value == "dp_not_in_month dp_disabled" || date_value == "dp_not_in_month dp_disabled") // 判斷當點到當月日曆上沒有的日子時
                        {
                            if (weekindex >= 5) // 週數為當月第5周以後
                            {
                                driver.FindElement(By.XPath("/html/body/div[2]/table[1]/tbody/tr/td[3]")).Click(); // 點 "下個月" 鍵
                                driver.FindElement(By.XPath("/html/body/div[2]/table[2]/tr[4]/td[4]")).Click(); // (切換月份後, 點擊當下月份日曆中間的日子)
                                goto quit;
                            }
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // 點 "開始查詢" 鍵
                        }
                        else
                        {
                            date.Click();
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // 點 "開始查詢" 鍵
                            Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
                        }
                    }
                }
            quit: monthindex++;
            }
            #endregion


            #region [自訂][結束] 從今天開始 "點擊小日曆日期" 查詢到一年前的今日
            Tools.ScrollPageUpOrDown(driver, 0);
            driver.Navigate().Refresh();
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[1]/label[1]")).Click(); // 點 "新台幣"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[2]/label[5]")).Click(); // 點 "自訂"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[4]/label[1]")).Click(); // 點 "走勢圖"

            Tools.ScrollPageUpOrDown(driver, 200);
            Thread.Sleep(100);

            for (int monthindex = 1; monthindex <= 13;) // 第 monthindex 個月份
            {
                for (int weekindex = 7; weekindex >= 1; weekindex--) //  第月第 weekindex - 1 周
                {
                    for (int dayindex = 7; dayindex >= 1; dayindex--) // 當周第 dayindex 天
                    {
                        driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[3]/div/div[2]/span/button")).Click(); // 點 "結束小日曆" icon
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[4]/table[2]/tr[{weekindex}]/td[{dayindex}]"));
                        string date_value = date.GetAttribute("class");

                        if (date_value == "dp_disabled" || date_value == "dp_weekend dp_disabled" || date_value == "dp_not_in_month dp_disabled" || date_value == "dp_not_in_month dp_weekend dp_disabled") // 判斷當點到當月日曆上沒有的日子時
                        {
                            if (weekindex <= 2) // 週數為當月前2周以前
                            {
                                driver.FindElement(By.XPath("/html/body/div[4]/table[1]/tbody/tr/td[1]")).Click(); // 點 "上個月" 鍵
                                driver.FindElement(By.XPath("/html/body/div[4]/table[2]/tr[4]/td[4]")).Click(); // (切換月份後, 點擊當下月份日曆中間的日子)
                                goto quit;
                            }
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // 點 "開始查詢" 鍵
                        }
                        else
                        {
                            if (weekindex == 2 && dayindex == 1) // 判斷當點到當月 "第一周第一天"時 (when 當月有這一天)
                            {
                                date.Click();
                                Tools.PageSnapshot($@"{snapshotpath}\結束日_{driver.FindElement(By.XPath("//*[@id='datepicker-to-compare']")).GetAttribute("value")}.png", driver);

                                driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[3]/div/div[2]/span/button")).Click(); // 點 "結束小日曆" icon
                                driver.FindElement(By.XPath("/html/body/div[4]/table[1]/tbody/tr/td[1]")).Click(); // 點 "上個月" 鍵
                                driver.FindElement(By.XPath("/html/body/div[4]/table[2]/tr[4]/td[4]")).Click(); // (切換月份後, 點擊當下月份日曆中間的日子)
                                goto quit;
                            }

                            date.Click();
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // 點 "開始查詢" 鍵
                            Tools.PageSnapshot($@"{snapshotpath}\結束日_{driver.FindElement(By.XPath("//*[@id='datepicker-to-compare']")).GetAttribute("value")}.png", driver);
                        }
                    }
                }
            quit: monthindex++;
            }
            #endregion


            #region [自訂][起始] 從一年前的今天開始 "欄位輸入日期" 查詢到今日
            int currentYear = Convert.ToInt32(System.DateTime.Now.ToString("yyyy"));
            int currentMonth = Convert.ToInt32(System.DateTime.Now.ToString("MM"));
            int currentDate = Convert.ToInt32(System.DateTime.Now.ToString("dd"));
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[1]/label[1]")).Click(); // 點 "新台幣"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[2]/label[5]")).Click(); // 點 "自訂"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[4]/label[1]")).Click(); // 點 "走勢圖"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();

            int maxFebruaryDays;

            for (int year = currentYear - 1; year <= currentYear;) // 2020 - 2021 year
            {
                switchYear:
                if (year == currentYear - 1) // 2020
                {
                    for (int mon = currentMonth; mon <= 12; mon++)
                    {
                        if (mon == currentMonth)
                        {
                            if (mon ==1 || mon == 3 || mon == 5|| mon == 7 || mon == 8 || mon == 10 || mon == 12)
                            {
                                for (int day = currentDate + 1; day <= 31; day++)
                                {
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);

                                    if (mon == 12 && day ==31)
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
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
                                }
                            }
                        }
                        else
                        {
                            if (mon == 1 || mon == 3 || mon == 5 || mon == 7 || mon == 8 || mon == 10 || mon == 12)
                            {
                                for (int day = 1; day <= 31; day++)
                                {
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);

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
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
                                }
                            }
                        }
                    }
                }
                else if (year == currentYear) // 2021
                {
                    for (int _mon = 1; _mon <= currentMonth; _mon++)
                    {
                        if (_mon == currentMonth)
                        {
                            for (int day = 1; day <= currentDate; day++)
                            {
                                driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{_mon}-{day}");
                                driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);

                                if (day == currentDate)
                                {
                                    goto finish;
                                }
                            }
                        }
                        else
                        {
                            if (_mon == 1 ||_mon == 3 || _mon == 5 || _mon == 7 || _mon == 8 || _mon == 10 || _mon == 12)
                            {
                                for (int day = 1; day <= 31; day++)
                                {
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{_mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
                                }
                            }
                            else if (_mon == 4 || _mon == 6 || _mon == 9 || _mon == 11)
                            {
                                for (int day = 1; day <= 30; day++)
                                {
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{_mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{_mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\起始日_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
                                }
                            }
                        }
                    }
                }
                finish:
                #endregion

                driver.Quit();
            }
        }
    }
}
