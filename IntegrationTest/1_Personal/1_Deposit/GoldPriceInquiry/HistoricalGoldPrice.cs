using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System.Threading;
using System;

namespace AutomatedTest.IntegrationTest.Personal.Deposit.GoldPriceInquiry
{
    public class ���v�����s�P�P���d��_M��:IntegrationTestBase
    {
        public ���v�����s�P�P���d��_M��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/wealth/gold/price/historical-price?dev=mobile"; // �n�JM�� "���v�����s�P�P���d��" ����
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void ���v�����s�P�P���d��(string browser)
        {
            StartTestCase(browser, "���v�����s�P_�P���d��", "York");
            INFO("");

            string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\GoldPriceInquiry";
            Tools.CreateSnapshotFolder(snapshotpath);
            Tools.CleanUPFolder(snapshotpath);

            #region [�ۭq][�_�l] �q�@�~�e�����Ѷ}�l "�I���p�����" �d�ߨ줵��
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[1]/label[1]")).Click(); // �I "�s�x��"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[2]/label[5]")).Click(); // �I "�ۭq"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[4]/label[1]")).Click(); // �I "���չ�"

            Tools.ScrollPageUpOrDown(driver, 200);
            Thread.Sleep(100);

            for (int monthindex = 1; monthindex <= 13;) // �� monthindex �Ӥ��
            {
                if (monthindex == 13) // �P�_��������U�����
                {
                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // �I "�}�l�d��" ��
                    string times = System.DateTime.Now.ToString("yyyy-MM-dd");
                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear(); // �M�Ű_�l������
                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys(times); // ����J��U��Ѥ��
                }
                for (int weekindex = 2; weekindex <= 7; weekindex++) //  �Ĥ�� weekindex - 1 �P
                {
                    for (int dayindex = 1; dayindex <= 7; dayindex++) // ��P�� dayindex ��
                    {
                        driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[3]/div/div[1]/span/button")).Click(); // �I "�_�l�p���" icon
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[2]/table[2]/tr[{weekindex}]/td[{dayindex}]"));

                        string date_value = date.GetAttribute("class");
                        if (date_value == "dp_disabled" || date_value == "dp_weekend dp_disabled" || date_value == "dp_not_in_month dp_disabled" || date_value == "dp_not_in_month dp_disabled") // �P�_���I������W�S������l��
                        {
                            if (weekindex >= 5) // �g�Ƭ�����5�P�H��
                            {
                                driver.FindElement(By.XPath("/html/body/div[2]/table[1]/tbody/tr/td[3]")).Click(); // �I "�U�Ӥ�" ��
                                driver.FindElement(By.XPath("/html/body/div[2]/table[2]/tr[4]/td[4]")).Click(); // (���������, �I����U�����䤤������l)
                                goto quit;
                            }
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // �I "�}�l�d��" ��
                        }
                        else
                        {
                            date.Click();
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // �I "�}�l�d��" ��
                            Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
                        }
                    }
                }
            quit: monthindex++;
            }
            #endregion


            #region [�ۭq][����] �q���Ѷ}�l "�I���p�����" �d�ߨ�@�~�e������
            Tools.ScrollPageUpOrDown(driver, 0);
            driver.Navigate().Refresh();
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[1]/label[1]")).Click(); // �I "�s�x��"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[2]/label[5]")).Click(); // �I "�ۭq"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[4]/label[1]")).Click(); // �I "���չ�"

            Tools.ScrollPageUpOrDown(driver, 200);
            Thread.Sleep(100);

            for (int monthindex = 1; monthindex <= 13;) // �� monthindex �Ӥ��
            {
                for (int weekindex = 7; weekindex >= 1; weekindex--) //  �Ĥ�� weekindex - 1 �P
                {
                    for (int dayindex = 7; dayindex >= 1; dayindex--) // ��P�� dayindex ��
                    {
                        driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[3]/div/div[2]/span/button")).Click(); // �I "�����p���" icon
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[4]/table[2]/tr[{weekindex}]/td[{dayindex}]"));
                        string date_value = date.GetAttribute("class");

                        if (date_value == "dp_disabled" || date_value == "dp_weekend dp_disabled" || date_value == "dp_not_in_month dp_disabled" || date_value == "dp_not_in_month dp_weekend dp_disabled") // �P�_���I������W�S������l��
                        {
                            if (weekindex <= 2) // �g�Ƭ����e2�P�H�e
                            {
                                driver.FindElement(By.XPath("/html/body/div[4]/table[1]/tbody/tr/td[1]")).Click(); // �I "�W�Ӥ�" ��
                                driver.FindElement(By.XPath("/html/body/div[4]/table[2]/tr[4]/td[4]")).Click(); // (���������, �I����U�����䤤������l)
                                goto quit;
                            }
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // �I "�}�l�d��" ��
                        }
                        else
                        {
                            if (weekindex == 2 && dayindex == 1) // �P�_���I���� "�Ĥ@�P�Ĥ@��"�� (when ��릳�o�@��)
                            {
                                date.Click();
                                Tools.PageSnapshot($@"{snapshotpath}\������_{driver.FindElement(By.XPath("//*[@id='datepicker-to-compare']")).GetAttribute("value")}.png", driver);

                                driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[3]/div/div[2]/span/button")).Click(); // �I "�����p���" icon
                                driver.FindElement(By.XPath("/html/body/div[4]/table[1]/tbody/tr/td[1]")).Click(); // �I "�W�Ӥ�" ��
                                driver.FindElement(By.XPath("/html/body/div[4]/table[2]/tr[4]/td[4]")).Click(); // (���������, �I����U�����䤤������l)
                                goto quit;
                            }

                            date.Click();
                            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click(); // �I "�}�l�d��" ��
                            Tools.PageSnapshot($@"{snapshotpath}\������_{driver.FindElement(By.XPath("//*[@id='datepicker-to-compare']")).GetAttribute("value")}.png", driver);
                        }
                    }
                }
            quit: monthindex++;
            }
            #endregion


            #region [�ۭq][�_�l] �q�@�~�e�����Ѷ}�l "����J���" �d�ߨ줵��
            int currentYear = Convert.ToInt32(System.DateTime.Now.ToString("yyyy"));
            int currentMonth = Convert.ToInt32(System.DateTime.Now.ToString("MM"));
            int currentDate = Convert.ToInt32(System.DateTime.Now.ToString("dd"));
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[1]/label[1]")).Click(); // �I "�s�x��"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[2]/label[5]")).Click(); // �I "�ۭq"
            driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[4]/label[1]")).Click(); // �I "���չ�"
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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);

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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);

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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);

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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
                                }
                            }
                            else if (_mon == 4 || _mon == 6 || _mon == 9 || _mon == 11)
                            {
                                for (int day = 1; day <= 30; day++)
                                {
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).Clear();
                                    driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).SendKeys($"{year}-{_mon}-{day}");
                                    driver.FindElement(By.XPath("//*[@id='panelMain']/div/div[2]/div/div[2]/section/div/div[1]/div[5]/label")).Click();
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
                                    Tools.PageSnapshot($@"{snapshotpath}\�_�l��_{driver.FindElement(By.XPath("//*[@id='datepicker-from-compare']")).GetAttribute("value")}.png", driver);
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
