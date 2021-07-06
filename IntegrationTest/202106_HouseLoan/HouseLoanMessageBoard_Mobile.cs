using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using Xunit;
using System;
using System.IO;
using System.Linq;
using AutomatedTest.Utilities;
using CsvHelper;
using System.Globalization;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.HouseLoan
{
    public class �жU�d�������ܵ�����r�վ�_PC�� : IntegrationTestBase // �ЫζU�گd������
    {
        public �жU�d�������ܵ�����r�վ�_PC��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/house-loan?dev=mobile";
        }

        private readonly string Version = "Mobile";
        private readonly string testcase_name = "�w��άd��";

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
      
        public void ���ܵ�����r�ˮ�(string browser)
        {
            StartTestCase(browser, "�жU�d�������ܵ�����r�վ�_PC��", "York");
            INFO("�T�{������r");
                try
                {
                    string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");

                    IWebElement loan_amount = driver.FindElement(By.XPath("//*[@id='loanAmount']"));
                    loan_amount.Clear();
                    loan_amount.SendKeys("10"); // ��ӶU�B��
                    System.Threading.Thread.Sleep(300);

                    IWebElement loan_purpose = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[3]/ul/li/span"));
                    loan_purpose.Click(); //�I�U�ڥγ~�U�Կ��
                    System.Threading.Thread.Sleep(300);
                    IWebElement buy_house = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[3]/ul/li/ul/li[2]/span"));
                    buy_house.Click(); //��"�ʶR�Ы�"
                    System.Threading.Thread.Sleep(300);

                    string csvpath = $@"{UserDataList.Upperfolderpath}\testdata\UserInfo.csv";
                    using (var reader = new StreamReader(csvpath)) //ŪCSV��
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        var records = csv.GetRecords<UserDataList>().ToList();
                        int n = 1;
                        foreach (var user in records)
                        {
                            if (n == 4)
                            {
                                IWebElement name = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[4]/input"));
                                name.Clear();
                                name.SendKeys(user.NAME); //��J�m�W
                                System.Threading.Thread.Sleep(300);

                                IWebElement ID = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[5]/input"));
                                ID.Clear();
                                ID.SendKeys(user.ID); //��J������
                                System.Threading.Thread.Sleep(300);

                                IWebElement cellphone = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[7]/input"));
                                cellphone.Clear();
                                cellphone.SendKeys(user.PHONE); // ��J������X
                                System.Threading.Thread.Sleep(300);

                                break;
                            }
                            n++;
                        }
                    }

                    IWebElement mailing_address = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[9]/div[1]/ul/li/span"));
                    mailing_address.Click(); //�I �q�T�a�} ���� �U�Կ��
                    IWebElement address_taipei = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[9]/div[1]/ul/li/ul/li[3]"));
                    address_taipei.Click(); // ��: �x�_��
                    IWebElement input_address = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[9]/div[3]/input"));
                    input_address.SendKeys("�������@�q1��"); // ��a�}

                    IWebElement house_address = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[10]/div[1]/ul/li/span"));
                    house_address.Click(); //�I �ЫΦ�m �U�Կ��
                    IWebElement house_taipei = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[10]/div[1]/ul/li/ul/li[3]"));
                    house_taipei.Click(); //��: �x�_��
                    
                    IWebElement contact_time = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[14]/ul/li/span"));
                    contact_time.Click(); //�I ��K�p���ɶ� �U�Կ��
                    IWebElement time_morning = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[14]/ul/li/ul/li[2]/span"));
                    time_morning.Click(); //��: ���W

                    IWebElement iagree = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[3]/div/div[1]/a[1]/span/span"));
                    iagree.Click(); //�I �ڤw�\Ū�æP�N

                    IWebElement submit = driver.FindElement(By.XPath("//*[@id='submit']"));
                    submit.Click(); //�I �e�X

                    WebDriverWait popup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                    popup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent()); // ���ݪ���alert popup window���X

                    string notification_wordings = driver.SwitchTo().Alert().Text;


                    string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\HouseLoanMessageBoard";
                    Tools.CreateSnapshotFolder(snapshotpath);
                    System.Threading.Thread.Sleep(100);
                    Tools.FullScreenshot($@"{snapshotpath}\{testcase_name}_{browser}_{Version}_{timesavepath}.png"); // �I�Ϸ�U�e��

                    driver.SwitchTo().Alert().Accept();
                }
                catch (Exception)
                {

                }
                driver.Quit();
        }    
    }
}
