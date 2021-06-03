using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Xunit;
using System;
using System.IO;
using System.Linq;
using References;
using IronXL;
using System.Text.RegularExpressions;
using CSVHeader;
using CsvHelper;
using System.Globalization;

namespace HouseLoanMessageBoardTest
{
    public class HouseLoanMessageBoard // �ЫζU�گd������
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/house-loan?dev=mobile";
        private readonly string Version = "Mobile";
        private readonly string testcase_name = "�w��άd��";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void HLMB_M(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100���������������e, �_�h����, ���������i�U�@�B.
                //driver.Manage().Window.Maximize();


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

                    using (var reader = new StreamReader(UserDataList.csvpath)) //ŪCSV��
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

                    string notification_wordings = driver.SwitchTo().Alert().Text;
                    System.Diagnostics.Debug.WriteLine(notification_wordings);
                    Tools.SnapshotFullScreen(@"D:\" + testcase_name+ "_" + browserType + "_" + Version + "_" + timesavepath + ".png"); // �I�Ϸ��U�e��

                    driver.SwitchTo().Alert().Accept();
                }
                catch (Exception)
                {

                }
                driver.Quit();
            }
        }    
    }
}