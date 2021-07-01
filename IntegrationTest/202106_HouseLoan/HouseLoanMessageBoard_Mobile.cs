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
    public class 房貸留言版提示視窗文字調整_PC版 : IntegrationTestBase // 房屋貸款留言測試
    {
        public 房貸留言版提示視窗文字調整_PC版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/house-loan?dev=mobile";
        }

        private readonly string Version = "Mobile";
        private readonly string testcase_name = "預售屋查詢";

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
      
        public void 提示視窗文字檢核(string browser)
        {
            StartTestCase(browser, "房貸留言版提示視窗文字調整_PC版", "York");
            INFO("確認視窗文字");
                try
                {
                    string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");

                    IWebElement loan_amount = driver.FindElement(By.XPath("//*[@id='loanAmount']"));
                    loan_amount.Clear();
                    loan_amount.SendKeys("10"); // 填申貸額度
                    System.Threading.Thread.Sleep(300);

                    IWebElement loan_purpose = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[3]/ul/li/span"));
                    loan_purpose.Click(); //點貸款用途下拉選單
                    System.Threading.Thread.Sleep(300);
                    IWebElement buy_house = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[3]/ul/li/ul/li[2]/span"));
                    buy_house.Click(); //選"購買房屋"
                    System.Threading.Thread.Sleep(300);

                    string csvpath = $@"{UserDataList.Upperfolderpath}\testdata\UserInfo.csv";
                    using (var reader = new StreamReader(csvpath)) //讀CSV檔
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
                                name.SendKeys(user.NAME); //輸入姓名
                                System.Threading.Thread.Sleep(300);

                                IWebElement ID = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[5]/input"));
                                ID.Clear();
                                ID.SendKeys(user.ID); //輸入身份證
                                System.Threading.Thread.Sleep(300);

                                IWebElement cellphone = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[7]/input"));
                                cellphone.Clear();
                                cellphone.SendKeys(user.PHONE); // 輸入手機號碼
                                System.Threading.Thread.Sleep(300);

                                break;
                            }
                            n++;
                        }
                    }

                    IWebElement mailing_address = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[9]/div[1]/ul/li/span"));
                    mailing_address.Click(); //點 通訊地址 縣市 下拉選單
                    IWebElement address_taipei = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[9]/div[1]/ul/li/ul/li[3]"));
                    address_taipei.Click(); // 選: 台北市
                    IWebElement input_address = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[9]/div[3]/input"));
                    input_address.SendKeys("中正路一段1號"); // 填地址

                    IWebElement house_address = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[10]/div[1]/ul/li/span"));
                    house_address.Click(); //點 房屋位置 下拉選單
                    IWebElement house_taipei = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[10]/div[1]/ul/li/ul/li[3]"));
                    house_taipei.Click(); //選: 台北市
                    
                    IWebElement contact_time = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[14]/ul/li/span"));
                    contact_time.Click(); //點 方便聯絡時間 下拉選單
                    IWebElement time_morning = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[2]/section/div[14]/ul/li/ul/li[2]/span"));
                    time_morning.Click(); //選: 早上

                    IWebElement iagree = driver.FindElement(By.XPath("//*[@id='mainContent']/div/div[3]/div[3]/div/div[1]/a[1]/span/span"));
                    iagree.Click(); //點 我已閱讀並同意

                    IWebElement submit = driver.FindElement(By.XPath("//*[@id='submit']"));
                    submit.Click(); //點 送出

                    WebDriverWait popup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                    popup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent()); // 等待直到alert popup window跳出

                    string notification_wordings = driver.SwitchTo().Alert().Text;


                    string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\HouseLoanMessageBoard";
                    Tools.CreateSnapshotFolder(snapshotpath);
                    System.Threading.Thread.Sleep(100);
                    Tools.FullScreenshot($@"{snapshotpath}\{testcase_name}_{browser}_{Version}_{timesavepath}.png"); // 截圖當下畫面

                    driver.SwitchTo().Alert().Accept();
                }
                catch (Exception)
                {

                }
                driver.Quit();
        }    
    }
}
