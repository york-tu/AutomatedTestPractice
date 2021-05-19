using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Xunit;
using System;
using System.Linq;

using System.IO;
using CSVHeader;
using CsvHelper;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;


namespace UnitTest1
{
    public class UnitTest1:TOOL
    {
        readonly string test_url = "https://www.esunbank.com.tw/bank/about/services/customer/message-board";
        private readonly string Version = "PC";
        public string _name ="";
        public string _id ="";
        public string _phone ="";
        public string _email ="";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void NAvigateToDpApp(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);


            using var driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                driver.Manage().Window.Maximize();


                for (int i = 2; i <= 22; i++) // initial i =2
                {
                    string timepath = System.DateTime.Now.ToString("yyyyMMddhhmmss");
                    int n = i - 1; // 第 n 個縣市

                    int j = 1; // initial j = 1
                    int[] arrray = new int[] { 1, 35, 32, 11, 3, 3, 3, 14, 2, 1, 1, 2, 1, 10, 14, 2, 1, 1, 1, 1, 15 }; //縣市對應分行數
                    while (j <= arrray[i - 2])
                    {
                        if (i == 22 && j == 14)
                        {
                            driver.Navigate().Refresh();
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000);
                        }
                        // 定義欄位XPath
                        String Country_DropDownList = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span"; // 縣市下拉選單XPath
                        String BranchName_DropDownList = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span"; // 分行下拉選單XPath
                        String Cuntry_Xpath = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[" + i + "]/span";
                        String BranchName_Xpath = "//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[" + j + "]/span";

                        // 所有欄位定位
                        IWebElement FullNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[1]/td[2]/input"));
                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath("//*[@id='citizenId']"));
                        IWebElement TelephoneColumn = driver.FindElement(By.XPath("//*[@id='phone']"));
                        IWebElement EMailColumn = driver.FindElement(By.XPath("//*[@id='email']"));
                        IWebElement BusinessItem_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/span"));
                        IWebElement CreditCardBusiness = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/ul/li[2]/span"));
                        IWebElement MessageArea = driver.FindElement(By.XPath("//*[@id='layout_0_maincontent_2_comments']"));
                        IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]/a"));
                        IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));

                        
                        using (var reader = new StreamReader(@"C:\Users\axn01\source\repos\XUnitAutoTest\UserData\UserInfo.csv"))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            
                            //csv.Configuration.HasHeaderRecord = false;

                            var records = csv.GetRecords<UserDataList>().ToList();
                            int k = 1;
                            foreach (var user in records)
                            {
                                if (k == 2) // 抓 CSV 第 k 行資料
                                {
                                     _name = user.NAME;
                                    _id = user.ID;
                                    _phone = user.PHONE;
                                    _email = user.EMAIL;


                                    break;
                                }
                                k++;
                            }

                            FullNameColumn.Clear();
                            FullNameColumn.SendKeys(_name); //填姓名
                            System.Threading.Thread.Sleep(100);

                            IdentityCardColumn.Clear();
                            IdentityCardColumn.SendKeys(_id); //填身分證
                            System.Threading.Thread.Sleep(100);

                            TelephoneColumn.Clear();
                            TelephoneColumn.SendKeys(_phone); //填連絡電話
                            System.Threading.Thread.Sleep(100);

                            EMailColumn.Clear();
                            EMailColumn.SendKeys(_email); //填 E-MAIL

                            Find_Element(driver, EMailColumn);
                            driver.FindElement(By.XPath(Country_DropDownList)).Click(); // 展開 "請選擇縣市" 下拉選單
                            System.Threading.Thread.Sleep(300);
                            driver.FindElement(By.XPath(Cuntry_Xpath)).Click(); // 點選一個 "縣市" 
                            System.Threading.Thread.Sleep(300);
                            System.Diagnostics.Debug.WriteLine("2. 點選選單第 " + n + " 個縣市");
                            System.Threading.Thread.Sleep(300);
                            


                            Find_Element(driver, EMailColumn);
                            driver.FindElement(By.XPath(BranchName_DropDownList)).Click(); // 展開 "分行" 下拉選單
                            System.Threading.Thread.Sleep(300);
                            driver.FindElement(By.XPath(BranchName_Xpath)).Click(); // 點選一個 "分行"
                            System.Threading.Thread.Sleep(300);
                            System.Diagnostics.Debug.WriteLine("4. 點選選單第 " + j + " 個分行");
                            System.Threading.Thread.Sleep(300);

                           

                            Find_Element(driver, EMailColumn);
                            System.Threading.Thread.Sleep(100);
                            BusinessItem_DropDownList.Click(); //點留言業務下拉選單
                            System.Threading.Thread.Sleep(300);
                            CreditCardBusiness.Click(); //選信用卡業務
                            System.Threading.Thread.Sleep(300);

                            MessageArea.Clear();
                            MessageArea.SendKeys(
                                "[Automation Test on " + browserType + " in " + Version + " ]" + "\r\n"
                                + "表單 on 第 " + n + " 個縣市第 " + j + " 個分行" + "\r\n"
                                //+ "=================================================" + "\r\n"
                                //+ "Total 21 區, 154 分行" + "\r\n"
                                //+ "縣市(地區index)(分行數量):" + "\r\n"
                                //+ "基隆市(1)(1),      臺北市(2)(35),      新北市(3)(32),      桃園市(4)(11),      新竹市(5)(3),      "
                                //+ "新竹縣(6)(3),      苗栗縣(7)(3),        臺中市(8)(14),      彰化縣(9)(2),        南投縣(10)(1)," + "\r\n"
                                //+ "雲林縣(11)(1),    嘉義市(12)(2),      嘉義縣(13)(1),      台南市(14)(10),    高雄市(15)(14),      "
                                //+ "屏東縣(16)(2),    宜蘭縣(17)(1),      花蓮縣(18)(1),      臺東縣(19)(1),      澎湖縣(20)(1),       海外地區(21)(15)."
                                ); // 填 "留言內容"

                            TakeScreenShot("d:\\第 " + n + " 個縣市第 " + j + " 個分行 snapshot_" + browserType + "_" + timepath +".png", driver); // snapshot當下畫面
                            System.Threading.Thread.Sleep(1000);
                           // ElementTakeScreenShot(driver, driver.FindElement(By.XPath("//*[@id='submit']")), "d:\\" + browserType + timepath +".png");
                            System.Threading.Thread.Sleep(3000);

                            //IHaveReadRadioButtin.Click(); // 點 "我已閱讀" radio button
                            // System.Threading.Thread.Sleep(100);
                            // System.Diagnostics.Debug.WriteLine("13. 勾選'我已閱讀'");

                            SCrollToElement(driver, MessageArea);
                            System.Threading.Thread.Sleep(3000);
                            //SubmitButton.Click(); // 點 "送出" 按鈕
                            System.Threading.Thread.Sleep(300);
                            //ElementTakeScreenShot(driver, driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]")), "d:\\" + browserType + timepath + ".png");
                            // WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                            // wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // 等待直到看到通知視窗 

                            // driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // 點通知視窗 "X" 按鈕
                            // System.Threading.Thread.Sleep(100);
                            // System.Diagnostics.Debug.WriteLine("15. 關掉通知視窗");

                            // WebDriverWait wait_to_see_submit_button = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                            // wait_to_see_submit_button.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id='submit']"))); // 等待直到看到"留言區"
                            // System.Threading.Thread.Sleep(100);

                            driver.FindElement(By.XPath("//*[@id='scrollUp']")).Click(); //畫面回到頂端
                            System.Threading.Thread.Sleep(500);
                            System.Diagnostics.Debug.WriteLine("16. 回到網頁頂");
                            System.Threading.Thread.Sleep(100);

                            j++;
                        }
                    }
                    driver.Quit();
                }
            }



        }
    }
}
