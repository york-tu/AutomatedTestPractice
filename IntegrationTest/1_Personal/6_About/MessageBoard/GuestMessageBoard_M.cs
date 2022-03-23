using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.IO;
using System.Linq;
using CsvHelper;
using System.Globalization;
using Xunit.Abstractions;


namespace AutomatedTest.IntegrationTest.About.MessageBoard
{
    public class 訪客留言版縣市資料源調整_PC:IntegrationTestBase
    {
        public 訪客留言版縣市資料源調整_PC (ITestOutputHelper output, Setup testsetup): base(output, testsetup)
        {
            testurl = domain + "https://easyfee.esunbank.com.tw/index.action";
        }
            
        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void 縣市分行選單檢核(string browser)
        {
            StartTestCase(browser, "縣市分行選單檢核", "York");
            INFO("縣市分行欄位Snapshot + 出Report");

            var aaa = driver.FindElement(By.CssSelector(".page_frame3")).Text;



            int country_dropdownList_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li")).Count;  //獲取 "縣市"下拉選單裡的縣市數

            for (int country_index = 1; country_index <= country_dropdownList_amount - 1; country_index++) // 初始值 = 1
            {
               

                IWebElement CountryDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span")); //  "縣市" 下拉選單
                CountryDropDownList.Click(); // 展開縣市選單

                string cuntry_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{country_index + 1}]/span"; // 縣市Xpath
                IWebElement SelectCountry = driver.FindElement(By.XPath(cuntry_xpath)); //  "縣市" 下拉選單
                SelectCountry.Click(); //點選第country_index個縣市

                int branch_dropdownlist_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li")).Count;  //獲取 "當下分行"下拉選單裡的分行數

                for (int branch_index = 1; branch_index <= branch_dropdownlist_amount; branch_index++) // 初始值 = 1
                {
                    if (country_index == country_dropdownList_amount - 1 && branch_index == 14) 
                    {
                        branch_index++;
                    }

                    IWebElement BranchDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span")); //  "分行" 下拉選單
                    BranchDropDownList.Click(); // 展開分行選單

                    string branch_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{branch_index}]/span";
                    IWebElement SelectBranch = driver.FindElement(By.XPath(branch_xpath)); //  "分行" 下拉選單
                    SelectBranch.Click(); //點選第branch_index個縣市


                    // 所有欄位定位
                    IWebElement FullNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[1]/td[2]/input"));
                    IWebElement TelephoneColumn = driver.FindElement(By.XPath("//*[@id='phone']"));
                    IWebElement IdentityCardColumn = driver.FindElement(By.XPath("//*[@id='citizenId']"));
                    IWebElement EMailColumn = driver.FindElement(By.XPath("//*[@id='email']"));
                    IWebElement BusinessItem_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/span"));
                    IWebElement CreditCardBusiness = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/ul/li[2]/span"));
                    IWebElement MessageArea = driver.FindElement(By.XPath("//*[@id='layout_0_maincontent_2_comments']"));
                    IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]/a"));
                    IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));



                    ///<summary>
                    ///從Excel檔中讀出UserData
                    ///</summary>
                    string csvpath = $@"{UserDataList.Upperfolderpath}\testdata\UserInfo.csv";
                    using (var reader = new StreamReader(csvpath))
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        var records = csv.GetRecords<UserDataList>().ToList();
                        Random csv_row = new Random();
                        int user_data_index = 1;
                        foreach (var user in records)
                        {
                            if (user_data_index == csv_row.Next(1, 7)) // 抓 CSV 第 random行資料
                            {
                                FullNameColumn.Clear();
                                FullNameColumn.SendKeys(user.NAME); //填姓名

                                IdentityCardColumn.Clear();
                                IdentityCardColumn.SendKeys(user.ID); //填身分證

                                TelephoneColumn.Clear();
                                TelephoneColumn.SendKeys(user.PHONE); //填連絡電話

                                EMailColumn.Clear();
                                EMailColumn.SendKeys(user.EMAIL); //填 E-MAIL
                                
                                break;
                            }
                            user_data_index++;
                        }

                        string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\GuestMessageBoard";
                        TestBase.CreateFolder(snapshotpath);
                        WARNING(TestBase.ElementSnapShotToReport(driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]"))));

                        //Tools.ElementTakeScreenShot(driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]")),
                        //    $@"{snapshotpath}\第 {country_index} 個縣市第 {branch_index} 個分行_欄位snapshot.png"); //元素截圖
                        //System.Threading.Thread.Sleep(100);

                        //BusinessItem_DropDownList.Click(); //點留言業務下拉選單
                        //CreditCardBusiness.Click(); //選信用卡業務


                        //MessageArea.Clear();
                        //MessageArea.SendKeys(
                        //    "[Automated Testing on " + browser + " in " + Version + " ]" + "\r\n"
                        //    + "第 " + country_index + " 個縣市第 " + branch_index + " 個分行" + "\r\n"
                        //    + "=================================================" + "\r\n"
                        //    + "Total 21 區, 154 分行" + "\r\n"
                        //    + "縣市(地區index)(分行數量):" + "\r\n"
                        //    + "基隆市(1)(1),      臺北市(2)(35),      新北市(3)(32),      桃園市(4)(11),      新竹市(5)(3),      "
                        //    + "新竹縣(6)(3),      苗栗縣(7)(3),        臺中市(8)(14),      彰化縣(9)(2),        南投縣(10)(1)," + "\r\n"
                        //    + "雲林縣(11)(1),    嘉義市(12)(2),      嘉義縣(13)(1),      台南市(14)(10),    高雄市(15)(14),      "
                        //    + "屏東縣(16)(2),    宜蘭縣(17)(1),      花蓮縣(18)(1),      臺東縣(19)(1),      澎湖縣(20)(1),       海外地區(21)(15)."
                        //    ); // 填 "留言內容"

                        //Tools.SCrollToElement(driver, TelephoneColumn);
                        //Tools.TakeScreenShot($@"{snapshotpath}\第 {country_index} 個縣市第 {branch_index} 個分行 snapshot_{browser}.png", driver); // snapshot當下畫面

                        // IHaveReadRadioButtin.Click(); // 點 "我已閱讀" radio button

                        // SubmitButton.Click(); // 點 "送出" 按鈕
                        // System.Threading.Thread.Sleep(300);

                        // ElementTakeScreenShot(driver, driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]")), "d:\\" + browserType + timepath + ".png");
                        
                        // WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                        // wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // 等待直到看到通知視窗 

                        // driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // 點通知視窗 "X" 按鈕
                        // System.Threading.Thread.Sleep(100);
                        // System.Diagnostics.Debug.WriteLine("15. 關掉通知視窗");

                        // WebDriverWait wait_to_see_submit_button = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                        // wait_to_see_submit_button.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id='submit']"))); // 等待直到看到"留言區"
                        // System.Threading.Thread.Sleep(100);

                       // driver.FindElement(By.XPath("//*[@id='scrollUp']")).Click(); //畫面回到頂端
                       // System.Threading.Thread.Sleep(500);
                    }
                }
            }
            CloseBrowser();
        }
    }
}
    

