using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using Xunit;
using Utilities;
using System.IO;
using System.Linq;
using CsvHelper;
using System.Globalization;


namespace GuestMessageBoardTest
{
    public class 訪客留言版縣市資料源調整_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/about/services/customer/message-board";
        private readonly string Version = "PC";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                driver.Manage().Window.Maximize();

                int country_dropdownList_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li")).Count;  //獲取 "縣市"下拉選單裡的縣市數
                
                for (int country_index = 1; country_index <= country_dropdownList_amount-1; country_index++) 
                {
                  
                    IWebElement CountryDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span")); //  "縣市" 下拉選單
                    CountryDropDownList.Click(); // 展開縣市選單

                    string cuntry_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{country_index + 1}]/span"; // 縣市Xpath
                    IWebElement SelectCountry = driver.FindElement(By.XPath(cuntry_xpath)); //  "縣市" 下拉選單
                    SelectCountry.Click(); //點選第country_index個縣市

                    int branch_dropdownlist_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li")).Count;  //獲取 "當下分行"下拉選單裡的分行數

                    for (int branch_index = 1; branch_index <= branch_dropdownlist_amount; branch_index++)
                    {
                        if (country_index == country_dropdownList_amount - 1 && branch_index == 14)
                        {
                            driver.Navigate().Refresh();
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000);
                        }


                        IWebElement BranchDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span")); //  "分行" 下拉選單
                        BranchDropDownList.Click(); // 展開分行選單

                        string branch_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{branch_index}]/span";
                        IWebElement SelectBranch = driver.FindElement(By.XPath(branch_xpath)); //  "分行" 下拉選單
                        SelectBranch.Click(); //點選第branch_index個縣市


                        string time = System.DateTime.Now.ToString("yyyyMMddHHmmss");

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


                        string csvpath = $@"{UserDataList.folderpath}\testdata\UserInfo.csv";
                        using (var reader = new StreamReader(csvpath))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            var records = csv.GetRecords<UserDataList>().ToList();
                            int user_data_index = 1;
                            foreach (var user in records)
                            {
                                if (user_data_index == 2) // 抓 CSV 第 user_data_index 行資料
                                {
                                    FullNameColumn.Clear();
                                    FullNameColumn.SendKeys(user.NAME); //填姓名
                                    //System.Threading.Thread.Sleep(100);

                                    IdentityCardColumn.Clear();
                                    IdentityCardColumn.SendKeys(user.ID); //填身分證
                                    //System.Threading.Thread.Sleep(100);

                                    TelephoneColumn.Clear();
                                    TelephoneColumn.SendKeys(user.PHONE); //填連絡電話
                                   // System.Threading.Thread.Sleep(100);

                                    EMailColumn.Clear();
                                    EMailColumn.SendKeys(user.EMAIL); //填 E-MAIL
                                    //System.Threading.Thread.Sleep(100);

                                    break;
                                }
                                user_data_index++;
                            }

                            string snapshotpath = $@"{UserDataList.folderpath}\SnapshotFolder\GuestMessageBoard";
                            Tools.CreateSnapshotFolder(snapshotpath);
                            //System.Threading.Thread.Sleep(100);

                            Tools.ElementTakeScreenShot(driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]")),
                                $@"{snapshotpath}\第 {country_index} 個縣市第 {branch_index} 個分行_欄位snapshot.png"); //元素截圖


                            //System.Threading.Thread.Sleep(100);
                            BusinessItem_DropDownList.Click(); //點留言業務下拉選單
                            //System.Threading.Thread.Sleep(300);
                            CreditCardBusiness.Click(); //選信用卡業務
                            //System.Threading.Thread.Sleep(300);

                            //MessageArea.Clear();
                            //MessageArea.SendKeys(
                            //    "[Automation Test on " + browserType + " in " + Version + " ]" + "\r\n"
                            //    + "表單 on 第 " + country_index + " 個縣市第 " + branch_index + " 個分行" + "\r\n"
                                //+ "=================================================" + "\r\n"
                                //+ "Total 21 區, 154 分行" + "\r\n"
                                //+ "縣市(地區index)(分行數量):" + "\r\n"
                                //+ "基隆市(1)(1),      臺北市(2)(35),      新北市(3)(32),      桃園市(4)(11),      新竹市(5)(3),      "
                                //+ "新竹縣(6)(3),      苗栗縣(7)(3),        臺中市(8)(14),      彰化縣(9)(2),        南投縣(10)(1)," + "\r\n"
                                //+ "雲林縣(11)(1),    嘉義市(12)(2),      嘉義縣(13)(1),      台南市(14)(10),    高雄市(15)(14),      "
                                //+ "屏東縣(16)(2),    宜蘭縣(17)(1),      花蓮縣(18)(1),      臺東縣(19)(1),      澎湖縣(20)(1),       海外地區(21)(15)."
                             //   ); // 填 "留言內容"

                           // Tools.SCrollToElement(driver, TelephoneColumn);
                            //Tools.TakeScreenShot($@"d:\第 {country_index} 個縣市第 {branch_index} 個分行 snapshot_{browserType}.png", driver); // snapshot當下畫面
                           // System.Threading.Thread.Sleep(500);




                            // IHaveReadRadioButtin.Click(); // 點 "我已閱讀" radio button
                            // System.Threading.Thread.Sleep(100);
                            // System.Diagnostics.Debug.WriteLine("13. 勾選'我已閱讀'");

                            // SCrollToElement(driver, MessageArea);
                            // System.Threading.Thread.Sleep(3000);

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

                            //driver.FindElement(By.XPath("//*[@id='scrollUp']")).Click(); //畫面回到頂端
                            //System.Threading.Thread.Sleep(500);




                        }
                    }
                }
                driver.Quit();
            }
        }
    }
}
    

