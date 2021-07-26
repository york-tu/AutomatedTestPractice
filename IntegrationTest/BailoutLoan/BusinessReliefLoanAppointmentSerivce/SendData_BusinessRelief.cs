using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using OpenQA.Selenium.Support.UI;
using System.IO;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.BailoutLoan
{
    public class 企業紓困貸款預約服務_送出資料:IntegrationTestBase
    {
        public 企業紓困貸款預約服務_送出資料(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/small-business/tools/apply/sbloan-appointment";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void 送出資料(string browser)
        {
            StartTestCase(browser, "企業紓困貸款預約服務_所有分行資料送出", "York");

            Tools.CreateSnapshotFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha");
            System.Threading.Thread.Sleep(100);
            Tools.CleanUPFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha"); //清空captcha資料夾

            int country_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li/ul/li")).Count; // 取縣市選單total縣市數
            for (int country_index = 1; country_index <= country_amount - 1; country_index++)
            {
                IWebElement Country_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li"));
                Country_DropDownList.Click(); // 展開 "請選擇縣市" 下拉選單

                string country_xpath = $"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li[{country_index + 1}]";
                IWebElement SelectCountry = driver.FindElement(By.XPath(country_xpath));
                SelectCountry.Click(); // 點選一個"縣市"

                int branch_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li")).Count;
                for (int branch_index = 1; branch_index <= branch_amount - 1; branch_index++) // index 1 = 第一個分行
                {
                    if (branch_index >= 2)
                    {
                        driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li")).Click(); // 重新點縣市選單
                        driver.FindElement(By.XPath(country_xpath)).Click(); // 重新選縣市
                    }
                    IWebElement Branch_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li[2]/span")); // 展開 "分行" 下拉選單
                    Branch_DropDownList.Click(); // 展開"分行"下拉選單
                    string branch_xpath = $"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li[{branch_index + 1}]/span";
                    IWebElement SelectBranch = driver.FindElement(By.XPath(branch_xpath));
                    SelectBranch.Click(); //點選一個"分行



                    /// <summary>
                    /// 填公司名稱
                    ///  </summary>
                    IWebElement CompanyNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]/input"));
                    CompanyNameColumn.Clear();
                    CompanyNameColumn.SendKeys($"AutomatedCompany {country_index}-{branch_index}"); // 填公司名稱



                    /// <summary>
                    /// 填公司統編
                    ///  </summary>
                    IWebElement CompanyUniformNumbersColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]/input"));
                    CompanyNameColumn.Clear();
                    CompanyNameColumn.SendKeys(Tools.CreateUniformNumber()); // 填公司統編



                    /// <summary>
                    /// 填負責人 聯絡人
                    ///  </summary>
                    IWebElement PrincipleColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[2]/input"));
                    IWebElement ContactPersonColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[4]/td[2]/input"));
                    PrincipleColumn.Clear();
                    PrincipleColumn.SendKeys($"Boss");
                    ContactPersonColumn.Clear();
                    ContactPersonColumn.SendKeys($"LittleBrother");



                    /// <summary>
                    /// 填連絡電話: 區碼+主碼+分機
                    ///  </summary>
                    IWebElement TelephoneAreaCodeColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[1]"));
                    IWebElement TelephoneMainColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[2]"));
                    IWebElement TelephoneExtensionColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[3]"));
                    TelephoneAreaCodeColumn.Clear();
                    TelephoneAreaCodeColumn.SendKeys($"02");
                    TelephoneMainColumn.Clear();
                    TelephoneMainColumn.SendKeys($"8825252");
                    TelephoneExtensionColumn.Clear();
                    TelephoneExtensionColumn.SendKeys($"9527");



                    /// <summary>
                    /// 填行動電話
                    ///  </summary>
                    IWebElement CellPhoneColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[1]"));
                    CellPhoneColumn.Clear();
                    CellPhoneColumn.SendKeys(Tools.CreateCellPhoneNumber());



                    ///<summary>
                    /// 洽詢紓困專案下拉選單
                    ///</summary>
                    IWebElement RescueProjectDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li")); // 紓困專案下拉選單
                    RescueProjectDropDownList.Click();
                    int rescueProjectCounts = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li")).Count; // 選項數量
                    Random rescueOption = new Random();
                    int projectSelected = rescueOption.Next(2, rescueProjectCounts + 1);
                    string selectRescueProjectXPath = $"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li[{projectSelected}]";
                    IWebElement SelectResuceProject = driver.FindElement(By.XPath(selectRescueProjectXPath));
                    SelectResuceProject.Click(); // 隨機選擇紓困專案



                    ///<summary>
                    /// 隨機選擇蒞行日期&時間
                    ///</summary>
                    IWebElement DateDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li"));
                    IWebElement TimeDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li"));
                    Tools.ScrollPageUpOrDown(driver, 700);
                    int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li")).Count; // 獲取蒞行日期下拉選單選項數量
                    int time_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li")).Count; // 獲取蒞行時間下拉選單選項數量

                    Random ran_date = new Random();
                    Random ran_time = new Random();
                    int ran_date_index = ran_date.Next(2, date_amount);
                    int ran_time_index = ran_time.Next(2, time_amount);

                    DateDropDownList.Click();  // 展開"蒞行日期"下拉選單
                    IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li[{ran_date_index}]/span"));
                    SelectDate.Click(); // 隨機點選一個"日期"
                    TimeDropDownList.Click(); // 展開"時段"下拉選單
                    IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li[{ran_time_index}]/span"));
                    SelectTime.Click(); // 隨機點選一個"時段"



                    ///<summary>
                    /// 點 "我已閱讀並同意"
                    ///</summary>
                    IWebElement IHaveReadRadioButton = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[11]/td[2]/div/a"));
                    IHaveReadRadioButton.Click();



                    ///<summary>
                    /// 點"確認"
                    ///</summary>
                    IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));
                    SubmitButton.Click(); // 點 送出



                    ///<summary>
                    /// 自動輸入圖片驗證碼
                    ///</summary>
                    int verify_count = 1; // verify_count: 紀錄retry驗證碼次數
                retryagain:
                    IWebElement CaptchaPicture = driver.FindElement(By.XPath("//*[@id='ImgCaptcha']")); // Captcha圖片位置
                    IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath("//*[@id='captchaValue']")); // 輸入驗證碼欄位

                    Tools.ElementSnapshotshot(CaptchaPicture, $@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //snapshot驗證碼圖片
                    string verify_code_result = TestBase.TesseractOCRIdentify($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png", 0.75); //解析出驗證碼

                    if (verify_count >= 10) // 依序刪除舊的picture
                    {
                        File.Delete($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count - 9}.png");
                    }
                    ImageVerificationCodeColumn.Clear();
                    ImageVerificationCodeColumn.SendKeys(verify_code_result); // 輸入驗證碼

                    if (driver.FindElement(By.Id("captchaWrong")).Text == "") // 檢查驗證碼錯誤訊息欄位是否為空
                    {
                        SubmitButton.Click(); // 點"確認"button

                        if (driver.FindElement(By.Id("captchaWrong")).Text == "")
                        {
                            goto gotoneststep;
                        }
                        else
                        {
                            CaptchaPicture.Click();
                            verify_count++;
                            goto retryagain;
                        }
                    }
                    else
                    {
                        CaptchaPicture.Click();
                        verify_count++;
                        goto retryagain;
                    }

                gotoneststep:
                    WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                    wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // 等待直到看到通知視窗 
                    System.Threading.Thread.Sleep(1000);
                    Tools.ScrollPageUpOrDown(driver, 700);
                    string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\BailoutLoan";
                    Tools.CreateSnapshotFolder(snapshotpath);
                    Tools.FullScreenshot($@"{snapshotpath}\第{country_index}縣市第{branch_index}分行_紓困專案{projectSelected - 1}_預約第{ran_date_index - 1}日第{ran_time_index - 1}時段.png"); //實作截圖

                    driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // 點通知視窗 "X" 按鈕
                    System.Threading.Thread.Sleep(1000);

                    driver.Navigate().GoToUrl(testurl);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);

                }
            }
            driver.Quit();
        }
    }
}


