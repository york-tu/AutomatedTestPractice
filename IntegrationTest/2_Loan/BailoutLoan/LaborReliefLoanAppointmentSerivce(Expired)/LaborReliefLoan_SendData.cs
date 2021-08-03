using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using OpenQA.Selenium.Support.UI;
using System.IO;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Personal.Loan.BailoutLoan
{
    public class 勞工紓困貸款預約服務_送出資料: IntegrationTestBase
    {
        public 勞工紓困貸款預約服務_送出資料(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/labor-loan-appointment#";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void 送出資料(string browser)
        {
            StartTestCase(browser, "勞工紓困貸款預約服務_所有分行資料送出", "York");

            Tools.CreateSnapshotFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha");
            System.Threading.Thread.Sleep(100);
            Tools.CleanUPFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha"); //清空captcha資料夾

            for (int i = 2; i <= 21; i++) // initial i =2
                {

                    int j = 1; // initial j = 1
                    int[] arrray = new int[] { 35, 1, 31, 1, 3, 3, 11, 3, 14, 2, 1, 2, 1, 1, 10, 14, 1, 2, 1, 1 }; //縣市對應分行數
                    while (j <= arrray[i - 2])
                    {
                        string timesavepath = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");

                        // 定義欄位XPath
                        string Cuntry_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{i}]/span";
                        string Branch_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{j + 1}]/span";


                        IWebElement FullNameColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.name_column_Xpath()));
                        FullNameColumn.Clear();
                        FullNameColumn.SendKeys($"機器人{i - 1}_{j}"); // 填姓名


                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.ID_column_XPath()));
                        IdentityCardColumn.Clear();
                        bool sex = false;
                        if (j % 2 == 0)
                            sex = true;
                        IdentityCardColumn.SendKeys(Tools.CreateIDNumber(sex, j % 21)); // 引用tools內身分證字號產生器 & 填身分證欄位


                        IWebElement CellPhoneColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.cellphone_column_XPath()));
                        CellPhoneColumn.Clear();
                        CellPhoneColumn.SendKeys(Tools.CreateCellPhoneNumber()); // 引用tools內電話號碼產生器 & 填電話號碼欄位


                        IWebElement BirthdayCalendarIcon = driver.FindElement(By.Id("datepicker1-button"));
                        BirthdayCalendarIcon.Click(); // 點開出生年月日小日曆 
                        Random Year = new Random();
                        Random Month = new Random();
                        // IWebElement Year_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]"));
                        IWebElement SelectYear = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[1]/select/option[{Year.Next(1, 121)}]")); // 隨機年份
                        SelectYear.Click();
                        //  IWebElement Month_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]"));
                        IWebElement SelectMonth = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[2]/select/option[{Month.Next(1, 12)}]")); // 隨機月份
                        SelectMonth.Click(); // 選日期

                    rerun:
                        Random week = new Random();
                        Random day = new Random();
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr[{week.Next(1, 5)}]/td[{day.Next(1, 7)}]")); // 隨機日期
                        string date_value = date.GetAttribute("class");
                        if (date_value == "is-empty") // 判斷當欄位為空時(表示當月沒有這日子), 重新產生隨機日期
                        {
                            goto rerun;
                        }
                        else
                        {
                            date.Click();
                        }

                        IWebElement Country_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.country_dropdownlist_XPath()));
                        Country_DropDownList.Click(); //展開"請選擇縣市"下拉選單

                        IWebElement SelectCountry = driver.FindElement(By.XPath(Cuntry_Xpath));
                        SelectCountry.Click(); // 點選一個"縣市"

                        IWebElement Branch_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.branch_dropdownlist_XPath()));
                        Branch_DropDownList.Click(); // 展開"分行"下拉選單

                        IWebElement SelectBranch = driver.FindElement(By.XPath(Branch_Xpath));
                        SelectBranch.Click(); //點選一個"分行"


                        IWebElement Date_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.date_dropdownlist_XPath()));
                        int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li")).Count;
                        int time_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li")).Count;
                        Random ran_date = new Random();
                        Random ran_time = new Random();

                        Date_DropDownList.Click(); //展開"蒞行日期"下拉選單
                        IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li[{ran_date.Next(2, date_amount)}]/span"));
                        SelectDate.Click(); // 循環點選一個"日期"
                        IWebElement Time_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.time_dropdownlist_XPath()));
                        Time_DropDownList.Click(); // 展開"時段"下拉選單
                        IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li[{ran_time.Next(2, time_amount)}]/span"));
                        SelectTime.Click(); // 循環點選一個"時段"

                        IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath(LaborReliefLoan_XPath.i_have_read_button_XPath()));
                        IHaveReadRadioButtin.Click(); // 點"我已閱讀" radio button

                        IWebElement SubmitButton = driver.FindElement(By.XPath(LaborReliefLoan_XPath.submit_button_XPath()));

                        IWebElement CaptchaPicture = driver.FindElement(By.XPath("//*[@id='ImgCaptcha']")); //圖片欄位

                       

                        int verify_count = 1; // verify_count: 紀錄retry驗證碼次數
                    retryagain:
                        IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.image_verify_code_column_XPath())); // 輸入驗證碼欄位

                        Tools.SCrollToElement(driver, FullNameColumn);
                        Tools.ElementSnapshotshot(CaptchaPicture, $@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //snapshot驗證碼圖片

                        if (verify_count >=10) // 依序刪除舊的picture
                        {
                            File.Delete($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count-9}.png");
                        }

                        string verify_code_result = TestBase.TesseractOCRIdentify($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png", 0.75); //解析出驗證碼

                        // Tools.ElementTakeScreenShot(CaptchaPicture, $@"{snapshotpath}\ImageVerifyCode.png");
                        // string verify_code_result = Tools.IronOCR($@"{snapshotpath}\ImageVerifyCode.png"); //解析出驗證碼: Iron_OCR
                        // string verify_code_result = Tools.BaiduOCR($@"{snapshotpath}\ImageVerifyCode.png"); //解析出驗證碼: Baidu_OCR

                        ImageVerificationCodeColumn.Clear();
                        ImageVerificationCodeColumn.SendKeys(verify_code_result); // 輸入驗證碼
                        if (driver.FindElement(By.Id("captchaWrong")).Text =="") // 檢查驗證碼錯誤訊息欄位是否為空
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

                        WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                        wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // 等待直到看到通知視窗 
                        System.Threading.Thread.Sleep(300);

                        Tools.CreateSnapshotFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\LaborReliefLoan");
                        Tools.PageSnapshot($@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\LaborReliefLoan\第{i - 1}縣市第{j}分行_申請_第{ran_date}日第{ran_time}時段.png", driver); //實作截圖


                        driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // 點通知視窗 "X" 按鈕
                        System.Threading.Thread.Sleep(3000);

                        driver.Navigate().GoToUrl(testurl);
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(100);

                        j++;
                    }
                }
            CloseBrowser();
            }
        }
    }


