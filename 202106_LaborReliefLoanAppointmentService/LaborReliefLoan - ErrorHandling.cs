using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Xunit;
using References;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text.RegularExpressions;


namespace LaborReliefLoanAppointmentServiceTest
{
    public class 勞工紓困貸款預約服務_ErrorHandling_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/labor-loan-appointment#";


        [Theory]
        [InlineData(BrowserType.Chrome)]
        [InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                driver.Manage().Window.Maximize();


                for (int i = 2; i <= 21; i++) // initial i =2
                {

                    int j = 1; // initial j = 1
                    int[] arrray = new int[] { 35, 1, 31, 1, 3, 3, 11, 3, 14, 2, 1, 2, 1, 1, 10, 14, 1, 2, 1, 1 }; //縣市對應分行數
                    while (j <= arrray[i - 2])
                    {
                        string timesavepath = System.DateTime.Now.ToString("yyyy-MM-dd_hhmm");

                        // 定義欄位XPath
                        string Cuntry_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{i}]/span";
                        string Branch_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{j + 1}]/span";

                        IWebElement SubmitButton = driver.FindElement(By.XPath(LaborReliefLoan_XPath.submit_button_XPath()));
                        SubmitButton.Click(); // 點"確認"button




                        ///<summary>
                        /// 檢查網頁上hyperlink是否正確
                        ///</summary>
                        string LaborReliefOnlineApplication_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[2]/a")).GetAttribute("href");
                        Assert.Equal("https://www.esunbank.com.tw/s/PersonalLoanApply/" +
                            "Landing/IDConfirm?MKP=eyJNS0RQVCI6bnVsbCwiTUtFSUQiOm51bGwsIk1LUFJOIjoiUDAwMDAwMjEiLCJNS1BKTiI6IkowMDAwMDM0IiwiTUtQSUQiOm51bGx9", 
                            LaborReliefOnlineApplication_hyperlink);
                        
                        string PersonalInformation_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[8]/td[2]/a")).GetAttribute("href");
                        Assert.Equal("https://www.esunbank.com.tw/bank/about/announcement/privacy/privacy-statement", PersonalInformation_hyperlink);




                        ///<summary>
                        /// 姓名欄位 (目前無檢核)
                        ///</summary>
                        IWebElement FullNameColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.name_column_Xpath()));
                        FullNameColumn.Clear();
                        FullNameColumn.SendKeys($"機器人{i - 1}_{j}"); // 填姓名




                        ///<summary>
                        /// 身分證字號欄位檢核
                        ///</summary>
                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.ID_column_XPath()));

                        Random random = new Random();
                        int random_number = random.Next(1, 21); //產生隨機1-21間數字
                        bool sex = false;
                        if (random_number % 2 == 0)
                            sex = true; // for 身分證字號tools 使用(男/女)

                        string[] IDNumbers = new string[] { "", "!@#%^", "許功蓋", "ABCDEF", "0897654321", "A987654321", 
                            "ａ１２３４５６７８９", "ａ123456789", "a12345678９", "A800000014", "AD12544441", "a123456788", "A123456788",
                            "a800000014", "aD12544441", "a123456788", "Ad30341957", "a941062183", "A970000026", "AD30341957", "A941062183", 
                            $"{Tools.CreateRandomString(10)}", $"{Tools.CreateIDNumber(sex, random_number)}" };
                        // 倒數第二組為隨機長度10大小寫英文+數字組合
                        // 最後一組透過身分證字號產生器產生符合規格字號

                        foreach (var input in IDNumbers)
                        {
                            IdentityCardColumn.Clear();
                            bool ResidentIDNumberCheck = Regex.IsMatch(input.ToUpper(), @"^[A-Z]{1}[0-9]{9}$"); // 使用正則表示式: 檢驗格式 [A~Z] {1}個數字 +  [0~9] {9}個數字
                            bool ForeignerIDNumberCheck = Regex.IsMatch(input.ToUpper(), @"^[A-Z]{1}[A-D8-9]{1}[0-9]{8}$");

                            IdentityCardColumn.SendKeys(input);
                            string id_error = driver.FindElement(By.Id("citizenId-error")).Text;
                            if (input == "")
                            {
                                Assert.Equal("必須填寫", id_error);
                            }
                            else if (ResidentIDNumberCheck != true && ForeignerIDNumberCheck != true)
                            {
                                Assert.Equal("請輸入有效的身分證字號", id_error);
                            }
                            else if (ResidentIDNumberCheck == true && Tools.CheckResidentID(input.ToUpper()) != true)
                            {
                                Assert.Equal("請輸入有效的身分證字號", id_error);
                            }
                            else if (ForeignerIDNumberCheck == true && Tools.CheckForeignerID(input.ToUpper()) != true)
                            {
                                Assert.Equal("請輸入有效的身分證字號", id_error);
                            }
                            else if (ForeignerIDNumberCheck == true && Tools.CheckForeignerID(input.ToUpper()) == true)
                            {
                                Assert.Equal("本服務目前僅限本國人申請", id_error);
                            }
                            else
                                Assert.Equal("", id_error);
                        }




                        ///<summary>
                        /// 電話號碼欄位檢核
                        ///</summary>
                        IWebElement CellPhoneColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.cellphone_column_XPath()));
                        string[] phones = new string[] { "ABCDEF", "0897654321", "", $"{Tools.CreateRandomNumber(10)}", $"{Tools.CreateCellPhoneNumber()}" };
                        // 倒數第二組: 隨機產生長度10數字
                        // 最後一組: 透過行動電話號碼產生器產生符合規格號碼
                        foreach (var input in phones)
                        {
                            int string_length = System.Text.Encoding.Default.GetBytes(input).Length; //  擷取字串長度bytes (UTF-8標準, 半形英數字 = 1 byte, 中文&全形英數字 = 3 bytes)

                            CellPhoneColumn.Clear();
                            bool CellPhoneNumberCheck = Regex.IsMatch(input, @"^09\d{8}$"); // 正則表示式: 09開頭後面8碼數字
                            CellPhoneColumn.SendKeys(input);
                            string cellphone_error = driver.FindElement(By.Id("cellPhone-error")).Text;
                            if (input == "")
                            {
                                Assert.Equal("必須填寫", cellphone_error);
                            }
                            else if (CellPhoneNumberCheck != true)
                            {
                                Assert.Equal("行動電話格式錯誤", cellphone_error);
                            }
                            else if (CellPhoneNumberCheck == true && string_length != 10) // 10位全半形數字 = 10 bytes長度
                            {
                                Assert.Equal("行動電話格式錯誤", cellphone_error);
                            }
                            else
                            {
                                Assert.Equal("", cellphone_error);
                            }
                        }




                        ///<summary>
                        /// 檢查"我已閱讀並同意" 錯誤訊息
                        ///</summary>
                        IWebElement IHaveReadRadioButton = driver.FindElement(By.XPath(LaborReliefLoan_XPath.i_have_read_button_XPath()));

                    rerun_ihaveread:
                        string i_have_read_status = IHaveReadRadioButton.GetAttribute("class");
                        string i_have_read_error = driver.FindElement(By.Id("apply-error")).Text;
                        if (i_have_read_status != "checked")
                        {
                            Assert.Equal("必須填寫", i_have_read_error);
                            IHaveReadRadioButton.Click(); // 點"我已閱讀" radio button
                            goto rerun_ihaveread;
                        }
                        else
                        {
                            Assert.Equal("", i_have_read_error);
                        }




                        ///<summary>
                        /// 出生年月日欄位
                        ///</summary>
                        IWebElement BirthdayCalendarIcon = driver.FindElement(By.Id("datepicker1-button"));
                        BirthdayCalendarIcon.Click(); // 點開出生年月日小日曆 
                        Random Year = new Random();
                        Random Month = new Random();
                        // IWebElement Year_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]"));
                        IWebElement SelectYear = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[1]/select/option[{Year.Next(1, 122)}]")); // 隨機年份
                        SelectYear.Click();
                        //  IWebElement Month_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]"));
                        IWebElement SelectMonth = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[2]/select/option[{Month.Next(1, 13)}]")); // 隨機月份
                        SelectMonth.Click(); // 選日期

                    rerun_date:
                        Random week = new Random();
                        Random day = new Random();
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr[{week.Next(1, 6)}]/td[{day.Next(1, 8)}]")); // 隨機日期
                        string date_value = date.GetAttribute("class");
                        if (date_value == "is-empty") // 判斷當欄位為空時(表示當月沒有這日子), 重新產生隨機日期
                        {
                            goto rerun_date;
                        }
                        else
                        {
                            date.Click();
                        }




                        ///<summary>
                        /// 縣市&分行欄位
                        ///</summary>
                        IWebElement Country_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.country_dropdownlist_XPath()));
                        Country_DropDownList.Click(); //展開"請選擇縣市"下拉選單
                        IWebElement SelectCountry = driver.FindElement(By.XPath(Cuntry_Xpath));
                        SelectCountry.Click(); // 點選一個"縣市"
                        IWebElement Branch_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.branch_dropdownlist_XPath()));
                        Branch_DropDownList.Click(); // 展開"分行"下拉選單
                        IWebElement SelectBranch = driver.FindElement(By.XPath(Branch_Xpath));
                        SelectBranch.Click(); //點選一個"分行"

                        IWebElement Date_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.date_dropdownlist_XPath()));
                        int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li")).Count; // 獲取蒞行日期下拉選單選項數量
                        int time_amount = driver.FindElements(By.XPath("//*[@id='mainform'']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li")).Count; // 獲取蒞行時間下拉選單選項數量

                        for (int m = 2; m <= date_amount; m++)
                        {
                            for (int n = 2; n <= time_amount; n++)
                            {
                                
                                Date_DropDownList.Click(); //展開"蒞行日期"下拉選單
                                IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li[{m}]/span"));
                                SelectDate.Click(); // 點選一個"日期"
                                IWebElement Time_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.time_dropdownlist_XPath()));
                                Time_DropDownList.Click(); // 展開"時段"下拉選單
                                IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li[{n}]/span"));
                                SelectTime.Click(); // 點選一個"時段"

                                string ErrorMessage = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/div[1]")).Text;
                                if (ErrorMessage == "") // 當欄位message為空, 即"該分行此時段名額已滿，請選擇其他時段。"
                                {
                                    string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\LaborReliefLoan";
                                    Tools.CreateSnapshotFolder(snapshotpath);
                                    System.Threading.Thread.Sleep(100);
                                    Tools.TakeScreenShot($@"{snapshotpath}\第{i - 1}縣市第{j}分行_第{m-1}日第{n-1}時段.png", driver); //實作截圖
                                }
                            }
                        }




                        ///<summary>
                        /// 檢查 "預約蒞行時間" 說明文字
                        ///</summary>
                        string appointment_visit_time_column = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]")).Text;
                        string appointment_visit_time_information = appointment_visit_time_column.Substring(appointment_visit_time_column.IndexOf("※"));
                        Assert.Equal("※為提升服務品質並進行顧客蒞行服務分流，本服務僅供每人每次預約1個時段，造成不便，請您見諒。", appointment_visit_time_information);




                       ///<summary>
                       /// 檢查 "圖型驗證碼" 錯誤訊息
                       ///</summary>
                    rerun_imageverification:
                        IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.image_verify_code_column_XPath()));
                        string image_verification_column_value = driver.FindElement(By.Id("captchaValue")).GetAttribute("value");

                        if (image_verification_column_value == "")
                        {
                            string image_verification_error = driver.FindElement(By.Id("captchaWrong")).Text;
                            Assert.Equal("請輸入驗證碼", image_verification_error);
                            ImageVerificationCodeColumn.SendKeys("9527");
                            goto rerun_imageverification;
                        }
                        else
                        {
                            SubmitButton.Click();
                            string image_verification_error = driver.FindElement(By.Id("captchaWrong")).Text;
                            Assert.Equal("驗證碼錯誤", image_verification_error);
                        }




                        ///<summary>
                        /// 檢查 "圖型驗證碼" 錯誤訊息
                        ///</summary>
                        string promot_wordings = driver.FindElement(By.XPath("//*[@id='mainform'']/div[9]/div[3]/div[2]/div/div[4]/div")).Text;
                        Assert.Equal("※提醒您，點擊「確認」送出預約後需完成線上申請書填寫，收到簡訊通知後才算成功完成預約服務，" +
                            "並請您於預約時間前往分行臨櫃申辦。若您未收到成功預約簡訊，為配合政府防疫措施，請勿直接至分行申請。", promot_wordings);

                        j++;
                    }
                }
                driver.Quit();
            }
        }
    }
}


