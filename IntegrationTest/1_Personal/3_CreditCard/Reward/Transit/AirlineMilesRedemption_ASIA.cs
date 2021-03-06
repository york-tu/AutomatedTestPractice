using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using Xunit;
using System;
using System.IO;
using System.Globalization;
using System.Linq;
using IronXL;
using CsvHelper;
using AutomatedTest.Utilities;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Personal.CreditCard.Reward.Transit
{
    public class 航空里程兌換精進_亞洲萬里通_M版:IntegrationTestBase
    {
        private readonly string Version = "Mobile";
        private readonly string testcase_name = "亞洲萬里通";

        public 航空里程兌換精進_亞洲萬里通_M版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/credit-card/reward/transit/asiamiles?dev=mobile";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void 航空里程兌換精進_欄位檢核(string browser)
        {
            string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");
            string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\AirlineMilesRedemption";
            string excel_savepath = snapshotpath + "\\TestReport.xlsx";

            WorkBook xlsWorkbook;
            WorkSheet xlsSheet;

            if (File.Exists(excel_savepath) == true && WorkBook.Load(excel_savepath).GetWorkSheet(testcase_name) == null)
            { // 判斷當excel檔存在 但 沒有"xxx"工作表 >>> 讀取該excel檔, new create "xxx" 工作表 
                xlsWorkbook = WorkBook.Load(excel_savepath);
                xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name);
            }
            else if (File.Exists(excel_savepath) == true && WorkBook.Load(excel_savepath).GetWorkSheet(testcase_name) != null)
            { // 判斷當excel檔存在 且 有"xxx"工作表 >>> 讀取該excel檔, 讀取"xxx" 工作表
                xlsWorkbook = WorkBook.Load(excel_savepath);
                xlsSheet = xlsWorkbook.GetWorkSheet(testcase_name);
            }
            else
            { // 判斷當excel檔不存在 >>> new create excel檔 & new create "xxx" 工作表
                xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLSX); //定義 excel格式為 XLSX
                xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name);
            }


            IWebElement go_to_redeem = driver.FindElement(By.XPath("//*[@id='layout_m_0_content_m_3_tab_0_HlkMeleageForm']"));
            go_to_redeem.Click(); // 點"前往兌換"
            System.Threading.Thread.Sleep(100);

            WebDriverWait redeem_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
            redeem_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id='login']"))); //等待直到看到 會員登入 dialog

            IWebElement creditcard_friend_login = driver.FindElement(By.XPath("//*[@id='layout_m_0_content_m_2_HlkCardLogin']/img"));
            creditcard_friend_login.Click();// 點"信用卡友登入"
            System.Threading.Thread.Sleep(100);

            string csvpath = $@"{UserDataList.Upperfolderpath}\testdata\UserInfo.csv";
            using (var reader = new StreamReader(csvpath)) //讀CSV檔
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                var records = csv.GetRecords<UserDataList>().ToList();
                int k = 1;
                foreach (var user in records)
                {
                    if (k == 6)
                    {
                        IWebElement ID = driver.FindElement(By.XPath("//*[@id='CHID_T']"));
                        IWebElement CreditCard = driver.FindElement(By.XPath("//*[@id='cardNO_T']"));
                        IWebElement Expire_Mon = driver.FindElement(By.XPath("//*[@id='endM']"));
                        IWebElement Expire_Year = driver.FindElement(By.XPath("//*[@id='endY']"));
                        IWebElement CheckCode = driver.FindElement(By.XPath("//*[@id='cardCVV_T']"));
                        IWebElement Submit_button = driver.FindElement(By.XPath("//*[@id='confirm_btn']"));

                        ID.Clear();
                        ID.SendKeys(user.ID); // 輸入身分證字號
                        System.Threading.Thread.Sleep(100);

                        CreditCard.Clear();
                        CreditCard.SendKeys(user.CARDID_FULL); // 輸入卡號
                        System.Threading.Thread.Sleep(100);

                        Expire_Mon.Clear();
                        Expire_Mon.SendKeys(user.MON); // 輸入有效月份
                        System.Threading.Thread.Sleep(100);

                        Expire_Year.Clear();
                        Expire_Year.SendKeys(user.YEAR); // 輸入有效年分
                        System.Threading.Thread.Sleep(100);

                        CheckCode.Clear();
                        CheckCode.SendKeys(user.CHECKCODE); //輸入驗證碼
                        System.Threading.Thread.Sleep(100);

                        Submit_button.Click(); // 送出
                        System.Threading.Thread.Sleep(100);

                        break;
                    }
                    k++;
                }
            }

            IWebElement redeem_again = driver.FindElement(By.XPath(""));
            redeem_again.Click(); // 前往兌換
            System.Threading.Thread.Sleep(100);
            IWebElement submit_redeem_setting = driver.FindElement(By.XPath(""));
            submit_redeem_setting.Click(); //點送出

            TestBase.PageSnapshot(driver,$@"{snapshotpath}\{testcase_name}_{browser}_{Version}_{timesavepath}.png");
            System.Threading.Thread.Sleep(500);


            //檢查點數輸入欄位
            IWebElement airline_input_point = driver.FindElement(By.XPath(""));
            string[] array_asia_point = new string[] { "１２３", "５６", "Ｂ", "@", "G", "0", "6666", "1", "20", "50", "999", "中文字", "@#$*", "！＠", "1a4", "" };

            //定義excel欄位
            xlsSheet["A1"].Value = "檢查 " + testcase_name + " '點數 & 換算' 欄位";
            xlsSheet["B1"].Value = "顯示訊息";
            xlsSheet["C1"].Value = "結果";
            xlsSheet["A1:C1"].Style.BottomBorder.SetColor("#ff6600"); // 底線"紅色"
            xlsSheet["A1:C1"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double; //加雙底線


            int i = 3;
            foreach (var keyin_point in array_asia_point) // 跑迴圈以逐筆輸入欄位
            {
                string check_position = "A" + i;
                string show_msg = "B" + i;
                string test_result = "C" + i;

                airline_input_point.Clear();
                airline_input_point.SendKeys(keyin_point);

                string total_point = driver.FindElement(By.Id("total")).GetAttribute("value");
                string to_mile = driver.FindElement(By.Id("mile")).GetAttribute("value");
                string display_keyin_point = airline_input_point.GetAttribute("value");
                string unit_point = driver.FindElement(By.Id("")).GetAttribute("value");

                submit_redeem_setting.Click();
                string count_error = driver.FindElement(By.Id("counter-error")).Text;

                xlsSheet[check_position].Value = "輸入 " + keyin_point;
                xlsSheet[show_msg].Value = count_error;
                xlsSheet[test_result].Value = " ==> 我要用 " + unit_point + " 點 x " + display_keyin_point + " = " + total_point + " 兌換" + testcase_name + " " + to_mile + " 哩";
                System.Threading.Thread.Sleep(100);
                i++;
            }

            if (File.Exists(excel_savepath) != true)
            {
                xlsWorkbook.SaveAs(excel_savepath);
            }
            else
            {
                xlsWorkbook.Save();
            }
            CloseBrowser();
        }
    }
}
