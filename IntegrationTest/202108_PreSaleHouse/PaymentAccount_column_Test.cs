using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.Text.RegularExpressions;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.PreSaleHouseTrustInquiryTest
{
    public class 預售屋信託查詢: IntegrationTestBase
    {
        public 預售屋信託查詢(ITestOutputHelper output, Setup testSetup): base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType =typeof (BrowserHelper))]
        public void 預售屋查詢精進_欄位檢核(string browser)
        {
            StartTestCase(browser, "查詢精進_輸入欄位檢核");

            INFO("預售屋信託查詢功能精進, 優化既有官網預售屋信託查詢結果, 針對繳款帳號欄位檢核測試");
            //[InlineData("", "必須填寫")]
            //[InlineData("12345", "最少 6 個字")]
            //[InlineData("中文字", "只可輸入數字")]
            //[InlineData("ＡＢＣＤＥＦ", "只可輸入數字")]
            //[InlineData("ａｂｃｄｅｆ", "只可輸入數字")]
            //[InlineData("abcdef", "只可輸入數字")]
            //[InlineData("ABCDEF", "只可輸入數字")]
            //[InlineData("@#$%^&*", "只可輸入數字")]
            //[InlineData("！＠＃＄％︿", "只可輸入數字")]
            //[InlineData("123abc456DEF", "只可輸入數字")]
            //[InlineData("1234567890123456789", "繳款帳號不存在")]
            //[InlineData("987654321", "企業代碼不存在")]



            //Array存取預測試的字串
            string[] array_payment_account = new string[] { "", "12345", "１２３４５６", "中文字", "ＡＢＣＤＥＦ", "ａｂｃｄｅｆ", "abcdef", "ABCDEF", "@#$%^&*", "！＠＃＄％︿", "123abc456DEF", "987654321", "1234567890123456789" };
            //string[] array_payment_account = new string[] { "", "12345" };

            foreach (string keyin in array_payment_account) // 跑迴圈以逐筆輸入欄位
            {

                IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
                PaymentAccount.Clear(); // 清除繳款欄位
                PaymentAccount.SendKeys(keyin); // 輸入值

               

                string actualPaymentAccount = PaymentAccount.GetAttribute("value"); // 輸入欄位取值
                

                IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));


                bool digital_check_result = Regex.IsMatch(actualPaymentAccount, @"^[0-9]*$"); // 判斷輸入字串是否為"全數字"
                bool expect_check_result = Regex.IsMatch(actualPaymentAccount, @"^\d{6,16}$"); // 判斷輸入字串是否為"6~16位全數字"


                ///<summary>
                /// 判斷當"輸入欄位為6~16位全數字" >>> 送出data後延遲等待5秒, 等網頁load完
                /// </summary>
                if (expect_check_result == true) 
                {
                    submit_button.Click(); // 點 送出
                    System.Threading.Thread.Sleep(6000);
                }
                else
                {
                    submit_button.Click(); // 點 送出
                    System.Threading.Thread.Sleep(100);
                }

                if (actualPaymentAccount == "") 
                {
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // 欄位旁error message
                    if (payment_account_error == "必須填寫")
                    {
                        PASS($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { payment_account_error}");
                        result = true;
                    }
                    else
                    {
                        FAIL($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { payment_account_error}");
                        result = false;
                    }
                }
                else if (digital_check_result == true && expect_check_result == false)
                {
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // 欄位旁error message
                    if (payment_account_error == "最少 6 個字")
                    {
                        PASS($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { payment_account_error}");
                        result = true;
                    }
                    else
                    {
                        FAIL($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { payment_account_error}");
                        result = false;
                    }
                }
                else if (digital_check_result == false)
                {
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // 欄位旁error message
                    if (payment_account_error == "只可輸入數字")
                    {
                        PASS($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { payment_account_error}");
                        result = true;
                    }
                    else
                    {
                        FAIL($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { payment_account_error}");
                        result = false;
                    }
                }
                else if (expect_check_result == true)
                {
                    string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 查詢結果錯誤訊息
                    if (searchresult == "繳款帳號不存在")
                    {
                        PASS($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { searchresult }");
                        result = true;
                    }
                    else
                    {
                        FAIL($"輸入: { keyin}, 欄位顯示: { actualPaymentAccount}, 顯示訊息: { searchresult }");
                        result = false;
                    }
                }
            
            }
            PASS(TestBase.ElementSnapShotToReport(driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table"))));
            CloseBrowser();
        }  
     }
}
