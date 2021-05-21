using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Xunit;
using System;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;

namespace PreSaleHouseTrustInquiry_PaymentAccount
{
    public class PaymentAccount// 檢款帳號欄位檢查
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";

        [Theory]
        [InlineData("", "必須填寫")]
        [InlineData("12345", "最少 6 個字")]
        [InlineData("中文字", "只可輸入數字")]
        [InlineData("ＡＢＣＤＥＦ", "只可輸入數字")]
        [InlineData("ａｂｃｄｅｆ", "只可輸入數字")]
        [InlineData("abcdef", "只可輸入數字")]
        [InlineData("ABCDEF", "只可輸入數字")]
        [InlineData("@#$%^&*", "只可輸入數字")]
        [InlineData("！＠＃＄％︿", "只可輸入數字")]
        [InlineData("123abc456DEF", "只可輸入數字")]
        [InlineData("1234567890123456789", "繳款帳號不存在")]
        [InlineData("987654321", "企業代碼不存在")]

       
        public void 繳款帳號欄位(string input, string expect_result)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(test_url);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.

            IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
            PaymentAccount.Clear(); // 清除繳款欄位
            PaymentAccount.SendKeys(input); // 輸入值

            string actualPaymentAccount = PaymentAccount.GetAttribute("value");

            IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));


            bool digital_check_result = Regex.IsMatch(actualPaymentAccount, @"^[0-9]*$"); // 判斷輸入字串是否為"全數字"
            bool expect_check_result = Regex.IsMatch(actualPaymentAccount, @"^\d{6,16}$"); // 判斷輸入字串是否為"6~16位全數字"

          
            if (expect_check_result == true) // 判斷當"輸入欄位為6~16位全數字" >>> 送出data後延遲等待5秒, 等網頁load完
            {
                submit_button.Click(); // 點 送出
                System.Threading.Thread.Sleep(6000);
            }
            else
            {
                submit_button.Click(); // 點 送出
                System.Threading.Thread.Sleep(100);
            }

            if (expect_check_result != true)
            {
                string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text;
                Assert.Equal(expect_result, payment_account_error);
                System.Diagnostics.Debug.WriteLine($"輸入: {input}, 實際: {actualPaymentAccount}, 顯示訊息: {payment_account_error}");
            }

            else if (expect_check_result == true)
            {
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text;
                Assert.Equal(expect_result, searchresult);
                System.Diagnostics.Debug.WriteLine($"輸入: {input}, 實際: {actualPaymentAccount}, 顯示訊息: {searchresult}");
            }
            else 
            {
                expect_result = "FAIL";
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                System.Diagnostics.Debug.WriteLine($"輸入: {input}, 實際: {actualPaymentAccount}, 顯示訊息: {searchresult}");
            }
            driver.Quit();
        }
     }
}
