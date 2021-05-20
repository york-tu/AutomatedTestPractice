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
        [InlineData("1234567890123456789", "企業代碼不存在")]
        [InlineData("987654321", "企業代碼不存在")]

       
        public void 繳款帳號欄位(string input, string expect_result)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(test_url);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.

            bool result = Regex.IsMatch(input, @"^[0-9]*$"); // 判斷輸入字串是否為"全數字"

            IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
            PaymentAccount.Clear(); // 清除繳款欄位
            PaymentAccount.SendKeys(input); // 輸入值

            string actualPaymentAccount = PaymentAccount.GetAttribute("value");

            IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));

            if (input.Length >= 6 && result == true) // 判斷當輸入字元為全數字且大於六位數時, 送出data後延遲等待5秒, 等網頁load完
            {
                submit_button.Click(); // 點 送出
                System.Threading.Thread.Sleep(6000);
            }
            else
            {
                submit_button.Click(); // 點 送出
                System.Threading.Thread.Sleep(100);
            }

            if (input == "")
            {
                string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text;
                Assert.Equal(expect_result, payment_account_error);
                System.Diagnostics.Debug.WriteLine("PASS_Case_1: " + input + payment_account_error);
            }
            else if (result != true)
            {
                string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text;
                Assert.Equal(expect_result, payment_account_error);
                System.Diagnostics.Debug.WriteLine("PASS_Case_3: " + input + "_" + payment_account_error);
            }
            else if (result == true && input.Length < 6) 
            {
                string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text;
                Assert.Equal(expect_result, payment_account_error);
                System.Diagnostics.Debug.WriteLine("PASS_Case_2: " + input + payment_account_error);
            }
            
            else if (result == true && input.Length >= 6 && input.Length <= 16)
            {
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text;
                Assert.Equal(expect_result, searchresult);
                System.Diagnostics.Debug.WriteLine("Case_4: " + input + "_" + searchresult);
            }
            else if (result == true && input.Length > 16) 
            {
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                Assert.Equal(expect_result, searchresult);
                System.Diagnostics.Debug.WriteLine("Case_5: " + input + "_" + searchresult);
            }
            else 
            {
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // 擷取"查詢結果" 欄位字串
                System.Diagnostics.Debug.WriteLine("Case_6: " + input + "_" +searchresult);
            }
            driver.Quit();
        }
     }
}
