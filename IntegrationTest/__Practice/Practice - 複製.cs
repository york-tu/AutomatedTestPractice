using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.IO;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.__Practice
{
    public class 練習_輸出Rerpot:IntegrationTestBase
    {
        public 練習_輸出Rerpot(ITestOutputHelper output, Setup testsetup) : base(output, testsetup)
        {
            testurl = domain + "https://www.esunbank.com.tw/s/PersonalLoanApply/Landing/IDConfirmDiversion?DG_landing";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]


        public void 輸出Rerpot(string browser)
        {
            StartTestCase(browser, "練習_輸出Rerpot", "York");
            INFO("練習輸出Rerpot");

            IWebElement IdentityCardColumn = driver.FindElement(By.XPath("//*[@id='ApplicantID']")); // 身分證字號輸入欄位
            IWebElement BirthdayColumn = driver.FindElement(By.XPath("//*[@id='ApplicantBirthDate']")); // 出生年月日欄位
            IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath("//*[@id='Captcha']")); // 驗證碼輸入欄位

            ImageVerificationCodeColumn.Click();
            BirthdayColumn.Click();

            ///<summary>
            /// 身分證字號欄位
            ///</summary>
            IdentityCardColumn.Clear();
            IdentityCardColumn.SendKeys("a123456789");


            ///<summary>
            /// 出生年月日欄位
            ///</summary>
            BirthdayColumn.Clear();
            BirthdayColumn.SendKeys("19550101");

            ///<summary>
            /// 驗證碼欄位
            ///</summary>

            IWebElement CaptchaPicture = driver.FindElement(By.XPath("//*[@id='esunOTPCaptchaHint']")); // 驗證碼圖片
            IWebElement NextButton = driver.FindElement(By.XPath("//*[@id='btnSubmitFormEsunOTP']")); // 下一步按鈕
            IWebElement RetryButton = driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[1]/div")); // Refresh 驗證碼圖片按鈕


            TestBase.CreateFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha");
            System.Threading.Thread.Sleep(100);
            TestBase.CleanUPFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha"); //清空Captcha資料夾



            int verify_count = 1; //紀錄驗證碼retry次數
        retryagain:
            System.Threading.Thread.Sleep(1000);
            string current_URL = driver.Url; //抓當下網頁URL

            if (current_URL != testurl) // 當下網頁不是初始網頁 >>> 代表驗證碼過, 已經PASS到下一頁
            {
                System.Threading.Thread.Sleep(3000);
                goto gotoNextStep; // 直接跳到程式結尾
            }


            TestBase.ElementSnapshotshot(CaptchaPicture, $@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); // snapshot"圖片驗證碼"

            if (verify_count >= 10) //依序刪除舊的captcha截圖
            {
                File.Delete($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count - 9}.png");
            }

            string verify_code_result = TestBase.TesseractOCRIdentify($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png", 1); //解析出驗證碼 (方法: Tesseract)
            // string verify_code_result = Tools.IronOCR($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //解析出驗證碼 (Iron)
            // string verify_code_result = Tools.BaiduOCR($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //解析出驗證碼 (百度)

            ImageVerificationCodeColumn.Clear();
            ImageVerificationCodeColumn.SendKeys(verify_code_result); // 輸入驗證碼
            System.Threading.Thread.Sleep(100);

            string NextButtonStatus = NextButton.GetAttribute("disabled"); // 獲取"下一步"button的狀態 "disabled" or not
            if (NextButtonStatus == "true")
            {
                RetryButton.Click();
                verify_count++;
                goto retryagain;
            }
            else
            {
                NextButton.Click();
                verify_count++;
                goto retryagain;
            }

        gotoNextStep:
            System.Threading.Thread.Sleep(2000);

            driver.Quit();
        }
    }
}


