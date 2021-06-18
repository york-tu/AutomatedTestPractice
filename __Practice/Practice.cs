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


namespace Practice
{
    public class 練習_自動輸入圖片驗證碼_玉山e指信貸
    {
        private readonly string test_url = "https://www.esunbank.com.tw/s/PersonalLoanApply/Landing/IDConfirmDiversion?DG_landing";

        [Theory]
        [InlineData(BrowserType.Chrome)]



        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                driver.Manage().Window.Maximize();

                string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\Practice";
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
                
                IWebElement CaptchaPicture = driver.FindElement(By.XPath("//*[@id='esunOTPCaptchaHint']")); // 圖片驗證碼圖片欄位
                IWebElement NextButton = driver.FindElement(By.XPath("//*[@id='btnSubmitFormEsunOTP']"));
                IWebElement RetryButton = driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[1]/div"));

                Tools.CreateSnapshotFolder(snapshotpath); 
                System.Threading.Thread.Sleep(100);

            retryagain:
                Tools.ElementTakeScreenShot(CaptchaPicture, $@"{snapshotpath}\ImageVerifyCode.png"); // 截圖"圖片驗證碼"圖片

                string verify_code_result = Tools.BaiduOCR($@"{snapshotpath}\ImageVerifyCode.png"); //解析出驗證碼
                ImageVerificationCodeColumn.SendKeys(verify_code_result); // 輸入驗證碼
                string image_verification_error = driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[2]")).Text; // 驗證碼錯誤訊息欄位

                if (driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[2]")).Text == "")
                {
                    NextButton.Click();

                    if (driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[2]")).Text == "")
                    {
                        goto gotoNextStep;
                    }
                    else
                    {
                        RetryButton.Click();
                        goto retryagain;
                    }
                }
                else
                {
                    RetryButton.Click();
                    goto retryagain;

                }
               gotoNextStep:
                System.Threading.Thread.Sleep(1000);
            }
             driver.Quit();
        }
    }
}


