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
using PictureRecognize;


namespace Practice
{
    public class 練習_自動輸入圖片驗證碼_玉山e指信貸:Tesseract_OCR
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
                                                                                      //driver.Manage().Window.Maximize();


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

                string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\Practice";
                Tools.CreateSnapshotFolder(snapshotpath);
                System.Threading.Thread.Sleep(100);

            retryagain:
                System.Threading.Thread.Sleep(1000);
                string time = System.DateTime.Now.ToString("MMddhhmmssffff");
                string fullnamepath = $@"{snapshotpath}\ImageVerifyCode_{time}.png";

                string current_URL = driver.Url; //抓當下網頁URL

                if (current_URL != test_url) // 當下網頁不是初始網頁 >>> 代表驗證碼過, 已經PASS到下一頁
                {
                    System.Threading.Thread.Sleep(5000);
                    goto gotoNextStep; // 直接跳到程式結尾
                }

                Tools.ElementTakeScreenShot(CaptchaPicture, fullnamepath); // snapshot"圖片驗證碼"

                string verify_code_result = TesseractOCRIdentify(fullnamepath); //解析出驗證碼 (方法: Tesseract)
                // string verify_code_result = Tools.IronOCR($@"{snapshotpath}\ImageVerifyCode.png"); //解析出驗證碼 (Iron)
                // string verify_code_result = Tools.BaiduOCR($@"{snapshotpath}\ImageVerifyCode.png"); //解析出驗證碼 (百度)
                
                ImageVerificationCodeColumn.Clear();
                ImageVerificationCodeColumn.SendKeys(verify_code_result); // 輸入驗證碼
                System.Threading.Thread.Sleep(100);

                string NextButtonStatus = NextButton.GetAttribute("disabled"); // 獲取"下一步"button的狀態 "disabled" or not
                if (NextButtonStatus == "true")
                {
                    RetryButton.Click();
                    goto retryagain;
                }
                else
                {
                    NextButton.Click();
                    goto retryagain;
                }

            gotoNextStep:
                System.Threading.Thread.Sleep(5000);
            }

            driver.Quit();
        }
    }
}


