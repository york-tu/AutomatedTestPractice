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
    public class �m��_�۰ʿ�J�Ϥ����ҽX_�ɤse���H�U: Tesseract_OCR
    {
        private readonly string test_url = "https://www.esunbank.com.tw/s/PersonalLoanApply/Landing/IDConfirmDiversion?DG_landing";

        [Theory]
        [InlineData(BrowserType.Chrome)]



        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                //driver.Manage().Window.Maximize();

                
                IWebElement IdentityCardColumn = driver.FindElement(By.XPath("//*[@id='ApplicantID']")); // �����Ҧr����J���
                IWebElement BirthdayColumn = driver.FindElement(By.XPath("//*[@id='ApplicantBirthDate']")); // �X�ͦ~������
                IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath("//*[@id='Captcha']")); // ���ҽX��J���

                ImageVerificationCodeColumn.Click();
                BirthdayColumn.Click();

                ///<summary>
                /// �����Ҧr�����
                ///</summary>
                IdentityCardColumn.Clear();
                IdentityCardColumn.SendKeys("a123456789");


                ///<summary>
                /// �X�ͦ~������
                ///</summary>
                BirthdayColumn.Clear();
                BirthdayColumn.SendKeys("19550101");

                ///<summary>
                /// ���ҽX���
                ///</summary>
                
                IWebElement CaptchaPicture = driver.FindElement(By.XPath("//*[@id='esunOTPCaptchaHint']")); // ���ҽX�Ϥ�
                IWebElement NextButton = driver.FindElement(By.XPath("//*[@id='btnSubmitFormEsunOTP']"));
                IWebElement RetryButton = driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[1]/div"));

                string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\Practice";
                Tools.CreateSnapshotFolder(snapshotpath); 
                System.Threading.Thread.Sleep(100);

            retryagain:
                string time = System.DateTime.Now.ToString("MMddhhmmssffff");
                string fullnamepath = $@"{snapshotpath}\ImageVerifyCode_{time}.png";

                Tools.ElementTakeScreenShot(CaptchaPicture, fullnamepath); // snapshot"�Ϥ����ҽX"

                string verify_code_result = TesseractOCRIdentify(fullnamepath); //�ѪR�X���ҽX (Tesseract)
                // string verify_code_result = Tools.IronOCR($@"{snapshotpath}\ImageVerifyCode.png"); //�ѪR�X���ҽX (Iron)
                //string verify_code_result = Tools.BaiduOCR($@"{snapshotpath}\ImageVerifyCode.png"); //�ѪR�X���ҽX (�ʫ�)

                ImageVerificationCodeColumn.SendKeys(verify_code_result); // ��J���ҽX
                string image_verification_error = driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[2]")).Text; // ���ҽX���~�T�����

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
                System.Threading.Thread.Sleep(100000);
            }
             driver.Quit();
        }
    }
}


