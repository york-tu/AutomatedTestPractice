using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using Xunit;
using AutomatedTest.Utilities;
using System.IO;
using AventStack.ExtentReports;

namespace Practice
{
    public class �m��_��XRerpot
    {
        private readonly string test_url = "https://www.esunbank.com.tw/s/PersonalLoanApply/Landing/IDConfirmDiversion?DG_landing";

        [Theory]
        [InlineData(BrowserType.Chrome)]

        public void TestCase(BrowserType browserType)
        {
            //Tools.KillProcess("chromedriver.exe");
            //Tools.KillProcess("conhostexe");

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
                IWebElement NextButton = driver.FindElement(By.XPath("//*[@id='btnSubmitFormEsunOTP']")); // �U�@�B���s
                IWebElement RetryButton = driver.FindElement(By.XPath("//*[@id='browserIE']/div/ul/li[3]/div/div[1]/div")); // Refresh ���ҽX�Ϥ����s

 
                Tools.CreateSnapshotFolder($@"{UserDataList.folderpath}\Captcha");
                System.Threading.Thread.Sleep(100);
                Tools.CleanUPFolder($@"{UserDataList.folderpath}\Captcha"); //�M��Captcha��Ƨ�



                int verify_count = 1; //�������ҽXretry����
            retryagain:
                System.Threading.Thread.Sleep(1000);
                string current_URL = driver.Url; //���U����URL

                if (current_URL != test_url) // ��U�������O��l���� >>> �N�����ҽX�L, �w�gPASS��U�@��
                {
                    System.Threading.Thread.Sleep(3000);
                    goto gotoNextStep; // ��������{������
                }

                
                Tools.ElementTakeScreenShot(CaptchaPicture, $@"{UserDataList.folderpath}\Captcha\CaptchaImage_{verify_count}.png"); // snapshot"�Ϥ����ҽX"
               
                if (verify_count >=10) //�̧ǧR���ª�captcha�I��
                {
                    File.Delete($@"{UserDataList.folderpath}\Captcha\CaptchaImage_{verify_count-9}.png");
                }

                //string verify_code_result = TesseractOCRIdentify($@"{UserDataList.folderpath}\Captcha\CaptchaImage_{verify_count}.png", 1); //�ѪR�X���ҽX (��k: Tesseract)
                // string verify_code_result = Tools.IronOCR($@"{ snapshotpath}\ImageVerifyCode.png"); //�ѪR�X���ҽX (Iron)
                // string verify_code_result = Tools.BaiduOCR($@"{snapshotpath}\ImageVerifyCode.png"); //�ѪR�X���ҽX (�ʫ�)
                
                ImageVerificationCodeColumn.Clear();
               // ImageVerificationCodeColumn.SendKeys(verify_code_result); // ��J���ҽX
                System.Threading.Thread.Sleep(100);

                string NextButtonStatus = NextButton.GetAttribute("disabled"); // ���"�U�@�B"button�����A "disabled" or not
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
            }

            driver.Quit();
        }
    }
}


