using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.IO;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.__Practice
{
    public class �m��_��XRerpot:IntegrationTestBase
    {
        public �m��_��XRerpot(ITestOutputHelper output, Setup testsetup) : base(output, testsetup)
        {
            testurl = domain + "https://www.esunbank.com.tw/s/PersonalLoanApply/Landing/IDConfirmDiversion?DG_landing";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]


        public void ��XRerpot(string browser)
        {
            StartTestCase(browser, "�m��_��XRerpot", "York");
            INFO("�m�߿�XRerpot");

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


            Tools.CreateSnapshotFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha");
            System.Threading.Thread.Sleep(100);
            Tools.CleanUPFolder($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha"); //�M��Captcha��Ƨ�



            int verify_count = 1; //�������ҽXretry����
        retryagain:
            System.Threading.Thread.Sleep(1000);
            string current_URL = driver.Url; //���U����URL

            if (current_URL != testurl) // ��U�������O��l���� >>> �N�����ҽX�L, �w�gPASS��U�@��
            {
                System.Threading.Thread.Sleep(3000);
                goto gotoNextStep; // ��������{������
            }


            Tools.ElementSnapshotshot(CaptchaPicture, $@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); // snapshot"�Ϥ����ҽX"

            if (verify_count >= 10) //�̧ǧR���ª�captcha�I��
            {
                File.Delete($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count - 9}.png");
            }

            string verify_code_result = TesseractOCRIdentify($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png", 1); //�ѪR�X���ҽX (��k: Tesseract)
            // string verify_code_result = Tools.IronOCR($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //�ѪR�X���ҽX (Iron)
            // string verify_code_result = Tools.BaiduOCR($@"{System.AppDomain.CurrentDomain.BaseDirectory}\Captcha\CaptchaImage_{verify_count}.png"); //�ѪR�X���ҽX (�ʫ�)

            ImageVerificationCodeColumn.Clear();
            ImageVerificationCodeColumn.SendKeys(verify_code_result); // ��J���ҽX
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

            driver.Quit();
        }
    }
}


