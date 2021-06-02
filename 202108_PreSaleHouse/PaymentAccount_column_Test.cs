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
    public class PaymentAccount// �˴ڱb������ˬd
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";

        [Theory]
        [InlineData("", "������g")]
        [InlineData("12345", "�̤� 6 �Ӧr")]
        [InlineData("����r", "�u�i��J�Ʀr")]
        [InlineData("�ϢТѢҢӢ�", "�u�i��J�Ʀr")]
        [InlineData("�������", "�u�i��J�Ʀr")]
        [InlineData("abcdef", "�u�i��J�Ʀr")]
        [InlineData("ABCDEF", "�u�i��J�Ʀr")]
        [InlineData("@#$%^&*", "�u�i��J�Ʀr")]
        [InlineData("�I�I���C�H�s", "�u�i��J�Ʀr")]
        [InlineData("123abc456DEF", "�u�i��J�Ʀr")]
        [InlineData("1234567890123456789", "ú�ڱb�����s�b")]
        [InlineData("987654321", "���~�N�X���s�b")]

       
        public void ú�ڱb�����(string input, string expect_result)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(test_url);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.

            IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
            PaymentAccount.Clear(); // �M��ú�����
            PaymentAccount.SendKeys(input); // ��J��

            string actualPaymentAccount = PaymentAccount.GetAttribute("value");

            IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));


            bool digital_check_result = Regex.IsMatch(actualPaymentAccount, @"^[0-9]*$"); // �P�_��J�r��O�_��"���Ʀr"
            bool expect_check_result = Regex.IsMatch(actualPaymentAccount, @"^\d{6,16}$"); // �P�_��J�r��O�_��"6~16����Ʀr"

          
            if (expect_check_result == true) // �P�_��"��J��쬰6~16����Ʀr" >>> �e�Xdata�᩵�𵥫�5��, ������load��
            {
                submit_button.Click(); // �I �e�X
                System.Threading.Thread.Sleep(6000);
            }
            else
            {
                submit_button.Click(); // �I �e�X
                System.Threading.Thread.Sleep(100);
            }

            if (expect_check_result != true)
            {
                string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text;
                Assert.Equal(expect_result, payment_account_error);
                System.Diagnostics.Debug.WriteLine($"��J: {input}, ���: {actualPaymentAccount}, ��ܰT��: {payment_account_error}");
            }

            else if (expect_check_result == true)
            {
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text;
                Assert.Equal(expect_result, searchresult);
                System.Diagnostics.Debug.WriteLine($"��J: {input}, ���: {actualPaymentAccount}, ��ܰT��: {searchresult}");
            }
            else 
            {
                expect_result = "FAIL";
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                System.Diagnostics.Debug.WriteLine($"��J: {input}, ���: {actualPaymentAccount}, ��ܰT��: {searchresult}");
            }
            driver.Quit();
        }
     }
}
