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
        [InlineData("1234567890123456789", "���~�N�X���s�b")]
        [InlineData("987654321", "���~�N�X���s�b")]

       
        public void ú�ڱb�����(string input, string expect_result)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl(test_url);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.

            bool result = Regex.IsMatch(input, @"^[0-9]*$"); // �P�_��J�r��O�_��"���Ʀr"

            IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
            PaymentAccount.Clear(); // �M��ú�����
            PaymentAccount.SendKeys(input); // ��J��

            string actualPaymentAccount = PaymentAccount.GetAttribute("value");

            IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));

            if (input.Length >= 6 && result == true) // �P�_���J�r�������Ʀr�B�j�󤻦�Ʈ�, �e�Xdata�᩵�𵥫�5��, ������load��
            {
                submit_button.Click(); // �I �e�X
                System.Threading.Thread.Sleep(6000);
            }
            else
            {
                submit_button.Click(); // �I �e�X
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
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                Assert.Equal(expect_result, searchresult);
                System.Diagnostics.Debug.WriteLine("Case_5: " + input + "_" + searchresult);
            }
            else 
            {
                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                System.Diagnostics.Debug.WriteLine("Case_6: " + input + "_" +searchresult);
            }
            driver.Quit();
        }
     }
}
