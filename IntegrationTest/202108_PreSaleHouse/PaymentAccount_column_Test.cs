using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.Text.RegularExpressions;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.PreSaleHouseTrustInquiryTest
{
    public class �w��ΫH�U�d��: IntegrationTestBase
    {
        public �w��ΫH�U�d��(ITestOutputHelper output, Setup testSetup): base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType =typeof (BrowserHelper))]
        public void �w��άd�ߺ�i_����ˮ�(string browser)
        {
            StartTestCase(browser, "�d�ߺ�i_��J����ˮ�");

            INFO("�w��ΫH�U�d�ߥ\���i, �u�ƬJ���x���w��ΫH�U�d�ߵ��G, �w��ú�ڱb������ˮִ���");
            //[InlineData("", "������g")]
            //[InlineData("12345", "�̤� 6 �Ӧr")]
            //[InlineData("����r", "�u�i��J�Ʀr")]
            //[InlineData("�ϢТѢҢӢ�", "�u�i��J�Ʀr")]
            //[InlineData("�������", "�u�i��J�Ʀr")]
            //[InlineData("abcdef", "�u�i��J�Ʀr")]
            //[InlineData("ABCDEF", "�u�i��J�Ʀr")]
            //[InlineData("@#$%^&*", "�u�i��J�Ʀr")]
            //[InlineData("�I�I���C�H�s", "�u�i��J�Ʀr")]
            //[InlineData("123abc456DEF", "�u�i��J�Ʀr")]
            //[InlineData("1234567890123456789", "ú�ڱb�����s�b")]
            //[InlineData("987654321", "���~�N�X���s�b")]



            //Array�s���w���ժ��r��
            string[] array_payment_account = new string[] { "", "12345", "������������", "����r", "�ϢТѢҢӢ�", "�������", "abcdef", "ABCDEF", "@#$%^&*", "�I�I���C�H�s", "123abc456DEF", "987654321", "1234567890123456789" };
            //string[] array_payment_account = new string[] { "", "12345" };

            foreach (string keyin in array_payment_account) // �]�j��H�v����J���
            {

                IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
                PaymentAccount.Clear(); // �M��ú�����
                PaymentAccount.SendKeys(keyin); // ��J��

               

                string actualPaymentAccount = PaymentAccount.GetAttribute("value"); // ��J������
                

                IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));


                bool digital_check_result = Regex.IsMatch(actualPaymentAccount, @"^[0-9]*$"); // �P�_��J�r��O�_��"���Ʀr"
                bool expect_check_result = Regex.IsMatch(actualPaymentAccount, @"^\d{6,16}$"); // �P�_��J�r��O�_��"6~16����Ʀr"


                ///<summary>
                /// �P�_��"��J��쬰6~16����Ʀr" >>> �e�Xdata�᩵�𵥫�5��, ������load��
                /// </summary>
                if (expect_check_result == true) 
                {
                    submit_button.Click(); // �I �e�X
                    System.Threading.Thread.Sleep(6000);
                }
                else
                {
                    submit_button.Click(); // �I �e�X
                    System.Threading.Thread.Sleep(100);
                }

                if (actualPaymentAccount == "") 
                {
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // ����error message
                    if (payment_account_error == "������g")
                    {
                        PASS($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { payment_account_error}");
                        result = true;
                    }
                    else
                    {
                        FAIL($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { payment_account_error}");
                        result = false;
                    }
                }
                else if (digital_check_result == true && expect_check_result == false)
                {
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // ����error message
                    if (payment_account_error == "�̤� 6 �Ӧr")
                    {
                        PASS($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { payment_account_error}");
                        result = true;
                    }
                    else
                    {
                        FAIL($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { payment_account_error}");
                        result = false;
                    }
                }
                else if (digital_check_result == false)
                {
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // ����error message
                    if (payment_account_error == "�u�i��J�Ʀr")
                    {
                        PASS($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { payment_account_error}");
                        result = true;
                    }
                    else
                    {
                        FAIL($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { payment_account_error}");
                        result = false;
                    }
                }
                else if (expect_check_result == true)
                {
                    string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �d�ߵ��G���~�T��
                    if (searchresult == "ú�ڱb�����s�b")
                    {
                        PASS($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { searchresult }");
                        result = true;
                    }
                    else
                    {
                        FAIL($"��J: { keyin}, ������: { actualPaymentAccount}, ��ܰT��: { searchresult }");
                        result = false;
                    }
                }
            
            }
            PASS(TestBase.ElementSnapShotToReport(driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table"))));
            CloseBrowser();
        }  
     }
}
