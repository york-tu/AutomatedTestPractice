using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Xunit;
using System;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;

namespace �w��ΫH�U�d��
{
    public class �w��ΫH�U�d��
    {
        readonly string test_url = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        readonly string excel_path = @"D:\PreSaleHouseTrustInquiry_TestReport.xls";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void PreSaleHouseTrustInquiry(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);

            using var driver = WebDriverInfra.Create_Browser(browserType);
            {
                string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");

                WorkBook xlsWorkbook;
                WorkSheet xlsSheet;
                if (File.Exists(excel_path) != true)
                {
                    xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLS); //�w�q excel�榡�� XLS
                    xlsSheet = xlsWorkbook.CreateWorkSheet("�w��άd��" + timesavepath); // Create excel�� "sheet"����
                }
                else
                {
                    xlsWorkbook = WorkBook.Load(excel_path); //Ū�� excel ��
                    xlsSheet = xlsWorkbook.CreateWorkSheet("�w��άd��" + timesavepath); // Create excel�� "sheet"����
                }

                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                //driver.Manage().Window.Maximize();

                IWebElement ProjectNameDropDownList = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul"));
                ProjectNameDropDownList.Click(); // �I�خפU�Կ��
                System.Threading.Thread.Sleep(100);
                IWebElement SelectProject = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul/li/ul/li[6]/span"));
                SelectProject.Click();// ��خ�
                System.Threading.Thread.Sleep(100);


                //Array�s���w���ժ��r��
                string[] array_payment_account = new string[] {"", "1234567890123456789","12345","������������", "����r", "�ϢТѢҢӢ�", "�������", "abcdef", "ABCDEF", "@#$%^&*", "�I�I���C�H�s", "123abc456DEF", "987654321"};
                
                //�w�qexcel���
                xlsSheet["A1"].Value = "�ˬd ú�ڱb�����";
                xlsSheet["A2"].Value = "�خצW��";
                xlsSheet["B2"].Value = "��ʿ�J";
                xlsSheet["C2"].Value = "����ڭ�";
                xlsSheet["D2"].Value = "��ܰT��";
                xlsSheet["E2"].Value = "�w�����G";
                xlsSheet["F2"].Value = "���յ��G";
                xlsSheet["A2:F2"].Style.BottomBorder.SetColor("#ff6600"); // ���u"����"
                xlsSheet["A2:F2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double; //�[�����u
                xlsSheet["A3"].Value = ProjectNameDropDownList.Text; // ���=�خצW��


                int i = 3;
                foreach (var keyin in array_payment_account) // �]�j��H�v����J���
                {
                    string check_position = "B" + i;
                    string actual_value = "C" + i;
                    string show_msg = "D" + i;
                    string expect_result = "E" + i;
                    string test_result = "F" + i;
                    

                    bool result = Regex.IsMatch(keyin, @"^[0-9]*$"); // �P�_��J�r��O�_��"���Ʀr"

                    IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
                    PaymentAccount.Clear(); // �M��ú�����
                    PaymentAccount.SendKeys(keyin); // ��J��

                    string actualPaymentAccount = PaymentAccount.GetAttribute("value");

                    IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));
                   
                    if (keyin.Length >= 6 && result == true) // �P�_���J�r�������Ʀr�B�j�󤻦�Ʈ�, �e�Xdata�᩵�𵥫�5��, ������load��
                    {
                        submit_button.Click(); // �I �e�X
                        System.Threading.Thread.Sleep(5000);
                    }
                    else
                    {
                        submit_button.Click(); // �I �e�X
                        System.Threading.Thread.Sleep(100);
                    }

                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text;

                    if (keyin == "" && payment_account_error == "������g")// �P�_���J��"�ŭ�" >>> �w����� "������g" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '������g'";
                        xlsSheet[test_result].Value = "PASS_Case_1";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_1: " + keyin);
                    }
                    else if (keyin.Length < 6 && payment_account_error == "�̤� 6 �Ӧr") // �P�_���J�r��"�p�󤻦��" >>> �w����� "�̤� 6 �Ӧr" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '�̤� 6 �Ӧr'";
                        xlsSheet[test_result].Value = "PASS_Case_2";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_2: " + keyin);
                    }
                    else if (result != true && payment_account_error == "�u�i��J�Ʀr") // �P�_���J"���O���Ʀr" >>> �w����� "�u�i��J�Ʀr" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '�u�i��J�Ʀr'";
                        xlsSheet[test_result].Value = "PASS_Case_3";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_3: " + keyin);
                    }

                    // �P�_���J "�������Ʀr" �B "�r�Ƥ��� 6~16 ��ƶ� >>> �w���d�ߵ��G��� "���~�N�X���s�b"
                    else if (result == true && keyin.Length >=6 && keyin.Length <=16 && driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text == "���~�N�X���s�b")
                    {   
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text;
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = searchresult;
                        xlsSheet[expect_result].Value = "��� '���~�N�X���s�b'";
                        xlsSheet[test_result].Value = "PASS_Case_4";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_4: " + keyin);
                    }
                    else if (keyin.Length > 16 && result == true) // �P�_���J�줸 "�j��16���" �B "�������Ʀr" >> �w���d�ߵ��G��� "���~�N�X���s�b (need to confirm)"
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = searchresult;
                        xlsSheet[expect_result].Value = "��� '[TBD]���~�N�X���s�b'";
                        xlsSheet[test_result].Value = "Need Manaul Check";
                        System.Diagnostics.Debug.WriteLine("PASS_Case_5: " + keyin);
                    }
                    else //�D�H�W���p���յ��Gfail
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = "Message 1: " + payment_account_error + "or Message 2: " + searchresult;
                        xlsSheet[test_result].Value = "FAIL";
                    }
                    i++;
                }

                if (File.Exists(excel_path) != true)
                {
                    xlsWorkbook.SaveAs(excel_path);
                }
                else
                {
                    xlsWorkbook.Save();
                }
            }
            driver.Quit();
        }
    }
}
