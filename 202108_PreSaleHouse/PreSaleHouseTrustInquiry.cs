using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using Xunit;
using System;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;
using Utilities;

namespace PreSaleHouseTrustInquiryTest
{
    public class �w��ΫH�U�d�ߺ�i_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        
        private readonly string testcase_name = "�w��άd��";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);

            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                string tim = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm");
                string snapshotfolderpath = $@"{UserDataList.folderpath}\SnapshotFolder\PreSaleHouseTrustInquiry";
                string excel_path = $@"{snapshotfolderpath}\TestReport.xlsx";

                Tools.CreateSnapshotFolder(snapshotfolderpath);
                System.Threading.Thread.Sleep(100);


                WorkBook xlsWorkbook;
                WorkSheet xlsSheet;

                if (File.Exists(excel_path) == true && WorkBook.Load(excel_path).GetWorkSheet(testcase_name) == null)
                { // �P�_��excel�ɦs�b �� �S��"�w��άd��"�u�@�� >>> Ū����excel��, new create "�w��άd��" �u�@�� 
                    xlsWorkbook = WorkBook.Load(excel_path); 
                    xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name); 
                }
                else if (File.Exists(excel_path) == true && WorkBook.Load(excel_path).GetWorkSheet(testcase_name) != null)
                { // �P�_��excel�ɦs�b �B ��"�w��άd��"�u�@�� >>> Ū����excel��, Ū��"�w��άd��" �u�@��
                    xlsWorkbook = WorkBook.Load(excel_path);
                    xlsSheet = xlsWorkbook.GetWorkSheet(testcase_name); 
                }
                else
                { // �P�_��excel�ɤ��s�b >>> new create excel�� & new create "�w��άd��" �u�@��
                    xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLSX); //�w�q excel�榡�� XLSX
                    xlsSheet = xlsWorkbook.CreateWorkSheet(testcase_name);
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
                string[] array_payment_account = new string[] {"", "1234567890123456789","12345","������������", "����r", "�ϢТѢҢӢ�", "�������", "abcdef", "ABCDEF", "@#$%^&*", "�I�I���C�H�s", "123abc456DEF", "987654321", ""};
                
                //�w�qexcel���
                xlsSheet["A1"].Value = "�ˬd ú�ڱb�����";
                xlsSheet["A2"].Value = "�خצW��";
                xlsSheet["B2"].Value = "��ʿ�J";
                xlsSheet["C2"].Value = "����ڭ�";
                xlsSheet["D2"].Value = "��ܰT��";
                xlsSheet["E2"].Value = "�w�����G(��ܰT��)";
                xlsSheet["F2"].Value = "���յ��G";
                xlsSheet["A2:F2"].Style.BottomBorder.SetColor("#ff6600"); // ���u"����"
                xlsSheet["A2:F2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double; //�[�����u
                xlsSheet["A3"].Value = ProjectNameDropDownList.Text; // ���=�خצW��


                int i = 3;
                foreach (string keyin in array_payment_account) // �]�j��H�v����J���
                {
                    string check_position = "B" + i;
                    string actual_value = "C" + i;
                    string show_msg = "D" + i;
                    string expect_result = "E" + i;
                    string test_result = "F" + i;

                    IWebElement PaymentAccount = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_TxtVacno']"));
                    PaymentAccount.Clear(); // �M��ú�����
                    PaymentAccount.SendKeys(keyin); // ��J��

                    IWebElement submit_button = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));

                    string actualPaymentAccount = PaymentAccount.GetAttribute("value"); // ������Ū���쪺��

                    bool expect_check_result = Regex.IsMatch(actualPaymentAccount, @"^\d{6,16}$"); // �P�_���̪��r��O�_��"6-16����Ʀr"
                    bool digital_check_result = Regex.IsMatch(actualPaymentAccount, @"^[0-9]*$"); // �P�_���̪��r��O�_��"���Ʀr"
                

                    if (expect_check_result == true) // �P�_���J�r���ŦX�w��(���Ʀr, 6~16���), �e�Xdata�᩵�𵥫�5��, ������load��
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

                    if (actualPaymentAccount == "")// �P�_���J��쬰"�ŭ�" >>> �w����� "������g" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '������g'";
                        xlsSheet[test_result].Value = "Case_1";
                    }
                    else if (actualPaymentAccount.Length < 6 && digital_check_result == true) // �P�_��"��J���Ʀr�p�󤻦��" >>> �w����� "�̤� 6 �Ӧr" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '�̤� 6 �Ӧr'";
                        xlsSheet[test_result].Value = "Case_2";
                    }
                    else if (digital_check_result != true) // �P�_���J���"�������Ʀr" >>> �w����� "�u�i��J�Ʀr" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '�u�i��J�Ʀr'";
                        xlsSheet[test_result].Value = "Case_3";
                    }

                    // �P�_���J "�������Ʀr" �B "�r�Ƥ��� 6~16 ��ƶ� >>> �w���d�ߵ��G��� "���~�N�X���s�b"
                    else if (expect_check_result == true)
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text;
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = searchresult;
                        xlsSheet[expect_result].Value = "��� '���~�N�X���s�b'";
                        xlsSheet[test_result].Value = "Case_4";
                    }
                    //else if (keyin.Length > 16 && result == true) // �P�_���J�줸 "�j��16���" �B "�������Ʀr" >> �w���d�ߵ��G��� "���~�N�X���s�b (need to confirm)"
                    //{
                    //    string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                    //    xlsSheet[check_position].Value = keyin;
                    //    xlsSheet[actual_value].Value = actualPaymentAccount;
                    //    xlsSheet[show_msg].Value = searchresult;
                    //    xlsSheet[expect_result].Value = "��� '[TBD]���~�N�X���s�b'";
                    //    xlsSheet[test_result].Value = "Need Manaul Check";
                    //}
                    else //�D�H�W���p���յ��Gfail
                    {
                        string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccount;
                        xlsSheet[show_msg].Value = $"Message 1: {payment_account_error} / Message 2: {searchresult}";
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
