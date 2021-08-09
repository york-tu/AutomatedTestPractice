using OpenQA.Selenium;
using Xunit;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;
using AutomatedTest.Utilities;
using Xunit.Abstractions;
using System;

namespace AutomatedTest.IntegrationTest.Personal.Trust
{
    public class �w��ΫH�U�d�ߺ�i_PC��: IntegrationTestBase
    {
        public �w��ΫH�U�d�ߺ�i_PC��(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/trust/announcement-and-inquiry/pre-construction-real-estate-trust-inquiry";
        }

        private readonly string testcase_name = "�w��άd��";

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]
        public void ����ˮִ���(string browser)
        {
            StartTestCase(browser, "�w��ΫH�U�d�ߺ�i_����ˮִ���", "York");
            INFO("");
            string time = System.DateTime.Now.ToString("yyyy-MM-dd_HHmm");
            string snapshotfolderpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\PreSaleHouseTrustInquiry";
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

                IWebElement ProjectNameDropDownList = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul"));
                ProjectNameDropDownList.Click(); // �I�خפU�Կ��
                System.Threading.Thread.Sleep(100);
                IWebElement SelectProject = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_PnlInput']/div[2]/table/tbody/tr[1]/td[2]/ul/li/ul/li[6]/span"));
                SelectProject.Click();// ��خ�
                System.Threading.Thread.Sleep(100);
                
                IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='layout_0_rightcontent_1_LbtnQuery']"));
                SubmitButton.Click();

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

                    

                    string actualPaymentAccountColumn = PaymentAccount.GetAttribute("value"); // ������Ū���쪺��

                    bool regexCheckResult = Regex.IsMatch(actualPaymentAccountColumn, @"^\d{6,16}$"); // �P�_���̪��r��O�_��"6-16����Ʀr"
                    bool numericResult = Regex.IsMatch(actualPaymentAccountColumn, @"^[0-9]*$"); // �P�_���̪��r��O�_��"���Ʀr"
                
                    string payment_account_error = driver.FindElement(By.Id("layout_0_rightcontent_1_TxtVacno-error")).Text; // �����~�T��

                    if (regexCheckResult == true) // �P�_���J�r���ŦX�w��(���Ʀr, 6~16���), �e�Xdata�᩵�𵥫�5��, ������load��
                    {
                    SubmitButton.Click(); // �I �e�X
                        System.Threading.Thread.Sleep(5000);
                    }
                    else
                    {
                    SubmitButton.Click(); // �I �e�X
                        System.Threading.Thread.Sleep(100);
                    }


                    if (string.IsNullOrWhiteSpace(keyin) == true)// �P�_���J��쬰"�ŭ�" >>> �w����� "������g" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '������g'";
                        xlsSheet[test_result].Value = "Case_1";
                        try
                        {
                           Assert.Equal("������g", payment_account_error);
                            PASS($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: ������g");
                        }
                         catch (System.Exception e)
                        {
                           ERROR(e.Message);
                           FAIL($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: ������g");
                           Assert.True(false);
                        }
                    }


                    else if (string.IsNullOrWhiteSpace(actualPaymentAccountColumn) != true && numericResult != true) //  �P�_���J�D�Ʀr  >>> �w����� "�u�i��J�Ʀr" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '�u�i��J�Ʀr'";
                        xlsSheet[test_result].Value = "Case_3";
                       try
                       {
                            Assert.Equal("�u�i��J�Ʀr", payment_account_error);
                            PASS($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: �u�i��J�Ʀr");
                       }
                       catch (System.Exception e)
                       {
                           ERROR(e.Message);
                           FAIL($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: �u�i��J�Ʀr");
                           Assert.True(false);
                       }
                    }
                    else if (string.IsNullOrWhiteSpace(actualPaymentAccountColumn) != true && regexCheckResult != true) // �P�_��"��J���Ʀr�p�󤻦��" >>> �w����� "�̤� 6 �Ӧr" warning
                    {
                        xlsSheet[check_position].Value = keyin;
                        xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                        xlsSheet[show_msg].Value = payment_account_error;
                        xlsSheet[expect_result].Value = "��� '�̤� 6 �Ӧr'";
                        xlsSheet[test_result].Value = "Case_2";

                         try
                         {
                            Assert.Equal("�̤�6�Ӧr", payment_account_error);
                            PASS($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: �̤�6�Ӧr");
                         }
                        catch (System.Exception e)
                         {
                           ERROR(e.Message);
                           FAIL($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: �̤�6�Ӧr");
                           Assert.True(false);
                         }
                    }
                    else if (string.IsNullOrWhiteSpace(actualPaymentAccountColumn) != true && regexCheckResult == true) // ��J "�������Ʀr" �B "�r�Ƥ��� 6~16 ��ƶ�
                    {
                    retry:
                        SubmitButton.Click();
                        System.Threading.Thread.Sleep(15000);
                        IWebElement verifycode = driver.FindElement(By.XPath("")); // ���ҽX���
                        verifycode.Clear();
                        verifycode.Click();
                        System.Threading.Thread.Sleep(15000);

                        SubmitButton.Click();

                        System.Threading.Thread.Sleep(15000);
                        string verifycode_error = driver.FindElement(By.Id("captchaWrong")).Text; // ���ҽX���~�T��
                        string actualVerifyCodeColumnValue = verifycode.GetAttribute("value"); // ���ҽX�����Ū�쪺��
                          if (string.IsNullOrWhiteSpace(actualVerifyCodeColumnValue) == true) // ��J�����ҽX
                          {
                             try
                             {
                                Assert.Equal("�п�J���ҽX", payment_account_error);
                                PASS($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: �п�J���ҽX");
                             }
                            catch (System.Exception e)
                             {
                               ERROR(e.Message);
                               FAIL($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: �п�J���ҽX");
                               Assert.True(false);
                             }
                          goto retry;
                         }
                          else if (string.IsNullOrWhiteSpace(actualVerifyCodeColumnValue) != true && string.IsNullOrWhiteSpace(verifycode_error) != true)
                          {
                            try
                             {
                                Assert.Equal("���ҽX���~", payment_account_error);
                                PASS($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: ���ҽX���~");
                             }
                            catch (System.Exception e)
                             {
                               ERROR(e.Message);
                               FAIL($"��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: ���ҽX���~");
                               Assert.True(false);
                             }
                            goto retry;
                          }
                          else if (string.IsNullOrWhiteSpace(actualVerifyCodeColumnValue) != true && string.IsNullOrWhiteSpace(verifycode_error) == true)
                           {
                                
                                INFO("");
                                INFO("�ˬd�d�ߵ��G");

                                string searchresult = driver.FindElement(By.Id("layout_0_rightcontent_1_LblMessage")).Text; // �^��"�d�ߵ��G" ���r��
                                
                                if (string.IsNullOrWhiteSpace(searchresult) != true)
                                {
                                xlsSheet[check_position].Value = keyin;
                                xlsSheet[actual_value].Value = actualPaymentAccountColumn;
                                xlsSheet[show_msg].Value = searchresult;
                                xlsSheet[expect_result].Value = "��� '���~�N�X���s�b / ú�ڱb�����s�b'";
                                xlsSheet[test_result].Value = "Case_4";

                                WARNING($"[NeedCheck][���T���ҽX], ��J�b��: {keyin}, ��J����{keyin.Length}, �����Ū��: {actualPaymentAccountColumn}, Ū�����: {actualPaymentAccountColumn.Length}, ��ڿ��~�T��: {payment_account_error}, �w�����~�T��: ���~�N�X���s�b / ú�ڱb�����s�b");
                                }

                                bool date_column = driver.FindElement(By.XPath("")).Displayed; // Result ��� "���" ������
                                bool summary_column = driver.FindElement(By.XPath("")).Displayed; // Result ��� "�K�n" ������
                                bool expenditure_column = driver.FindElement(By.XPath("")).Displayed; // Result ��� "��X" ������
                                bool deposit_column = driver.FindElement(By.XPath("")).Displayed; // Result ��� "�s�J" ������
                                bool surolus_column = driver.FindElement(By.XPath("")).Displayed; // Result ��� "���l" ������
                                
                                try
                                {
                                    Assert.True(date_column);
                                    Assert.True(summary_column );
                                    Assert.True(expenditure_column);
                                    Assert.True(deposit_column);
                                    Assert.True(surolus_column);
                                    PASS("�d�ߵ��G�ŦX�w��");
                                }
                                catch (Exception e)
                                {   
                                    ERROR(e.Message);
                                    PASS("�d�ߵ��G���ŦX�w��");
                                    Assert.True(false);
                                }
                          }
                    }
                }

                if (File.Exists(excel_path) != true)
                {
                    xlsWorkbook.SaveAs(excel_path);
                }
                else
                {
                    xlsWorkbook.Save();
                }
            CloseBrowser();
        }
    }
}
