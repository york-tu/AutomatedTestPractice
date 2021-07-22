using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.Text.RegularExpressions;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.BailoutLoan.LaborReliefLoanAppointmentService
{
    public class ���~�V�x�U�ڹw���A��_����ˮ�:IntegrationTestBase
    {
        public ���~�V�x�U�ڹw���A��_����ˮ�(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/small-business/tools/apply/sbloan-appointment";
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void ����ˮ�(string browser)
        {
            StartTestCase(browser, "���~�V�x�U�ڹw���A��_����ˮ�", "York");
            INFO("�T�{����ˮ֥��T");


            INFO("�ˬd�����Whyperlink�O�_���T");
            string business_bailout_zone_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[1]/span/a")).GetAttribute("href"); // ��� "���~�V�x�M��" ���W�s�����}
            string expect_business_bailout_zone_path = "https://www.esunbank.com.tw/bank/marketing/loan/rescueplan-all#sortorder400";
            if (business_bailout_zone_hyperlink == expect_business_bailout_zone_path)
            {
                PASS("���~�V�x�M�� hyperlink �ŦX�w��");
            }
            else
            {
                FAIL($"���~�V�x�M�� link: {business_bailout_zone_hyperlink}, �w�� link: {expect_business_bailout_zone_path}");
            }
            string existing_loan_relief_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[2]/a")).GetAttribute("href"); // ��� �J���U�ڼe�w �� "�W�s��" ���}
            string expect_existing_loan_relief_path = "https://www.esunbank.com.tw/bank/small-business/tools/apply/contact-me";
            if (existing_loan_relief_hyperlink == expect_existing_loan_relief_path)
            {
                PASS("�J���U�ڼe�w hyperlink �ŦX�w��");
            }
            else
            {
                FAIL($"�J���U�ڼe�w link: {existing_loan_relief_hyperlink}, �w�� link: {expect_existing_loan_relief_path}");
            }
            IWebElement ExistLoanReliefButton = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[1]")); 
            string existing_loan_relief_button = ExistLoanReliefButton.GetAttribute("href"); //��� �J���U�ڼe�w "���s" ���s�����}
            string expect_existing_loan_relief_button_path = "https://www.esunbank.com.tw/bank/small-business/tools/apply/contact-me";
            if(existing_loan_relief_button == expect_existing_loan_relief_button_path)
            {
                PASS("�J���U�ڼe�w'����'�s���ŦX�w��");
            }
            else
            {
                FAIL($"�J���U�ڼe�w'����'�s��: {existing_loan_relief_button}, �w�� link: {expect_existing_loan_relief_button_path}");
            }

            
            ///<summary>
            /// �ˬd���� "�W����"
            ///</summary>
            INFO("");
            INFO("�ˬd�����W����");
            string upper_wordings_first_line = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[1]")).Text;
            string upper_wordings_second_line = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[2]")).Text;
            string expect_upper_first_line_wordings = "���]���̴��������I�H�y���ޡA��ĳ�i�ܥ��~�V�x�M�ϸԾ\�ǳƸ�Ƹ�T�A�A�i��u�W�w���A�ȡA�H�[�t�ӿ�y�{�C";
            string expect_upper_second_line_wordings = "�������������~�V�x�s�U�ץ�M�ΡA�ӽСu�J���U�ڼe�w�v�A���I��J���U�ڼe�w�e���w�������C";
            if (upper_wordings_first_line == expect_upper_first_line_wordings)
            {
                PASS("�W���װϤ��1: �ŦX�w��");
            }
            else
            {
                FAIL($"�W���װϤ��1�ԭz: {upper_wordings_first_line}, �w�� link: {expect_upper_first_line_wordings}");
            }
            if (upper_wordings_second_line == expect_upper_second_line_wordings)
            {
                PASS("�W���װϤ��2: �ŦX�w��");
            }
            else
            {
                FAIL($"�W���װϤ��2�ԭz: {upper_wordings_second_line}, �w�� link: {expect_upper_second_line_wordings}");
            }
                    
            IWebElement TelephoneExtensionColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[3]"));
            TelephoneExtensionColumn.SendKeys("QWE!@#$%"); // ��������J���~��r
            IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));
            SubmitButton.Click(); // �I �e�X


     
            ///<summary>
            ///���q�W������ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�: ���q�W��");
            string[] CompanyNameArray = new string[] {"", "       ", "�Ϣ�", "ab", "��", "��", "�^����뤤��Ң��Ÿ�#$%!@~&*^%", "����r�[�^��r�[�Ʀr_�ϢТѢҢӢԢբ֢עء�abcdefghij_1234567890_ABCDEFGHIJ_", "GoldCompany"};
            foreach (var companyName in CompanyNameArray)
            {
                IWebElement CompanyNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[2]/input"));
                CompanyNameColumn.Clear();
                CompanyNameColumn.SendKeys(companyName); // �񤽥q�W��
                string companyNameErrorWordings = driver.FindElement(By.Id("companyName-error")).Text; // ����ˮֿ��~�T��
                string actualColumnCompanyName = CompanyNameColumn.GetAttribute("value"); // ������Ū�쪺��

                if(string.IsNullOrWhiteSpace(companyName) == true)
                {
                    if(companyNameErrorWordings == "������g")
                    {
                        PASS($"��J: {companyName}, ���~�T��: {companyNameErrorWordings}, �w�����~�T��: ������g");
                    }
                    else
                    {
                        FAIL($"��J: {companyName}, ���~�T��: {companyNameErrorWordings}, �w�����~�T��: ������g");
                    }
                }
                else if (string.IsNullOrWhiteSpace(companyName) != true && companyName.Length < 2)
                {
                    if(companyNameErrorWordings == "�̤�2�Ӧr")
                    {
                        PASS($"��J: {companyName}, ���~�T��: {companyNameErrorWordings}, �w�����~�T��: �̤�2�Ӧr");
                    }
                    else
                    {
                        FAIL($"��J: {companyName}, ���~�T��: {companyNameErrorWordings}, �w�����~�T��: �̤�2�Ӧr");
                    }
                }
                else if (companyName.Length >= 50)
                {
                    if(actualColumnCompanyName.Length == 50)
                    {
                        PASS($"��J: {companyName}, ��J����:{companyName.Length}, ������: {actualColumnCompanyName.Length}, �w��: ���׭���50�r��");
                    }
                    else
                    {
                        FAIL($"��J: {companyName}, ��J����:{companyName.Length}, ������: {actualColumnCompanyName.Length}, �w��: ���׭���50�r��");
                    }
                }
                else
                {
                    WARNING($"NeedCheck, ��J: {companyName}, ���~�T��: {companyNameErrorWordings}");
                }    
            }



            ///<summary>
            ///���q�νs����ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�: ���q�νs");
            string[] CompanyUniformNumbersArray = new string[] {"", "       ", "12345", "�������Ң��", "abcdefgh", "����r�[�^��Ʀr", "!@#$%^&*(", "1232148123123", "55665566", "@�^%^&��b�rc", "91301434666666", "91301434"};
            foreach (var companyUniformNumber in CompanyUniformNumbersArray)
            {
                IWebElement CompanyUniformNumbersColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]/input"));
                CompanyUniformNumbersColumn.Clear();
                CompanyUniformNumbersColumn.SendKeys(companyUniformNumber); // �񤽥q�νs
                string companyUniformNumbersErrorWordings = driver.FindElement(By.Id("vatNumber-error")).Text; // ����ˮֿ��~�T��
                string actualColumnCompanyUniformNumbers = CompanyUniformNumbersColumn.GetAttribute("value"); // ������Ū�쪺��

                if (string.IsNullOrWhiteSpace(companyUniformNumber) == true )
                {
                    if(companyUniformNumbersErrorWordings == "������g")
                    {
                        PASS($"��J: {companyUniformNumber}, ���~�T��: {companyUniformNumbersErrorWordings}, �w�����~�T��: ������g");
                    }
                    else
                    {
                        FAIL($"��J: {companyUniformNumber}, ���~�T��: {companyUniformNumbersErrorWordings}, �w�����~�T��: ������g");
                    }
                }
                else if (Tools.CheckUniformNumber(actualColumnCompanyUniformNumbers) != true)
                {
                    if(companyUniformNumbersErrorWordings == "�Τ@�s�����~")
                    {
                        PASS($"��J: {companyUniformNumber}, ��J����: {companyUniformNumber.Length}, �����Ū��: {actualColumnCompanyUniformNumbers}, ������: {actualColumnCompanyUniformNumbers.Length}, ���~�T��: {companyUniformNumbersErrorWordings}, �w�����~�T��: �Τ@�s�����~, �w������: 8 �r��");
                    }
                    else
                    {
                        FAIL($"��J: {companyUniformNumber}, ��J����: {companyUniformNumber.Length}, �����Ū��: {actualColumnCompanyUniformNumbers}, ������: {actualColumnCompanyUniformNumbers.Length}, ���~�T��: {companyUniformNumbersErrorWordings}, �w�����~�T��: �Τ@�s�����~, �w������: 8 �r��");
                    }
                }
                else if (Tools.CheckUniformNumber(actualColumnCompanyUniformNumbers) == true)
                {
                       PASS($"��J: {companyUniformNumber}, ��J����: {companyUniformNumber.Length}, �����Ū��: {actualColumnCompanyUniformNumbers}, ������: {actualColumnCompanyUniformNumbers.Length}, ���~�T��: {companyUniformNumbersErrorWordings}, �w�����~�T��: n/a, �w������: 8 �r��");
                }
            }



             ///<summary>
            ///�t�d�H����ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�: �t�d�H");
            string[] PrincipalArray = new string[] {"", "       ", "�Ϣ�", "ab", "�\�\�\����", "�\�\�\Bluce", "��", "��", "�^����뤤��Ң��Ÿ�#$%!@~&*^%", "����r�[�^��r�[�Ʀr_�ϢТѢҢӢԢբ֢עء�abcdefghij_1234567890_ABCDEFGHIJ_", "GoldPrinciple"};
            foreach (var Principal in PrincipalArray)
            {
                IWebElement PrincipalColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[3]/td[2]/input"));
                PrincipalColumn.Clear();
                PrincipalColumn.SendKeys(Principal); // ��t�d�H
                string principalColumnErrorWordings = driver.FindElement(By.Id("name-error")).Text; // ����ˮֿ��~�T��
                string actualColumnPrincipal = PrincipalColumn.GetAttribute("value"); // ������Ū�쪺��

                if(string.IsNullOrWhiteSpace(Principal) == true)
                {
                    if(principalColumnErrorWordings == "������g")
                    {
                        PASS($"��J: {Principal}, ���~�T��: {principalColumnErrorWordings}, �w�����~�T��: ������g");
                    }
                    else
                    {
                        FAIL($"��J: {Principal}, ���~�T��: {principalColumnErrorWordings}, �w�����~�T��: ������g");
                    }
                }
                else
                {
                    WARNING($"NeedCheck, ��J:{Principal}, ���~�T��: {principalColumnErrorWordings}");
                }
            }



            ///<summary>
            ///�p���H����ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�: �p���H");
            string[] ContactPersonArray = new string[] {"", "       ", "�Ϣ�", "ab", "�\�\�\����", "�\�\�\Bluce", "��", "��", "�^����뤤��Ң��Ÿ�#$%!@~&*^%", "����r�[�^��r�[�Ʀr_�ϢТѢҢӢԢբ֢עء�abcdefghij_1234567890_ABCDEFGHIJ_", "GoldPrinciple"};
            foreach (var contactPerson in ContactPersonArray)
            {
                IWebElement ContactPersonColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[4]/td[2]/input"));
                ContactPersonColumn.Clear();
                ContactPersonColumn.SendKeys(contactPerson); // ��t�d�H
                string contactPersonColumnErrorWordings = driver.FindElement(By.Id("contactPerson-error")).Text; // ����ˮֿ��~�T��
                string actualColumnContactPerson= ContactPersonColumn.GetAttribute("value"); // ������Ū�쪺��

                if(string.IsNullOrWhiteSpace(contactPerson) == true)
                {
                    if(contactPersonColumnErrorWordings == "������g")
                    {
                        PASS($"��J: {contactPerson}, ���~�T��: {contactPersonColumnErrorWordings}, �w�����~�T��: ������g");
                    }
                    else
                    {
                        FAIL($"��J: {contactPerson}, ���~�T��: {contactPersonColumnErrorWordings}, �w�����~�T��: ������g");
                    }
                }
                else
                {
                    WARNING($"NeedCheck, ��J:{contactPerson}, ���~�T��: {contactPersonColumnErrorWordings}");
                }
            }



            ///<summary>
            ///�s���q��-�ϽX-����ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�: �s���q��-�ϽX");
            string[] telephoneAreaCodeArray = new string[] {"", "       ", "12345", "�������Ң��", "abcdefgh", "����r�[�^��Ʀr", "!@#$%^&*(", "1232148123123", "55665566", "@�^%^&��b�rc", "��������", "12�G", "91301434", "3a45", "6789", "02"};
            foreach (var areaCode in telephoneAreaCodeArray)
            {
                bool telephoneAreaCodeCheck = Regex.IsMatch(areaCode, @"^[0-9]*$");

                IWebElement TelephoneAreaCodeColumn  = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[1]"));
                TelephoneAreaCodeColumn.Clear();
                TelephoneAreaCodeColumn.SendKeys(areaCode); 
                string telephoneArea_error = driver.FindElement(By.Id("companyTelArea-error")).Text; // ����ˮֿ��~�T��
                string actualColumnAreaCodes = TelephoneAreaCodeColumn.GetAttribute("value"); // ������Ū�쪺��

                if (string.IsNullOrWhiteSpace(areaCode) == true ) // ���J���ŭ�
                {
                    if(telephoneArea_error == "������g")
                    {
                        PASS($"��J: {areaCode}, ���~�T��: {telephoneArea_error}, �w�����~�T��: ������g");
                    }
                    else
                    {
                        FAIL($"��J: {areaCode}, ���~�T��: {telephoneArea_error}, �w�����~�T��: ������g");
                    }
                }
                else if (string.IsNullOrWhiteSpace(areaCode) != true && telephoneAreaCodeCheck != true) // ���J���O�ŭ� & ��J����"���Ʀr"
                {
                    if(telephoneArea_error == "�ȯ��J�Ʀr")
                    {
                        PASS($"��J: {areaCode}, ���~�T��: {telephoneArea_error}, �w�����~�T��: �ȯ��J�Ʀr");
                    }
                    else
                    {
                        FAIL($"��J: {areaCode}, ���~�T��: {telephoneArea_error}, �w�����~�T��: �ȯ��J�Ʀr");
                    }
                }
                else if (telephoneAreaCodeCheck ==true) // ���J�����Ʀr
                {
                    if (telephoneArea_error == "" && actualColumnAreaCodes.Length <=4) // ��S�����~�T�� �B �Ʀr���� < 4��
                    {
                         PASS($"��J: {areaCode}, ��J����: {areaCode.Length}, �����Ū��: {actualColumnAreaCodes}, ������: {actualColumnAreaCodes.Length}, ���~�T��: {telephoneArea_error}, �w�����~�T��: n/a, �w������: 4��ƥH��");
                    }
                    else
                    {
                        FAIL($"��J: {areaCode}, ��J����: {areaCode.Length}, �����Ū��: {actualColumnAreaCodes}, ������: {actualColumnAreaCodes.Length}, ���~�T��: {telephoneArea_error}, �w�����~�T��: n/a, �w������: 4��ƥH��");
                    }
                }
                else // �H�WCase�Ҥ��O��, �ݤH�ucheck
                {
                       WARNING($"NeedCheck, ��J: {areaCode}, �����Ū��: {actualColumnAreaCodes}, ������: {actualColumnAreaCodes.Length}, ���~�T��: {telephoneArea_error}");
                }
            }
            ///<summary>
            ///�s���q��-�D�X-����ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�: �s���q��-�D�X");
            string[] telephoneMainArray = new string[] {"", "       ", "12345", "�������Ң��", "ABCDEFG", "abcdefgh", "����r�[�^��Ʀr", "!@#$%^&*(", "����������������", "123������78", "55665566", "@�^%^&��b�rc", "��������", "12�G", "67890123", "3a4567b8", "6789", "2345678923456789", "8825252"};
            foreach (var telephoneMain in telephoneMainArray )
            {
                bool telephoneMainCheck = Regex.IsMatch(telephoneMain, @"^[0-9]*$");

                IWebElement TelephoneMainColumn  = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[5]/td[2]/input[2]"));
                TelephoneMainColumn.Clear();
                TelephoneMainColumn.SendKeys(telephoneMain); 
                string telephoneMain_error = driver.FindElement(By.Id("tel-error")).Text; // ����ˮֿ��~�T��
                string actualColumnTelephoneMain = TelephoneMainColumn.GetAttribute("value"); // ������Ū�쪺��

                if (string.IsNullOrWhiteSpace(telephoneMain) == true ) // ���J���ŭ�
                {
                    if(telephoneMain_error == "������g")
                    {
                        PASS($"��J: {telephoneMain}, ���~�T��: {telephoneMain_error}, �w�����~�T��: ������g");
                    }
                    else
                    {
                        FAIL($"��J: {telephoneMain}, ���~�T��: {telephoneMain_error}, �w�����~�T��: ������g");
                    }
                }
                else if (string.IsNullOrWhiteSpace(telephoneMain) != true && telephoneMainCheck != true) // ���J�����ŭȦ���J"�������Ʀr"
                {
                    if(telephoneMain_error == "�ȯ��J�Ʀr")
                    {
                        PASS($"��J: {telephoneMain}, ���~�T��: {telephoneMain_error}, �w�����~�T��: �ȯ��J�Ʀr");
                    }
                    else
                    {
                        FAIL($"��J: {telephoneMain}, ���~�T��: {telephoneMain_error}, �w�����~�T��: �ȯ��J�Ʀr");
                    }
                }
                else if (telephoneMainCheck ==true) // ���J�����Ʀr
                {
                    if (telephoneMain_error == "" && actualColumnTelephoneMain.Length <=8) // ��S�����~�T�� �B �Ʀr���� <= 8��
                    {
                         PASS($"��J: {telephoneMain}, ��J����: {telephoneMain.Length}, �����Ū��: {actualColumnTelephoneMain}, ������: {actualColumnTelephoneMain.Length}, ���~�T��: {telephoneMain_error}, �w�����~�T��: n/a, �w������: 8 ��ƥH��");
                    }
                    else
                    {
                        FAIL($"��J: {telephoneMain}, ��J����: {telephoneMain.Length}, �����Ū��: {actualColumnTelephoneMain}, ������: {actualColumnTelephoneMain.Length}, ���~�T��: {telephoneMain_error}, �w�����~�T��: n/a, �w������: 8 ��ƥH��");
                    }
                }
                else // �H�WCase�Ҥ��O��, �ݤH�ucheck
                {
                       WARNING($"NeedCheck, ��J: {telephoneMain}, �����Ū��: {actualColumnTelephoneMain}, ������: {actualColumnTelephoneMain.Length}, ���~�T��: {telephoneMain_error}");
                }
            }
            ///<summary>
            ///�s���q��-����-����ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�: �s���q��-����");
            string[] telephoneExtensionArray = new string[] {"", "       ", "12345", "�������Ң��", "ABCDEFG", "abcdefgh", "����r�[�^��Ʀr", "!@#$%^&*(", "����������������", "123������78", "55665566", "@�^%^&��b�rc", "��������", "12�G", "67890123", "3a4567b8", "6789", "2345678923456789", "8825252", "123", "131517131517", "9527"};
            foreach (var telephoneExtension in telephoneExtensionArray )
            {
                bool telephoneExtensionCheck = Regex.IsMatch(telephoneExtension, @"^[0-9]*$");

                TelephoneExtensionColumn.Clear();
                TelephoneExtensionColumn.SendKeys(telephoneExtension); 
                string telephoneExt_error = driver.FindElement(By.Id("telext-error")).Text; // ����ˮֿ��~�T��
                string actualColumnTelephoneExtension = TelephoneExtensionColumn.GetAttribute("value"); // ������Ū�쪺��

                if (string.IsNullOrWhiteSpace(telephoneExtension) == true ) // ���J���ŭ�
                {
                    if(telephoneExt_error == "������g")
                    {
                        PASS($"��J: {telephoneExtension}, ���~�T��: { telephoneExt_error }, �w�����~�T��: n/a (�D����)");
                    }
                    else
                    {
                        FAIL($"��J: {telephoneExtension}, ���~�T��: { telephoneExt_error }, �w�����~�T��: n/a (�D����)");
                    }
                }
                else if (string.IsNullOrWhiteSpace(telephoneExtension) != true && telephoneExtensionCheck  != true) // ���J�����ŭȦ���J"�������Ʀr"
                {
                    if(telephoneExt_error  == "�ȯ��J�Ʀr")
                    {
                        PASS($"��J: {telephoneExtension}, ���~�T��: {telephoneExt_error }, �w�����~�T��: �ȯ��J�Ʀr");
                    }
                    else
                    {
                        FAIL($"��J: {telephoneExtension}, ���~�T��: {telephoneExt_error }, �w�����~�T��: �ȯ��J�Ʀr");
                    }
                }
                else if ( telephoneExtensionCheck == true  && telephoneExtension.Length < 6 )  // ���J�����Ʀr �B �Ʀr���� < 6��
                {
                    if (telephoneExt_error == "") // ��S�����~�T����
                    {
                        PASS($"��J: {telephoneExtension}, ��J����: {telephoneExtension.Length}, �����Ū��: {actualColumnTelephoneExtension}, ������: {actualColumnTelephoneExtension.Length}, ���~�T��: {telephoneExt_error }, �w�����~�T��: n/a, �w������: 6 ��ƥH��");
                    }
                    else
                    {
                        FAIL($"��J: {telephoneExtension}, ��J����: {telephoneExtension.Length}, �����Ū��: {actualColumnTelephoneExtension}, ������: {actualColumnTelephoneExtension.Length}, ���~�T��: {telephoneExt_error }, �w�����~�T��: n/a, �w������: 6 ��ƥH��");
                    }
                }
                else if (telephoneExtensionCheck ==true  &&  telephoneExtension.Length >= 6) // ���J�����Ʀr �B �Ʀr���� >= 6��
                {
                     if (actualColumnTelephoneExtension.Length == 6) // ������ڪ��׬�6��
                    {
                        PASS($"��J: {telephoneExtension}, ��J����: {telephoneExtension.Length}, �����Ū��: {actualColumnTelephoneExtension}, ������: {actualColumnTelephoneExtension.Length}, ���~�T��: {telephoneExt_error }, �w�����~�T��: n/a, �w������: 6 ��ƥH��");
                    }
                    else
                    {
                        FAIL($"��J: {telephoneExtension}, ��J����: {telephoneExtension.Length}, �����Ū��: {actualColumnTelephoneExtension}, ������: {actualColumnTelephoneExtension.Length}, ���~�T��: {telephoneExt_error }, �w�����~�T��: n/a, �w������: 6 ��ƥH��");
                    }
                }
            }



            ///<summary>
            /// ��ʹq�ܸ��X����ˮ�
            ///</summary>
            INFO("");
            INFO("����ˮ�:��ʹq�ܸ��X");
            string[] phones = new string[] { "", "          ","!@#$%^", "�\�\�\", "ABCDEF", "0897654321", "A987654321", "��987654321", "098765432��", "0987 543 1", "0901234567890", "0910313414" };
            foreach (var input in phones)
            {
                // int string_length = System.Text.Encoding.Default.GetBytes(input).Length; //  �^���r�����bytes (UTF-8�з�, �b�έ^�Ʀr = 1 byte, ����&���έ^�Ʀr = 3 bytes)
                bool numericCheck = Regex.IsMatch(input, @"^[0-9]*$"); // �P�_��J�O�_�����Ʀr
                bool cellPhoneNumberCheck = Regex.IsMatch(input, @"^09\d{8}$"); // �P�_��J�O�_����ʹq��09�}�Y�榡 (���h��ܦ�: 09�}�Y�᭱8�X�Ʀr) 
                
                IWebElement CellphoneColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[6]/td[2]/input"));
                CellphoneColumn.Clear();
                CellphoneColumn.SendKeys(input);
                string cellphone_error = driver.FindElement(By.Id("cellPhone-error")).Text; // ����ˮֿ��~�T��
                string actualColumnCellphone = CellphoneColumn.GetAttribute("value"); // �����Ū�쪺��
                 
                 if (string.IsNullOrWhiteSpace(input) == true ) // ���J���ŭ�
                {
                    if(cellphone_error== "������g")
                    {
                        PASS($"��J: {input}, ���~�T��: {cellphone_error}, �w�����~�T��: ������g");
                    }
                    else
                    {
                        FAIL($"��J: {input}, ���~�T��: {cellphone_error}, �w�����~�T��: ������g");
                    }
                }
                else if (numericCheck != true) // ���J"�������Ʀr"
                {
                    if(cellphone_error == "��ʹq�ܮ榡���~")
                    {
                        PASS($"��J: {input}, ���~�T��: {cellphone_error}, �w�����~�T��: ��ʹq�ܮ榡���~");
                    }
                    else
                    {
                        FAIL($"��J: {input}, ���~�T��: {cellphone_error}, �w�����~�T��: ��ʹq�ܮ榡���~");
                    }
                }
                else if (numericCheck == true && cellPhoneNumberCheck != true) // ���J�����Ʀr �� �D��ʹq�ܮ榡
                {
                    if (cellphone_error == "��ʹq�ܮ榡���~") 
                    {
                        PASS($"��J: {input}, ���~�T��: {cellphone_error}, �w�����~�T��: ��ʹq�ܮ榡���~");
                    }
                    else
                    {
                        FAIL($"��J: {input}, ���~�T��: {cellphone_error}, �w�����~�T��: ��ʹq�ܮ榡���~");
                    }
                }
                 else if (cellPhoneNumberCheck == true)
                {
                    PASS($"��J: {input}, ��J����: {input.Length}, �����Ū��: {actualColumnCellphone}, ������: {actualColumnCellphone.Length}, ���~�T��: {cellphone_error}, �w�����~�T��: n/a, �w������: 09�}�Y��10��Ʀr");
                }
                else // �H�WCase�Ҥ��O��, �ݤH�ucheck
                {
                       WARNING($"NeedCheck, ��J: {input}, ��J����: {input.Length}, �����Ū��: {actualColumnCellphone}, ������: {actualColumnCellphone.Length}, ���~�T��: {cellphone_error}, �w�����~�T��: n/a, �w������: 09�}�Y��10��Ʀr");
                }
            }



            ///<summary>
            /// �����V�x�M�פU�Կ��
            ///</summary>
            INFO("");
            INFO("�ˬd:�V�x�M�פU�Կ��ﶵ");
            int rescueProjectCounts = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li")).Count; // �ﶵ�ƶq
            for (int i = 2; i < rescueProjectCounts; i++)
			{
                IWebElement RescueProjectDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li")); // �V�x�M�פU�Կ��
                RescueProjectDropDownList.Click();
                IWebElement SelectResuceProject = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[7]/td[2]/ul/li/ul/li[{i}]"));
                SelectResuceProject.Click();
                WARNING($"�V�x��� {i - 1}: {RescueProjectDropDownList.Text}");
			}



            ///<summary>
            /// ����&�������
            ///</summary>
            INFO("");
            INFO("�ˬd:����&����U�Կ��ﶵ");
            IWebElement Country_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li"));
            Country_DropDownList.Click(); //�i�}"�п�ܿ���"�U�Կ��
            IWebElement SelectCountry = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[1]/li/ul/li[2]/span"));
            SelectCountry.Click(); // �I��"�x�_��"
            IWebElement Branch_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]"));
            Branch_DropDownList.Click(); // �i�}"����"�U�Կ��
            IWebElement SelectBranch = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[8]/td[2]/div/ul[2]/li/ul/li[2]/span"));
            SelectBranch.Click(); //�I��"��~�B"



            ///<summary>
            /// �Y����&�ɶ����
            ///</summary>
            INFO("");
            INFO("�ˬd:�Y����&�ɶ��U�Կ��ﶵ");
            Tools.ScrollPageUpOrDown(driver, 700);
             int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li")).Count; // ����Y�����U�Կ��ﶵ�ƶq
             int time_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li")).Count; // ����Y��ɶ��U�Կ��ﶵ�ƶq

            for (int m= 2; m < date_amount; m++)
			{
                IWebElement DateDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li")); 
                DateDropDownList.Click();  // �i�}�Y�����U�Կ��
                IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[1]/li/ul/li[{m}]/span")); 
                SelectDate.Click(); // �I��@��"���"
                WARNING($"�� {m - 1} �ӥi����: {DateDropDownList.Text}");
			}
             for (int n= 2; n < time_amount; n++)
			{
                IWebElement TimeDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li"));
                TimeDropDownList .Click(); // �i�}"�ɬq"�U�Կ��
                IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]/div/ul[2]/li/ul/li[{n}]/span"));
                SelectTime.Click(); // �I��@��"�ɬq"
                WARNING($"�� {n - 1} �ӥi����: {TimeDropDownList.Text}");
			}



            ///<summary>
            /// �ˬd "�w���Y��ɶ�" ������r
            ///</summary>
           INFO("");
           INFO("�ˬd:�w���Y��ɶ�������r");
           string appointment_visit_time_column = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[9]/td[2]")).Text;
           string appointment_visit_time_information = appointment_visit_time_column.Substring(appointment_visit_time_column.IndexOf("��"));
            if (appointment_visit_time_information == "�������ɪA�ȫ~��öi���U�ȻY��A�Ȥ��y�A���A�ȶȨѨC�H�C���w��1�Ӯɬq�A�y�����K�A�бz���̡C")
            {
                PASS("�ŦX�w��");
            }
            else
            {
                FAIL($"���~�T��: {appointment_visit_time_information}, �w�����~�T��: �������ɪA�ȫ~��öi���U�ȻY��A�Ȥ��y�A���A�ȶȨѨC�H�C���w��1�Ӯɬq�A�y�����K�A�бz���̡C");
            }


            ///<summary>
            /// �ˬd "�ڤw�\Ū�æP�N" ���~�T��
            ///</summary>
           INFO("");
           INFO("�ˬd:�ڤw�\Ū�æP�N���~�T��");
           IWebElement IHaveReadRadioButton = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[11]/td[2]/div/a"));
           
        rerun_ihaveread:
            string i_have_read_status = IHaveReadRadioButton.GetAttribute("class"); // �ڤw�\Ū radio button ���A (checked: �w�Ŀ�)
            string i_have_read_error = driver.FindElement(By.Id("apply-error")).Text; // �ڤw�\Ū�����~�T��
            if (i_have_read_status != "checked")
            { 
                if (i_have_read_error == "������g")
                {
                    PASS($"���Ŀ���~�T��: {i_have_read_error}, �w�����~�T��: ������g");
                }
                else
                {
                    FAIL($"���Ŀ���~�T��: {i_have_read_error}, �w�����~�T��: ������g");
                }
                IHaveReadRadioButton.Click(); // �I"�ڤw�\Ū" radio button
                 goto rerun_ihaveread;
            }
            else
            {
                if (i_have_read_error == "")
                {
                    PASS($"�w�Ŀ���~�T��: {i_have_read_error}, �w�����~�T��: n/a");
                }
                else
                {
                    FAIL($"�w�Ŀ���~�T��: {i_have_read_error}, �w�����~�T��: n/a");
                }
            }


            ///<summary>
            /// �ˬd "�ϫ����ҽX" ���~�T��
            /// </summary>
            INFO("");
            INFO("�ˬd: �ϫ����ҽX���~�T��");
            Tools.ScrollPageUpOrDown(driver, 800);
            IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath("//*[@id='captchaValue']")); // ���ҽX��J���
        rerun_CaptchaSteps:
            string image_verification_column_value = driver.FindElement(By.Id("captchaValue")).GetAttribute("value"); // ���ҽX�����Ū�쪺��
            string image_verification_error = driver.FindElement(By.Id("captchaWrong")).Text; // ���ҽX���~�T��
            if (string.IsNullOrWhiteSpace(image_verification_column_value) == true)
            {
                if (image_verification_error == "�п�J���ҽX")
                {
                     PASS($"����J, ���~�T��: {image_verification_error}, �w�����~�T��: �п�J���ҽX");
                }
                else
                {
                    FAIL($"����J, ���~�T��: {image_verification_error}, �w�����~�T��: �п�J���ҽX");
                }    
                ImageVerificationCodeColumn.SendKeys("9527");
                SubmitButton.Click();
                goto rerun_CaptchaSteps;
            }
            else
            {
                if (image_verification_error == "���ҽX���~")
                {
                     PASS($"��J���~�����ҽX, ���~�T��: {image_verification_error}, �w�����~�T��: ���ҽX���~");
                }
                else
                {
                    FAIL($"FAIL, ��J���~�����ҽX, ���~�T��: {image_verification_error}, �w�����~�T��: ���ҽX���~");
                }  
            }



            ///<summary>
            ///�ˬd�Ӹ�k�i���ƶ��W�s��
            ///</summary>
            INFO("");
            INFO("�ˬd: �Ӹ�k�i���ƶ��W�s��");
            string personal_indormation_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/div[2]/table/tbody/tr[11]/td[2]/a")).GetAttribute("href"); // ����Ӹ�k�i���ƶ����W�s��
            if (personal_indormation_hyperlink == "https://www.esunbank.com.tw/bank/about/announcement/privacy/privacy-statement")
            {
                PASS("�Ӹ�k�i���ƶ��W�s�� �ŦX�w��");
            }
            else
            {
                FAIL($"�Ӹ�k�i���ƶ� link: {personal_indormation_hyperlink}, �w�� link: https://www.esunbank.com.tw/bank/about/announcement/privacy/privacy-statement");
            }


            ///<summary>
            ///�ˬd�U���װϴ��ܤ��
            ///</summary>
            INFO("");
            INFO("�ˬd: �U���װϴ��ܤ��");
            string promot_wordings = driver.FindElement(By.XPath("//*[@id='mainform']/div[7]/div[3]/div[2]/div[1]/span[3]")).Text;
            if (promot_wordings == "�������z�A�I���u�T�{�v�e�X�w����A�Ш̹w���e�������{�d�ӽСA�Y�z����w���ɶ��Ӧ�A���t�X�F�����̱��I�A�ФŪ����ܤ���ӽСC")
            {
                PASS("�U���װϴ��ܤ�� �ŦX�w��");
            }
            else
            {
                FAIL($"�U���װϴ��ܤ��:  ��ڱԭz: {promot_wordings}, �w���ԭz: �������z�A�I���u�T�{�v�e�X�w����A�Ш̹w���e�������{�d�ӽСA�Y�z����w���ɶ��Ӧ�A���t�X�F�����̱��I�A�ФŪ����ܤ���ӽСC");
            }

            CloseBrowser();
        }
    }
}


