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


namespace LaborReliefLoanAppointmentServiceTest
{
    public class �Ҥu�V�x�U�ڹw���A��_ErrorHandling_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/loan/tools/apply/labor-loan-appointment#";


        [Theory]
        [InlineData(BrowserType.Chrome)]
        [InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                driver.Manage().Window.Maximize();


                for (int i = 2; i <= 21; i++) // initial i =2
                {

                    int j = 1; // initial j = 1
                    int[] arrray = new int[] { 35, 1, 31, 1, 3, 3, 11, 3, 14, 2, 1, 2, 1, 1, 10, 14, 1, 2, 1, 1 }; //�������������
                    while (j <= arrray[i - 2])
                    {
                        string timesavepath = System.DateTime.Now.ToString("yyyy-MM-dd_hhmm");

                        // �w�q���XPath
                        string Cuntry_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{i}]/span";
                        string Branch_Xpath = $"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{j + 1}]/span";

                        IWebElement SubmitButton = driver.FindElement(By.XPath(LaborReliefLoan_XPath.submit_button_XPath()));
                        SubmitButton.Click(); // �I"�T�{"button




                        ///<summary>
                        /// �ˬd�����Whyperlink�O�_���T
                        ///</summary>
                        string LaborReliefOnlineApplication_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[2]/a")).GetAttribute("href");
                        Assert.Equal("https://www.esunbank.com.tw/s/PersonalLoanApply/" +
                            "Landing/IDConfirm?MKP=eyJNS0RQVCI6bnVsbCwiTUtFSUQiOm51bGwsIk1LUFJOIjoiUDAwMDAwMjEiLCJNS1BKTiI6IkowMDAwMDM0IiwiTUtQSUQiOm51bGx9", 
                            LaborReliefOnlineApplication_hyperlink);
                        
                        string PersonalInformation_hyperlink = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[8]/td[2]/a")).GetAttribute("href");
                        Assert.Equal("https://www.esunbank.com.tw/bank/about/announcement/privacy/privacy-statement", PersonalInformation_hyperlink);




                        ///<summary>
                        /// �m�W��� (�ثe�L�ˮ�)
                        ///</summary>
                        IWebElement FullNameColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.name_column_Xpath()));
                        FullNameColumn.Clear();
                        FullNameColumn.SendKeys($"�����H{i - 1}_{j}"); // ��m�W




                        ///<summary>
                        /// �����Ҧr������ˮ�
                        ///</summary>
                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.ID_column_XPath()));

                        Random random = new Random();
                        int random_number = random.Next(1, 21); //�����H��1-21���Ʀr
                        bool sex = false;
                        if (random_number % 2 == 0)
                            sex = true; // for �����Ҧr��tools �ϥ�(�k/�k)

                        string[] IDNumbers = new string[] { "", "!@#%^", "�\�\�\", "ABCDEF", "0897654321", "A987654321", 
                            "�颰����������������", "��123456789", "a12345678��", "A800000014", "AD12544441", "a123456788", "A123456788",
                            "a800000014", "aD12544441", "a123456788", "Ad30341957", "a941062183", "A970000026", "AD30341957", "A941062183", 
                            $"{Tools.CreateRandomString(10)}", $"{Tools.CreateIDNumber(sex, random_number)}" };
                        // �˼ƲĤG�լ��H������10�j�p�g�^��+�Ʀr�զX
                        // �̫�@�ճz�L�����Ҧr�����;����ͲŦX�W��r��

                        foreach (var input in IDNumbers)
                        {
                            IdentityCardColumn.Clear();
                            bool ResidentIDNumberCheck = Regex.IsMatch(input.ToUpper(), @"^[A-Z]{1}[0-9]{9}$"); // �ϥΥ��h��ܦ�: ����榡 [A~Z] {1}�ӼƦr +  [0~9] {9}�ӼƦr
                            bool ForeignerIDNumberCheck = Regex.IsMatch(input.ToUpper(), @"^[A-Z]{1}[A-D8-9]{1}[0-9]{8}$");

                            IdentityCardColumn.SendKeys(input);
                            string id_error = driver.FindElement(By.Id("citizenId-error")).Text;
                            if (input == "")
                            {
                                Assert.Equal("������g", id_error);
                            }
                            else if (ResidentIDNumberCheck != true && ForeignerIDNumberCheck != true)
                            {
                                Assert.Equal("�п�J���Ī������Ҧr��", id_error);
                            }
                            else if (ResidentIDNumberCheck == true && Tools.CheckResidentID(input.ToUpper()) != true)
                            {
                                Assert.Equal("�п�J���Ī������Ҧr��", id_error);
                            }
                            else if (ForeignerIDNumberCheck == true && Tools.CheckForeignerID(input.ToUpper()) != true)
                            {
                                Assert.Equal("�п�J���Ī������Ҧr��", id_error);
                            }
                            else if (ForeignerIDNumberCheck == true && Tools.CheckForeignerID(input.ToUpper()) == true)
                            {
                                Assert.Equal("���A�ȥثe�ȭ�����H�ӽ�", id_error);
                            }
                            else
                                Assert.Equal("", id_error);
                        }




                        ///<summary>
                        /// �q�ܸ��X����ˮ�
                        ///</summary>
                        IWebElement CellPhoneColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.cellphone_column_XPath()));
                        string[] phones = new string[] { "ABCDEF", "0897654321", "", $"{Tools.CreateRandomNumber(10)}", $"{Tools.CreateCellPhoneNumber()}" };
                        // �˼ƲĤG��: �H�����ͪ���10�Ʀr
                        // �̫�@��: �z�L��ʹq�ܸ��X���;����ͲŦX�W�渹�X
                        foreach (var input in phones)
                        {
                            int string_length = System.Text.Encoding.Default.GetBytes(input).Length; //  �^���r�����bytes (UTF-8�з�, �b�έ^�Ʀr = 1 byte, ����&���έ^�Ʀr = 3 bytes)

                            CellPhoneColumn.Clear();
                            bool CellPhoneNumberCheck = Regex.IsMatch(input, @"^09\d{8}$"); // ���h��ܦ�: 09�}�Y�᭱8�X�Ʀr
                            CellPhoneColumn.SendKeys(input);
                            string cellphone_error = driver.FindElement(By.Id("cellPhone-error")).Text;
                            if (input == "")
                            {
                                Assert.Equal("������g", cellphone_error);
                            }
                            else if (CellPhoneNumberCheck != true)
                            {
                                Assert.Equal("��ʹq�ܮ榡���~", cellphone_error);
                            }
                            else if (CellPhoneNumberCheck == true && string_length != 10) // 10����b�μƦr = 10 bytes����
                            {
                                Assert.Equal("��ʹq�ܮ榡���~", cellphone_error);
                            }
                            else
                            {
                                Assert.Equal("", cellphone_error);
                            }
                        }




                        ///<summary>
                        /// �ˬd"�ڤw�\Ū�æP�N" ���~�T��
                        ///</summary>
                        IWebElement IHaveReadRadioButton = driver.FindElement(By.XPath(LaborReliefLoan_XPath.i_have_read_button_XPath()));

                    rerun_ihaveread:
                        string i_have_read_status = IHaveReadRadioButton.GetAttribute("class");
                        string i_have_read_error = driver.FindElement(By.Id("apply-error")).Text;
                        if (i_have_read_status != "checked")
                        {
                            Assert.Equal("������g", i_have_read_error);
                            IHaveReadRadioButton.Click(); // �I"�ڤw�\Ū" radio button
                            goto rerun_ihaveread;
                        }
                        else
                        {
                            Assert.Equal("", i_have_read_error);
                        }




                        ///<summary>
                        /// �X�ͦ~������
                        ///</summary>
                        IWebElement BirthdayCalendarIcon = driver.FindElement(By.Id("datepicker1-button"));
                        BirthdayCalendarIcon.Click(); // �I�}�X�ͦ~���p��� 
                        Random Year = new Random();
                        Random Month = new Random();
                        // IWebElement Year_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[1]"));
                        IWebElement SelectYear = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[1]/select/option[{Year.Next(1, 122)}]")); // �H���~��
                        SelectYear.Click();
                        //  IWebElement Month_dropdownlist = driver.FindElement(By.XPath("/html/body/div[3]/div/div[2]"));
                        IWebElement SelectMonth = driver.FindElement(By.XPath($"/html/body/div[3]/div/div[2]/select/option[{Month.Next(1, 13)}]")); // �H�����
                        SelectMonth.Click(); // ����

                    rerun_date:
                        Random week = new Random();
                        Random day = new Random();
                        IWebElement date = driver.FindElement(By.XPath($"/html/body/div[3]/table/tbody/tr[{week.Next(1, 6)}]/td[{day.Next(1, 8)}]")); // �H�����
                        string date_value = date.GetAttribute("class");
                        if (date_value == "is-empty") // �P�_����쬰�Ů�(��ܷ��S���o��l), ���s�����H�����
                        {
                            goto rerun_date;
                        }
                        else
                        {
                            date.Click();
                        }




                        ///<summary>
                        /// ����&�������
                        ///</summary>
                        IWebElement Country_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.country_dropdownlist_XPath()));
                        Country_DropDownList.Click(); //�i�}"�п�ܿ���"�U�Կ��
                        IWebElement SelectCountry = driver.FindElement(By.XPath(Cuntry_Xpath));
                        SelectCountry.Click(); // �I��@��"����"
                        IWebElement Branch_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.branch_dropdownlist_XPath()));
                        Branch_DropDownList.Click(); // �i�}"����"�U�Կ��
                        IWebElement SelectBranch = driver.FindElement(By.XPath(Branch_Xpath));
                        SelectBranch.Click(); //�I��@��"����"

                        IWebElement Date_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.date_dropdownlist_XPath()));
                        int date_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li")).Count; // ����Y�����U�Կ��ﶵ�ƶq
                        int time_amount = driver.FindElements(By.XPath("//*[@id='mainform'']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li")).Count; // ����Y��ɶ��U�Կ��ﶵ�ƶq

                        for (int m = 2; m <= date_amount; m++)
                        {
                            for (int n = 2; n <= time_amount; n++)
                            {
                                
                                Date_DropDownList.Click(); //�i�}"�Y����"�U�Կ��
                                IWebElement SelectDate = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li/ul/li[{m}]/span"));
                                SelectDate.Click(); // �I��@��"���"
                                IWebElement Time_DropDownList = driver.FindElement(By.XPath(LaborReliefLoan_XPath.time_dropdownlist_XPath()));
                                Time_DropDownList.Click(); // �i�}"�ɬq"�U�Կ��
                                IWebElement SelectTime = driver.FindElement(By.XPath($"//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li/ul/li[{n}]/span"));
                                SelectTime.Click(); // �I��@��"�ɬq"

                                string ErrorMessage = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/div[1]")).Text;
                                if (ErrorMessage == "") // �����message����, �Y"�Ӥ��榹�ɬq�W�B�w���A�п�ܨ�L�ɬq�C"
                                {
                                    string snapshotpath = System.AppDomain.CurrentDomain.BaseDirectory + "SnapshotFolder\\LaborReliefLoan";
                                    Tools.CreateSnapshotFolder(snapshotpath);
                                    System.Threading.Thread.Sleep(100);
                                    Tools.TakeScreenShot($@"{snapshotpath}\��{i - 1}������{j}����_��{m-1}���{n-1}�ɬq.png", driver); //��@�I��
                                }
                            }
                        }




                        ///<summary>
                        /// �ˬd "�w���Y��ɶ�" ������r
                        ///</summary>
                        string appointment_visit_time_column = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]")).Text;
                        string appointment_visit_time_information = appointment_visit_time_column.Substring(appointment_visit_time_column.IndexOf("��"));
                        Assert.Equal("�������ɪA�ȫ~��öi���U�ȻY��A�Ȥ��y�A���A�ȶȨѨC�H�C���w��1�Ӯɬq�A�y�����K�A�бz���̡C", appointment_visit_time_information);




                       ///<summary>
                       /// �ˬd "�ϫ����ҽX" ���~�T��
                       ///</summary>
                    rerun_imageverification:
                        IWebElement ImageVerificationCodeColumn = driver.FindElement(By.XPath(LaborReliefLoan_XPath.image_verify_code_column_XPath()));
                        string image_verification_column_value = driver.FindElement(By.Id("captchaValue")).GetAttribute("value");

                        if (image_verification_column_value == "")
                        {
                            string image_verification_error = driver.FindElement(By.Id("captchaWrong")).Text;
                            Assert.Equal("�п�J���ҽX", image_verification_error);
                            ImageVerificationCodeColumn.SendKeys("9527");
                            goto rerun_imageverification;
                        }
                        else
                        {
                            SubmitButton.Click();
                            string image_verification_error = driver.FindElement(By.Id("captchaWrong")).Text;
                            Assert.Equal("���ҽX���~", image_verification_error);
                        }




                        ///<summary>
                        /// �ˬd "�ϫ����ҽX" ���~�T��
                        ///</summary>
                        string promot_wordings = driver.FindElement(By.XPath("//*[@id='mainform'']/div[9]/div[3]/div[2]/div/div[4]/div")).Text;
                        Assert.Equal("�������z�A�I���u�T�{�v�e�X�w����ݧ����u�W�ӽЮѶ�g�A����²�T�q����~�⦨�\�����w���A�ȡA" +
                            "�ýбz��w���ɶ��e�������{�d�ӿ�C�Y�z�����즨�\�w��²�T�A���t�X�F�����̱��I�A�ФŪ����ܤ���ӽСC", promot_wordings);

                        j++;
                    }
                }
                driver.Quit();
            }
        }
    }
}


