using System;
using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using System.IO;
using System.Linq;
using CsvHelper;
using System.Globalization;
using Xunit.Abstractions;


namespace AutomatedTest.IntegrationTest.About.MessageBoard
{
    public class �X�ȯd����������Ʒ��վ�_PC:IntegrationTestBase
    {
        public �X�ȯd����������Ʒ��վ�_PC (ITestOutputHelper output, Setup testsetup): base(output, testsetup)
        {
            testurl = domain + "https://easyfee.esunbank.com.tw/index.action";
        }
            
        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void �����������ˮ�(string browser)
        {
            StartTestCase(browser, "�����������ˮ�", "York");
            INFO("�����������Snapshot + �XReport");

            var aaa = driver.FindElement(By.CssSelector(".page_frame3")).Text;



            int country_dropdownList_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li")).Count;  //��� "����"�U�Կ��̪�������

            for (int country_index = 1; country_index <= country_dropdownList_amount - 1; country_index++) // ��l�� = 1
            {
               

                IWebElement CountryDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span")); //  "����" �U�Կ��
                CountryDropDownList.Click(); // �i�}�������

                string cuntry_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{country_index + 1}]/span"; // ����Xpath
                IWebElement SelectCountry = driver.FindElement(By.XPath(cuntry_xpath)); //  "����" �U�Կ��
                SelectCountry.Click(); //�I���country_index�ӿ���

                int branch_dropdownlist_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li")).Count;  //��� "��U����"�U�Կ��̪������

                for (int branch_index = 1; branch_index <= branch_dropdownlist_amount; branch_index++) // ��l�� = 1
                {
                    if (country_index == country_dropdownList_amount - 1 && branch_index == 14) 
                    {
                        branch_index++;
                    }

                    IWebElement BranchDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span")); //  "����" �U�Կ��
                    BranchDropDownList.Click(); // �i�}������

                    string branch_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{branch_index}]/span";
                    IWebElement SelectBranch = driver.FindElement(By.XPath(branch_xpath)); //  "����" �U�Կ��
                    SelectBranch.Click(); //�I���branch_index�ӿ���


                    // �Ҧ����w��
                    IWebElement FullNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[1]/td[2]/input"));
                    IWebElement TelephoneColumn = driver.FindElement(By.XPath("//*[@id='phone']"));
                    IWebElement IdentityCardColumn = driver.FindElement(By.XPath("//*[@id='citizenId']"));
                    IWebElement EMailColumn = driver.FindElement(By.XPath("//*[@id='email']"));
                    IWebElement BusinessItem_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/span"));
                    IWebElement CreditCardBusiness = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/ul/li[2]/span"));
                    IWebElement MessageArea = driver.FindElement(By.XPath("//*[@id='layout_0_maincontent_2_comments']"));
                    IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]/a"));
                    IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));



                    ///<summary>
                    ///�qExcel�ɤ�Ū�XUserData
                    ///</summary>
                    string csvpath = $@"{UserDataList.Upperfolderpath}\testdata\UserInfo.csv";
                    using (var reader = new StreamReader(csvpath))
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        var records = csv.GetRecords<UserDataList>().ToList();
                        Random csv_row = new Random();
                        int user_data_index = 1;
                        foreach (var user in records)
                        {
                            if (user_data_index == csv_row.Next(1, 7)) // �� CSV �� random����
                            {
                                FullNameColumn.Clear();
                                FullNameColumn.SendKeys(user.NAME); //��m�W

                                IdentityCardColumn.Clear();
                                IdentityCardColumn.SendKeys(user.ID); //�񨭤���

                                TelephoneColumn.Clear();
                                TelephoneColumn.SendKeys(user.PHONE); //��s���q��

                                EMailColumn.Clear();
                                EMailColumn.SendKeys(user.EMAIL); //�� E-MAIL
                                
                                break;
                            }
                            user_data_index++;
                        }

                        string snapshotpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}\SnapshotFolder\GuestMessageBoard";
                        TestBase.CreateFolder(snapshotpath);
                        WARNING(TestBase.ElementSnapShotToReport(driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]"))));

                        //Tools.ElementTakeScreenShot(driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]")),
                        //    $@"{snapshotpath}\�� {country_index} �ӿ����� {branch_index} �Ӥ���_���snapshot.png"); //�����I��
                        //System.Threading.Thread.Sleep(100);

                        //BusinessItem_DropDownList.Click(); //�I�d���~�ȤU�Կ��
                        //CreditCardBusiness.Click(); //��H�Υd�~��


                        //MessageArea.Clear();
                        //MessageArea.SendKeys(
                        //    "[Automated Testing on " + browser + " in " + Version + " ]" + "\r\n"
                        //    + "�� " + country_index + " �ӿ����� " + branch_index + " �Ӥ���" + "\r\n"
                        //    + "=================================================" + "\r\n"
                        //    + "Total 21 ��, 154 ����" + "\r\n"
                        //    + "����(�a��index)(����ƶq):" + "\r\n"
                        //    + "�򶩥�(1)(1),      �O�_��(2)(35),      �s�_��(3)(32),      ��饫(4)(11),      �s�˥�(5)(3),      "
                        //    + "�s�˿�(6)(3),      �]�߿�(7)(3),        �O����(8)(14),      ���ƿ�(9)(2),        �n�뿤(10)(1)," + "\r\n"
                        //    + "���L��(11)(1),    �Ÿq��(12)(2),      �Ÿq��(13)(1),      �x�n��(14)(10),    ������(15)(14),      "
                        //    + "�̪F��(16)(2),    �y����(17)(1),      �Ὤ��(18)(1),      �O�F��(19)(1),      ���(20)(1),       ���~�a��(21)(15)."
                        //    ); // �� "�d�����e"

                        //Tools.SCrollToElement(driver, TelephoneColumn);
                        //Tools.TakeScreenShot($@"{snapshotpath}\�� {country_index} �ӿ����� {branch_index} �Ӥ��� snapshot_{browser}.png", driver); // snapshot��U�e��

                        // IHaveReadRadioButtin.Click(); // �I "�ڤw�\Ū" radio button

                        // SubmitButton.Click(); // �I "�e�X" ���s
                        // System.Threading.Thread.Sleep(300);

                        // ElementTakeScreenShot(driver, driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]")), "d:\\" + browserType + timepath + ".png");
                        
                        // WebDriverWait wait_to_see_popsup_window = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                        // wait_to_see_popsup_window.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("/html/body/div[5]/div/div/a"))); // ���ݪ���ݨ�q������ 

                        // driver.FindElement(By.XPath("/html/body/div[5]/div/div/a")).Click(); // �I�q������ "X" ���s
                        // System.Threading.Thread.Sleep(100);
                        // System.Diagnostics.Debug.WriteLine("15. �����q������");

                        // WebDriverWait wait_to_see_submit_button = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
                        // wait_to_see_submit_button.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id='submit']"))); // ���ݪ���ݨ�"�d����"
                        // System.Threading.Thread.Sleep(100);

                       // driver.FindElement(By.XPath("//*[@id='scrollUp']")).Click(); //�e���^�쳻��
                       // System.Threading.Thread.Sleep(500);
                    }
                }
            }
            CloseBrowser();
        }
    }
}
    

