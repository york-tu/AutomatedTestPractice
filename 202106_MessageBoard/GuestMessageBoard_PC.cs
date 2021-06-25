using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using Xunit;
using Utilities;
using System.IO;
using System.Linq;
using CsvHelper;
using System.Globalization;


namespace GuestMessageBoardTest
{
    public class �X�ȯd����������Ʒ��վ�_PC
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/about/services/customer/message-board";
        private readonly string Version = "PC";

        [Theory]
        [InlineData(BrowserType.Chrome)]
        //[InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {
                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                driver.Manage().Window.Maximize();

                int country_dropdownList_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li")).Count;  //��� "����"�U�Կ��̪�������
                
                for (int country_index = 1; country_index <= country_dropdownList_amount-1; country_index++) 
                {
                  
                    IWebElement CountryDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span")); //  "����" �U�Կ��
                    CountryDropDownList.Click(); // �i�}�������

                    string cuntry_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[1]/li/ul/li[{country_index + 1}]/span"; // ����Xpath
                    IWebElement SelectCountry = driver.FindElement(By.XPath(cuntry_xpath)); //  "����" �U�Կ��
                    SelectCountry.Click(); //�I���country_index�ӿ���

                    int branch_dropdownlist_amount = driver.FindElements(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li")).Count;  //��� "��U����"�U�Կ��̪������

                    for (int branch_index = 1; branch_index <= branch_dropdownlist_amount; branch_index++)
                    {
                        if (country_index == country_dropdownList_amount - 1 && branch_index == 14)
                        {
                            driver.Navigate().Refresh();
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000);
                        }


                        IWebElement BranchDropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span")); //  "����" �U�Կ��
                        BranchDropDownList.Click(); // �i�}������

                        string branch_xpath = $"//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]/td[2]/div/ul[2]/li/ul/li[{branch_index}]/span";
                        IWebElement SelectBranch = driver.FindElement(By.XPath(branch_xpath)); //  "����" �U�Կ��
                        SelectBranch.Click(); //�I���branch_index�ӿ���


                        string time = System.DateTime.Now.ToString("yyyyMMddHHmmss");

                        // �Ҧ����w��
                        IWebElement FullNameColumn = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[1]/td[2]/input"));
                        IWebElement IdentityCardColumn = driver.FindElement(By.XPath("//*[@id='citizenId']"));
                        IWebElement TelephoneColumn = driver.FindElement(By.XPath("//*[@id='phone']"));
                        IWebElement EMailColumn = driver.FindElement(By.XPath("//*[@id='email']"));
                        IWebElement BusinessItem_DropDownList = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/span"));
                        IWebElement CreditCardBusiness = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[1]/td[2]/ul/li/ul/li[2]/span"));
                        IWebElement MessageArea = driver.FindElement(By.XPath("//*[@id='layout_0_maincontent_2_comments']"));
                        IWebElement IHaveReadRadioButtin = driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[4]/table/tbody/tr[2]/td[2]/div[1]/a"));
                        IWebElement SubmitButton = driver.FindElement(By.XPath("//*[@id='submit']"));


                        string csvpath = $@"{UserDataList.folderpath}\testdata\UserInfo.csv";
                        using (var reader = new StreamReader(csvpath))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            var records = csv.GetRecords<UserDataList>().ToList();
                            int user_data_index = 1;
                            foreach (var user in records)
                            {
                                if (user_data_index == 2) // �� CSV �� user_data_index ����
                                {
                                    FullNameColumn.Clear();
                                    FullNameColumn.SendKeys(user.NAME); //��m�W
                                    //System.Threading.Thread.Sleep(100);

                                    IdentityCardColumn.Clear();
                                    IdentityCardColumn.SendKeys(user.ID); //�񨭤���
                                    //System.Threading.Thread.Sleep(100);

                                    TelephoneColumn.Clear();
                                    TelephoneColumn.SendKeys(user.PHONE); //��s���q��
                                   // System.Threading.Thread.Sleep(100);

                                    EMailColumn.Clear();
                                    EMailColumn.SendKeys(user.EMAIL); //�� E-MAIL
                                    //System.Threading.Thread.Sleep(100);

                                    break;
                                }
                                user_data_index++;
                            }

                            string snapshotpath = $@"{UserDataList.folderpath}\SnapshotFolder\GuestMessageBoard";
                            Tools.CreateSnapshotFolder(snapshotpath);
                            //System.Threading.Thread.Sleep(100);

                            Tools.ElementTakeScreenShot(driver.FindElement(By.XPath("//*[@id='mainform']/div[9]/div[4]/div[2]/table/tbody/tr[5]")),
                                $@"{snapshotpath}\�� {country_index} �ӿ����� {branch_index} �Ӥ���_���snapshot.png"); //�����I��


                            //System.Threading.Thread.Sleep(100);
                            BusinessItem_DropDownList.Click(); //�I�d���~�ȤU�Կ��
                            //System.Threading.Thread.Sleep(300);
                            CreditCardBusiness.Click(); //��H�Υd�~��
                            //System.Threading.Thread.Sleep(300);

                            //MessageArea.Clear();
                            //MessageArea.SendKeys(
                            //    "[Automation Test on " + browserType + " in " + Version + " ]" + "\r\n"
                            //    + "��� on �� " + country_index + " �ӿ����� " + branch_index + " �Ӥ���" + "\r\n"
                                //+ "=================================================" + "\r\n"
                                //+ "Total 21 ��, 154 ����" + "\r\n"
                                //+ "����(�a��index)(����ƶq):" + "\r\n"
                                //+ "�򶩥�(1)(1),      �O�_��(2)(35),      �s�_��(3)(32),      ��饫(4)(11),      �s�˥�(5)(3),      "
                                //+ "�s�˿�(6)(3),      �]�߿�(7)(3),        �O����(8)(14),      ���ƿ�(9)(2),        �n�뿤(10)(1)," + "\r\n"
                                //+ "���L��(11)(1),    �Ÿq��(12)(2),      �Ÿq��(13)(1),      �x�n��(14)(10),    ������(15)(14),      "
                                //+ "�̪F��(16)(2),    �y����(17)(1),      �Ὤ��(18)(1),      �O�F��(19)(1),      ���(20)(1),       ���~�a��(21)(15)."
                             //   ); // �� "�d�����e"

                           // Tools.SCrollToElement(driver, TelephoneColumn);
                            //Tools.TakeScreenShot($@"d:\�� {country_index} �ӿ����� {branch_index} �Ӥ��� snapshot_{browserType}.png", driver); // snapshot��U�e��
                           // System.Threading.Thread.Sleep(500);




                            // IHaveReadRadioButtin.Click(); // �I "�ڤw�\Ū" radio button
                            // System.Threading.Thread.Sleep(100);
                            // System.Diagnostics.Debug.WriteLine("13. �Ŀ�'�ڤw�\Ū'");

                            // SCrollToElement(driver, MessageArea);
                            // System.Threading.Thread.Sleep(3000);

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

                            //driver.FindElement(By.XPath("//*[@id='scrollUp']")).Click(); //�e���^�쳻��
                            //System.Threading.Thread.Sleep(500);




                        }
                    }
                }
                driver.Quit();
            }
        }
    }
}
    

