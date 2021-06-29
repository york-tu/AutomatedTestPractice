using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using Xunit;
using System;
using AutomatedTest.Utilities;
using System.Drawing;

namespace SpotExchangeRateQueryTest
{
    public class �Y���ײv���ȶR��~�����s�T�w�m��_M 
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/deposit/rate/forex/foreign-exchange-rates?dev=mobile"; // �n�JM���Y���ײv����

        [Theory]
        [InlineData(BrowserType.Chrome)]
        [InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);

            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {

                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100�������������e, �_�h����, ���������i�U�@�B.
                driver.Manage().Window.Position = new Point(400, 0); //�]�w�����}�Ҧb�e�������m
                driver.Manage().Window.Size = new Size(640,800); // �]�w�}�Ҫ������j�p

                string browsername = ((RemoteWebDriver)driver).Capabilities.GetCapability("browserName").ToString(); // �����s��������
                //  string browsername = ((RemoteWebDriver)driver).Capabilities.GetCapability("browserVersion").ToString(); // �����s��������

                int[] currencylist = new int[] { 1, 4, 7, 10, 13, 16, 19, 22, 25, 27, 29, 31, 33, 35, 37 }; // �w�q��"V"�����OXPath�s��
                string[] currencynamelist = new string[] { "USD", "CNY", "HKD", "JPY", "EUR", "AUD", "CAD", "GBP", "ZAR", "NZD", "CHF", "SEK", "SGD", "MXN", "THB" }; // ���O�r�� for �I���ɦW��
                
                int k = 0; // for �I�Ϩ� currencynamelist index��
                foreach (var currency in currencylist)
                {
                   
                    string V_Xpath = $"//*[@id='BoardRate']/tbody/tr[{currency}]/td[4]"; //�w�q"V" button�s���W�h
                    string V_dropdown_content_XPath = $"//*[@id='BoardRate']/tbody/tr[{currency+1}]";
                    string HlkEBankBuy = "//*[@id='layout_m_0_content_m_3_tab_content_m_0_HlkEBankBuy']"; //�w�q�u�X���"���ȶR�~��"��m
                    string HlkEBankSell = "//*[@id='layout_m_0_content_m_3_tab_content_m_0_HlkEBankSell']"; //�w�q�u�X���"���Ƚ�~��"��m
                    
                    int V_button_count = driver.FindElements(By.XPath(V_Xpath)).Count;
                    //int HlkEBankBuy_button = driver.FindElements(By.XPath(HlkEBankBuy)).Count;
                    // int HlkEBankSell_button = driver.FindElements(By.XPath(HlkEBankSell)).Count;

                    IWebElement bottom_button = driver.FindElement(By.ClassName("fixed_block"));

                    string snapshotfolderpath = $@"{UserDataList.folderpath}\SnapshotFolder\SpotExchangeRateQueryTest";
                    Tools.CreateSnapshotFolder(snapshotfolderpath);

                    /// <summary>
                    /// �ˬd�� "�i�J������" �O�_�ߧY�ݪ���m�����ȥ~��������s
                    /// </summary>
                    // Assert.Equal("bottom: -100px;", bottom_button.GetAttribute("style"));
                    // Assert.Equal(1, HlkEBankBuy_button); 
                    // Assert.Equal(1, HlkEBankSell_button);

                    if (V_button_count >= 1) // �ˬd�e�����O�_�s�b�Ӥ���, count >=1 (�s�b1��element�H�W), else: null (�S����element)
                    {
                        IWebElement V_button = driver.FindElement(By.XPath(V_Xpath));
                        IWebElement V_dropdown_content = driver.FindElement(By.XPath(V_dropdown_content_XPath));

                        if (currency >= 7 && currency <= 25) // �קK�I������button, ��index����7~25����, �u�ʵe����target element����
                        {
                            Tools.SCrollToElement(driver, driver.FindElement(By.XPath($"//*[@id='BoardRate']/tbody/tr[{currency - 6}]/td[4]")));
                        }

                        V_button.Click(); // �I�� "V" >>>�i�}"�u�f�ײv"���
                        System.Threading.Thread.Sleep(500);

                        if (browsername == "Firefox") //���ù��I��
                        {
                            string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // ������U�ɶ�
                            Tools.SnapshotFullScreen($@"{snapshotfolderpath}\{currencynamelist[k]} �i�} fullsnapshot {browserType}_{time}.png");
                            System.Threading.Thread.Sleep(100);
                        }
                        else // �����I��
                        {
                            string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // ������U�ɶ�
                            Tools.TakeScreenShot($@"{snapshotfolderpath}\{currencynamelist[k]} �i�} fullsnapshot {browserType}_{time}.png", driver);
                            System.Threading.Thread.Sleep(100);
                        }

                        /// <summary>
                        /// �ˬd�� "�I�}V����" �O�_�ݪ���m�����ȥ~��������s
                        /// <summary>
                        int HlkEBankBuy_button = driver.FindElements(By.XPath(HlkEBankBuy)).Count;
                        int HlkEBankSell_button = driver.FindElements(By.XPath(HlkEBankSell)).Count;
                        Assert.Equal("dropdown_content none show", V_dropdown_content.GetAttribute("class"));
                        Assert.Equal("bottom: 0px;", bottom_button.GetAttribute("style"));
                        Assert.Equal(1, HlkEBankBuy_button);
                        Assert.Equal(1, HlkEBankSell_button);


                        V_button.Click(); // �I�� "V" >>>���X"�u�f�ײv"���
                        System.Threading.Thread.Sleep(500);

                        if (browsername == "Firefox") //���ù��I��
                        {
                            string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // ������U�ɶ�
                            Tools.SnapshotFullScreen($@"{snapshotfolderpath}\{currencynamelist[k]} ���X fullsnapshot {browserType}_{time}.png");
                            System.Threading.Thread.Sleep(100);
                        }
                        else // �����I��
                        {
                            string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // ������U�ɶ�
                            Tools.TakeScreenShot($@"{snapshotfolderpath}\{currencynamelist[k]} ���Xsnapshot {browserType}_{time}.png", driver);
                            System.Threading.Thread.Sleep(100);
                        }

                        /// <summary>
                        /// �ˬd�� "���XV����" �O�_�ݪ���m�����ȥ~��������s
                        /// </summary>
                        Assert.Equal("dropdown_content none", V_dropdown_content.GetAttribute("class"));
                        Assert.Equal("bottom: -100px;", bottom_button.GetAttribute("style"));
                        Assert.Equal(1, HlkEBankBuy_button);
                        Assert.Equal(1, HlkEBankSell_button);
                    }
                    k++;
                }
            }
            driver.Quit();
        }
    }
}
