using OpenQA.Selenium;
using Xunit;
using AutomatedTest.Utilities;
using Xunit.Abstractions;

namespace AutomatedTest.IntegrationTest.Personal.Deposit.Forex
{
    public class 即期匯率網銀買賣外幣按鈕固定置底_M版:IntegrationTestBase
    {
        public 即期匯率網銀買賣外幣按鈕固定置底_M版(ITestOutputHelper output, Setup testSetup) : base(output, testSetup)
        {
            testurl = "https://www.esunbank.com.tw/bank/personal/deposit/rate/forex/foreign-exchange-rates?dev=mobile"; // 登入M版即期匯率網頁
        }

        [Theory]
        [MemberData(nameof(BrowserHelper.BrowserList), MemberType = typeof(BrowserHelper))]

        public void 送出資料(string browser)
        {
            StartTestCaseForCustomizedSize(browser, "即期匯率網銀買賣外幣按鈕固定置底測試","York",  400, 0, 640, 800);
            INFO("");

            string snapshotfolderpath = $@"{System.AppDomain.CurrentDomain.BaseDirectory}SnapshotFolder\SpotExchangeRateQueryTest";
            TestBase.CreateFolder(snapshotfolderpath);

            int[] currencylist = new int[] { 1, 4, 7, 10, 13, 16, 19, 22, 25, 27, 29, 31, 33, 35, 37 }; // 定義有"V"的幣別XPath編號
            string[] currencynamelist = new string[] { "USD", "CNY", "HKD", "JPY", "EUR", "AUD", "CAD", "GBP", "ZAR", "NZD", "CHF", "SEK", "SGD", "MXN", "THB" }; // 幣別字串 for 截圖檔名用

            int k = 0; // for 截圖取 currencynamelist index用
            foreach (var currency in currencylist)
            {

                string V_Xpath = $"//*[@id='BoardRate']/tbody/tr[{currency}]/td[4]"; //定義"V" button編號規則
                string V_dropdown_content_XPath = $"//*[@id='BoardRate']/tbody/tr[{currency + 1}]";
                string HlkEBankBuy = "//*[@id='layout_m_0_content_m_3_tab_content_m_0_HlkEBankBuy']"; //定義彈出選單"網銀買外幣"位置
                string HlkEBankSell = "//*[@id='layout_m_0_content_m_3_tab_content_m_0_HlkEBankSell']"; //定義彈出選單"網銀賣外幣"位置

                int V_button_count = driver.FindElements(By.XPath(V_Xpath)).Count;
                //int HlkEBankBuy_button = driver.FindElements(By.XPath(HlkEBankBuy)).Count;
                // int HlkEBankSell_button = driver.FindElements(By.XPath(HlkEBankSell)).Count;

                IWebElement bottom_button = driver.FindElement(By.ClassName("fixed_block"));



                /// <summary>
                /// 檢查當 "進入網頁後" 是否立即看的到置底網銀外幣交易按鈕
                /// </summary>
                // Assert.Equal("bottom: -100px;", bottom_button.GetAttribute("style"));
                // Assert.Equal(1, HlkEBankBuy_button); 
                // Assert.Equal(1, HlkEBankSell_button);

                if (V_button_count >= 1) // 檢查畫面中是否存在該元素, count >=1 (存在1個element以上), else: null (沒有該element)
                {
                    IWebElement V_button = driver.FindElement(By.XPath(V_Xpath));
                    IWebElement V_dropdown_content = driver.FindElement(By.XPath(V_dropdown_content_XPath));

                    if (currency >= 7 && currency <= 25) // 避免點擊不到button, 當index介於7~25之間, 滾動畫面到target element附近
                    {
                        TestBase.SCrollToElement(driver, driver.FindElement(By.XPath($"//*[@id='BoardRate']/tbody/tr[{currency - 6}]/td[4]")));
                    }

                    V_button.Click(); // 點擊 "V" >>>展開"優惠匯率"選單
                    System.Threading.Thread.Sleep(500);

                    if (browser == "Firefox") //全螢幕截圖
                    {
                        string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        TestBase.FullScreenshot($@"{snapshotfolderpath}\{currencynamelist[k]} 展開 fullsnapshot {browser}_{time}.png");
                        System.Threading.Thread.Sleep(100);
                    }
                    else // 網頁截圖
                    {
                        string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        TestBase.PageSnapshot(driver,$@"{snapshotfolderpath}\{currencynamelist[k]} 展開 fullsnapshot {browser}_{time}.png");
                        System.Threading.Thread.Sleep(100);
                    }

                    /// <summary>
                    /// 檢查當 "點開V選單後" 是否看的到置底網銀外幣交易按鈕
                    /// <summary>
                    int HlkEBankBuy_button = driver.FindElements(By.XPath(HlkEBankBuy)).Count;
                    int HlkEBankSell_button = driver.FindElements(By.XPath(HlkEBankSell)).Count;
                    Assert.Equal("dropdown_content none show", V_dropdown_content.GetAttribute("class"));
                    Assert.Equal("bottom: 0px;", bottom_button.GetAttribute("style"));
                    Assert.Equal(1, HlkEBankBuy_button);
                    Assert.Equal(1, HlkEBankSell_button);


                    V_button.Click(); // 點擊 "V" >>>收合"優惠匯率"選單
                    System.Threading.Thread.Sleep(500);

                    if (browser == "Firefox") //全螢幕截圖
                    {
                        string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        TestBase.FullScreenshot($@"{snapshotfolderpath}\{currencynamelist[k]} 收合 fullsnapshot {browser}_{time}.png");
                        System.Threading.Thread.Sleep(100);
                    }
                    else // 網頁截圖
                    {
                        string time = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        TestBase.PageSnapshot(driver,$@"{snapshotfolderpath}\{currencynamelist[k]} 收合snapshot {browser}_{time}.png");
                        System.Threading.Thread.Sleep(100);
                    }

                    /// <summary>
                    /// 檢查當 "收合V選單後" 是否看的到置底網銀外幣交易按鈕
                    /// </summary>
                    Assert.Equal("dropdown_content none", V_dropdown_content.GetAttribute("class"));
                    Assert.Equal("bottom: -100px;", bottom_button.GetAttribute("style"));
                    Assert.Equal(1, HlkEBankBuy_button);
                    Assert.Equal(1, HlkEBankSell_button);


                    /// <summary>
                    /// 檢查置底網銀外幣交易按鈕hyperlink
                    /// </summary>
                    string HlkEBankBuy_hyperlink = driver.FindElement(By.XPath(HlkEBankBuy)).GetAttribute("href");
                    string HlkEBankSell_hyperlink = driver.FindElement(By.XPath(HlkEBankSell)).GetAttribute("href");
                    Assert.Equal("https://j3x8a.app.goo.gl/Buy-Foreign-Currency-web", HlkEBankBuy_hyperlink);
                    Assert.Equal("https://j3x8a.app.goo.gl/Sell-Foreign-Currency-web", HlkEBankSell_hyperlink);
                }
                k++;
            }
            driver.Quit();
        }
    }
}
