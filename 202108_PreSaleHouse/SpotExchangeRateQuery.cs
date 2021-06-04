using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using Xunit;
using System;
using System.IO;
using IronXL;
using System.Text.RegularExpressions;
using References;
using System.Drawing;
using OpenQA.Selenium.Remote;

namespace SpotExchangeRateQueryTest
{
    public class 即期匯率網銀買賣外幣按鈕固定置底_M 
    {
        private readonly string test_url = "https://www.esunbank.com.tw/bank/personal/deposit/rate/forex/foreign-exchange-rates?dev=mobile"; // 登入M版即期匯率網頁

        [Theory]
        [InlineData(BrowserType.Chrome)]
        [InlineData(BrowserType.Firefox)]


        public void TestCase(BrowserType browserType)
        {
            System.Threading.Thread.Sleep(100);

            using IWebDriver driver = WebDriverInfra.Create_Browser(browserType);
            {

                driver.Navigate().GoToUrl(test_url);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1000); //100秒內載完網頁內容, 否則報錯, 載完提早進下一步.
                driver.Manage().Window.Position = new Point(400, 0); //設定網頁開啟在畫面什麼位置
                driver.Manage().Window.Size = new Size(640,800); // 設定開啟的網頁大小
                string browserdriver = driver.GetType().Name.ToString(); // 偵測瀏覽器type
                string browsername = browserdriver.Remove(browserdriver.Length-6, 6); //只取browserdriver瀏覽器版本字串, 移掉字尾"Driver"字串

                int[] currencylist = new int[] { 1, 4, 7, 10, 13, 16, 19, 22, 25, 27, 29, 31, 33, 35, 37 }; // 定義有"V"的幣別XPath編號
                string[] currencynamelist = new string[] { "USD", "CNY", "HKD", "JPY", "EUR", "AUD", "CAD", "GBP", "ZAR", "NZD", "CHF", "SEK", "SGD", "MXN", "THB" }; // 幣別字串 for 截圖檔名用
                
                int k = 0; // for 截圖取 currencynamelist index用
                foreach (var currency in currencylist)
                {
                   
                    string V_Xpath = $"//*[@id='BoardRate']/tbody/tr[{currency}]/td[4]"; //定義"V" button編號規則
                    string V_dropdown_content_XPath = $"//*[@id='BoardRate']/tbody/tr[{currency+1}]";
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
                            Tools.SCrollToElement(driver, driver.FindElement(By.XPath($"//*[@id='BoardRate']/tbody/tr[{currency - 6}]/td[4]")));
                        }

                        V_button.Click(); // 點擊 "V" >>>展開"優惠匯率"選單
                        System.Threading.Thread.Sleep(500);

                        //if (browsername == "Firefox") //全螢幕截圖
                        //{
                        //    string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        //    Tools.SnapshotFullScreen($@"D:\Snapshot_Folder\SpotExchangeRateQueryTest_BuyAndSellForeignButtonCheck\{currencynamelist[k]} 展開 fullsnapshot {browserType}_{timesavepath}.png");
                        //    System.Threading.Thread.Sleep(100);
                        //}
                        //else // 網頁截圖
                        //{
                        //    string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        //    Tools.TakeScreenShot($@"D:\Snapshot_Folder\SpotExchangeRateQueryTest_BuyAndSellForeignButtonCheck\{currencynamelist[k]} 展開 snapshot {browserType}_{timesavepath}.png", driver); 
                        //    System.Threading.Thread.Sleep(100);
                        //}

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

                        //if (browsername == "FirefoxDriver") //全螢幕截圖
                        //{
                        //    string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        //    Tools.SnapshotFullScreen($@"D:\Snapshot_Folder\SpotExchangeRateQueryTest_BuyAndSellForeignButtonCheck\{currencynamelist[k]} 收合 fullsnapshot {browserType}_{timesavepath}.png");
                        //    System.Threading.Thread.Sleep(100); 
                        //}
                        //else // 網頁截圖
                        //{
                        //    string timesavepath = System.DateTime.Now.ToString("yyyyMMdd'-'HHmm"); // 偵測當下時間
                        //    Tools.TakeScreenShot($@"D:\Snapshot_Folder\SpotExchangeRateQueryTest_BuyAndSellForeignButtonCheck\{currencynamelist[k]} 收合snapshot {browserType}_{timesavepath}.png", driver);
                        //    System.Threading.Thread.Sleep(100);
                        //}

                        /// <summary>
                        /// 檢查當 "收合V選單後" 是否看的到置底網銀外幣交易按鈕
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
