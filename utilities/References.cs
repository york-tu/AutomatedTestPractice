using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using System.Drawing;

namespace References
{
    public class Tools
    {

        public static void Find_Element(IWebDriver driver, IWebElement element) // 畫面定位 (selenium)
        {
            Actions action = new Actions(driver);
            action.MoveToElement(element).Perform();

        }
        public static void TakeScreenShot(string savepath, IWebDriver driver) // 截圖當下畫面 (selenium)
        {
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(savepath, ScreenshotImageFormat.Png);
        }

        public static void ElementTakeScreenShot(IWebElement webElement, string savepath) // Only截元件圖 (selenium)
        {
            var elementScreenshot = (webElement as ITakesScreenshot).GetScreenshot();
            elementScreenshot.SaveAsFile(savepath);
        } 

        public static void SnapshotFullScreen (string savepath) //全螢幕截圖 (C#)
        {
            Bitmap myimage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myimage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0),new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            myimage.Save(savepath);
        }
        public static void SCrollToElement(IWebDriver driver, IWebElement element) // 畫面滾到target位置
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", element); // Viewport對頂部對齊
            Actions act = new Actions(driver);
            act.MoveToElement(element).Perform();

        }

        public static string CreateIDNumber(bool sex, int city)
        {
            //身分證開頭英文 (參考網址: https://ithelp.ithome.com.tw/articles/10202484)
            /*              
             (1)英文代號以下表轉換成數字 
       　　　A=10 台北市 city索引值(0)　　J=18 新竹縣　city索引值(9)　　 S=26 高雄縣 city索引值
       　　　B=11 台中市 city索引值(1)　　K=19 苗栗縣　city索引值(10)　　T=27 屏東縣 city索引值(16) 
       　　　C=12 基隆市 city索引值(2)　　L=20 台中縣　city索引值     　 U=28 花蓮縣 city索引值(17) 
       　　　D=13 台南市 city索引值(3)　　M=21 南投縣　city索引值(11)　　V=29 台東縣 city索引值(18) 
       　　　E=14 高雄市 city索引值(4)　　N=22 彰化縣　city索引值(12)　  W=32 金門縣 city索引值(19) 
       　　　F=15 台北縣 city索引值(5)　　O=35 新竹市　city索引值(13)　　X=30 澎湖縣 city索引值(20) 
       　　　G=16 宜蘭縣 city索引值(6)　　P=23 雲林縣　city索引值(14)　　Y=31 陽明山 city索引值 
       　　　H=17 桃園縣 city索引值(7)　　Q=24 嘉義縣　city索引值(15)　  Z=33 連江縣 city索引值(21)
       　　  I=34 嘉義市 city索引值(8)　　R=25 台南縣　city索引值
         */

            string[] county_E = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
                      "M", "N", "O", "P", "Q", "T", "U", "V", "W", "X", "Z" };
            //對應數字 (索引值)
            int[] county_i = { 10, 11, 12, 13, 14, 15, 16, 17, 34, 18, 19, 21, 22, 35,
                       23, 24, 27, 28, 29, 32, 30, 33 };
            Random r = new Random();
            string id = county_E[city];
            int c_i = county_i[city];
            string s = "2";
            if (sex) s = "1";
            int rand_i = r.Next(0, 10000000);
            //計算
            int check = (c_i / 10) + 9 * (c_i - (c_i / 10) * 10) + Convert.ToInt32(s) * 8;
            for (int i = 7; i >= 1; i--)
            {
                check += ((rand_i / (int)Math.Pow(10, i - 1)) % 10) % 10 * i;
            }
            check = (10 - (check % 10)) % 10;
            //計算審核碼
            id += s + rand_i.ToString().PadLeft(7, '0') + check.ToString();
            return id;

        } // 亂數產生有效身分證字號
        public static string CreateCellPhoneNumber() // 亂數產生有效行動電話號碼
        {
            Random phone = new Random();
            string cellphonenumber = $"09{(phone.Next(100000000) + 100000000).ToString().Substring(1)}";
            return cellphonenumber;
        }
    }
}
