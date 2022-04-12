using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using IronOcr;
using AventStack.ExtentReports;
using S22.Imap;
using AutomatedTest.Utilities;
using System.Collections.Generic;
using System.Threading;
using System.Text;

namespace AutomatedTest.Utilities
{
    public class TestBase
    {
        protected IWebDriver driver;
        private int _interval = 500;
        private int _timeout = 60;
        protected ExtentReports extentReports;
        protected ExtentTest ExtentTObj;
        private static Image img;
        private static Bitmap bimg;
        private static string Security_Code;

        public TestBase(string browserArg, ExtentTest extentTObj)
        {
            BrowserHelper browser = new BrowserHelper(browserArg);
            //this.driver = browser.driver;
            this.ExtentTObj = extentTObj;
        }
        public TestBase(string browserArg, int timeout, int interval)
        {
            BrowserHelper browser = new BrowserHelper(browserArg);
            //this.driver = browser.driver;
            _timeout = timeout;
            _interval = interval;
        }

        #region 截圖附上report
        /// <summary>
        /// 網頁截圖 (for attach to report)
        /// </summary>
        /// <param name="driver">指定IWebDriver </param>
        /// <returns>ImageString</returns>
        public static string PageSnapshotToReport(IWebDriver driver) 
        {
            var img_HTML = string.Empty;
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            var screenshotByArray = ss.AsByteArray;
            var temp_inBase64 = Convert.ToBase64String(screenshotByArray);
            img_HTML = "<img style=\"width: 900px;\" src=\"data:image/jpg;base64, " + temp_inBase64 + "\"/>";
            return img_HTML;
        }
        /// <summary>
        /// 元件截圖 (for attach to report)
        /// </summary>
        /// <param name=" webElement">指定IWebElement</param>
        /// <returns>ImageString</returns>
        public static string ElementSnapShotToReport(IWebElement webElement) 
        {
            var img_HTML = string.Empty;
            Screenshot elementScreenshot = (webElement as ITakesScreenshot).GetScreenshot();
            var elementscreenshotByArray = elementScreenshot.AsByteArray;
            var temp_inBase64 = Convert.ToBase64String(elementscreenshotByArray);
            img_HTML = "<img style=\"width: 720px;\" src=\"data:image/png;base64, " + temp_inBase64 + "\"/>";
            return img_HTML;
        }
        /// <summary>
        /// 全螢幕截圖 (for attach to report)
        /// </summary>
        /// <param name="imagefilepath">圖片路徑</param>
        /// <returns>ImageString</returns>
        public static string FullScreenshotToReport(string imagefilepath)
        {
            var img_HTML = string.Empty;
            Bitmap bmp = new Bitmap(imagefilepath);
            MemoryStream ms = new MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            byte[] arr = new byte[ms.Length];
            ms.Position = 0;
            ms.Read(arr, 0, (int)ms.Length);
            ms.Close();
            var temp_inBase64 = Convert.ToBase64String(arr);
            img_HTML = "<img style=\"width: 720px;\" src=\"data:image/png;base64, " + temp_inBase64 + "\"/>";
            return img_HTML;
        }
        #endregion

        #region 圖片辨識
        /// <summary>
        /// 讀取驗證碼圖片
        /// </summary>
        /// <param name="sourcepath">圖片來源路徑</param>
        /// <returns>image</returns>
        private static Image Getimage(string sourcepath) 
        {
            Bitmap img = new Bitmap(sourcepath);
            return img;
        }
        /// <summary>
        /// 執行圖片色階處理
        /// </summary>
        /// <returns>解析出文字</returns>
        public static string executing() 
        {
            Image lookimg;

            //轉灰階
            convert2GrayScale();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());

            //去雜線
            RemoteNoiseLineByPixels();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());

            //去外框
            ClearPictureBorder(2);

            //去雜點
            RemoteNoisePointByPixels();
            lookimg = Image.FromHbitmap(bimg.GetHbitmap());

            //補點
            AddNoisePointByPixels();
            img = Image.FromHbitmap(bimg.GetHbitmap());
            parseCaptchaStr(img);

            return Security_Code;
        }
        /// <summary>
        /// Tesseract圖片驗證碼辨識
        /// </summary>
        /// <param name="filepath">檔案路徑</param>
        /// <param name="imagescale">圖像比例</param>
        /// <returns>字串</returns>
        public static string TesseractOCRIdentify(string filepath, double imagescale) 
        {
            img = Getimage(filepath);
            img = new Bitmap(img, (int)(img.Width * imagescale), (int)(img.Height * imagescale));
            bimg = new Bitmap(img);
            string words = executing();
            return words;
        }
        private static void convert2GrayScale() // 轉灰階
        {
            for (int i = 0; i < bimg.Width; i++)
            {
                for (int j = 0; j < bimg.Height; j++)
                {
                    Color pixelColor = bimg.GetPixel(i, j);
                    byte r = pixelColor.R;
                    byte g = pixelColor.G;
                    byte b = pixelColor.B;

                    byte gray = (byte)(0.299 * (float)r + 0.587 * (float)g + 0.114 * (float)b);
                    r = g = b = gray;
                    pixelColor = Color.FromArgb(r, g, b);

                    bimg.SetPixel(i, j, pixelColor);
                }
            }
        }
        private static void ClearPictureBorder(int pBorderWidth) //去邊框
        {
            for (int i = 0; i < bimg.Height; i++)
            {
                for (int j = 0; j < bimg.Width; j++)
                {
                    if (i < pBorderWidth || j < pBorderWidth || j > bimg.Width - 1 - pBorderWidth || i > bimg.Height - 1 - pBorderWidth)
                        bimg.SetPixel(j, i, Color.FromArgb(255, 255, 255));
                }
            }
        }
        private class NoisePoint
        {
            public int X { get; set; }
            public int Y { get; set; }
        }
        private static void RemoteNoisePointByPixels() //去雜點
        {
            List<NoisePoint> points = new List<NoisePoint>();

            for (int k = 0; k < 5; k++)
            {
                for (int i = 0; i < bimg.Height; i++)
                    for (int j = 0; j < bimg.Width; j++)
                    {
                        int flag = 0;
                        int garyVal = 255;
                        // 檢查上相鄰像素
                        if (i - 1 > 0 && bimg.GetPixel(j, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && bimg.GetPixel(j, i + 1).R != garyVal) flag++;
                        if (j - 1 > 0 && bimg.GetPixel(j - 1, i).R != garyVal) flag++;
                        if (j + 1 < bimg.Width && bimg.GetPixel(j + 1, i).R != garyVal) flag++;
                        if (i - 1 > 0 && j - 1 > 0 && bimg.GetPixel(j - 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j - 1 > 0 && bimg.GetPixel(j - 1, i + 1).R != garyVal) flag++;
                        if (i - 1 > 0 && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i + 1).R != garyVal) flag++;

                        if (flag < 3)
                            points.Add(new NoisePoint() { X = j, Y = i });
                    }
                foreach (NoisePoint point in points)
                    bimg.SetPixel(point.X, point.Y, Color.FromArgb(255, 255, 255));
            }
        }
        private static void RemoteNoiseLineByPixels() //去噪音線
        {
            for (int i = 0; i < bimg.Height; i++)
                for (int j = 0; j < bimg.Width; j++)
                {
                    int grayValue = bimg.GetPixel(j, i).R;
                    if (grayValue <= 255 && grayValue >= 160)
                        bimg.SetPixel(j, i, Color.FromArgb(255, 255, 255));
                }
        }
        private static void AddNoisePointByPixels()  //補點
        {
            List<NoisePoint> points = new List<NoisePoint>();

            for (int k = 0; k < 1; k++)
            {
                for (int i = 0; i < bimg.Height; i++)
                    for (int j = 0; j < bimg.Width; j++)
                    {
                        int flag = 0;
                        int garyVal = 255;
                        // 檢查上相鄰像素
                        if (i - 1 > 0 && bimg.GetPixel(j, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && bimg.GetPixel(j, i + 1).R != garyVal) flag++;
                        if (j - 1 > 0 && bimg.GetPixel(j - 1, i).R != garyVal) flag++;
                        if (j + 1 < bimg.Width && bimg.GetPixel(j + 1, i).R != garyVal) flag++;
                        if (i - 1 > 0 && j - 1 > 0 && bimg.GetPixel(j - 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j - 1 > 0 && bimg.GetPixel(j - 1, i + 1).R != garyVal) flag++;
                        if (i - 1 > 0 && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i - 1).R != garyVal) flag++;
                        if (i + 1 < bimg.Height && j + 1 < bimg.Width && bimg.GetPixel(j + 1, i + 1).R != garyVal) flag++;

                        if (flag >= 7)
                            points.Add(new NoisePoint() { X = j, Y = i });
                    }
                foreach (NoisePoint point in points)
                    bimg.SetPixel(point.X, point.Y, Color.FromArgb(0, 0, 0));
            }
        }
        private static void parseCaptchaStr(Image image) // 讀出字串
        {
            Ocr ocr = new Ocr();
            Bitmap BmpSource = new Bitmap(image);
            ocr.DoOCRMultiThred(BmpSource, "eng");
            Security_Code = ocr.GetSecurity_Code();
        }
        /// <summary>
        /// 百度_圖片辨識
        /// </summary>
        /// <param name="imagepath">圖片路徑</param>
        /// <returns>字串</returns>
        public static string BaiduOCR(string imagepath)  
        {
            var API_KEY = "cohIahxAt7HveHLYSHYK6G5N"; // "FGPi0QpCbZxZxBaN6dvqt87X";
            var SECRET_KEY = "e8SAsDIWSK9NPUKviYiPQNlfaVDXQSY5"; // "HunNq6XsLjF3a7aCAuirVaVQO7CKBuwW";

            var client = new Baidu.Aip.Ocr.Ocr(API_KEY, SECRET_KEY);
            client.Timeout = 60000;  // 修改超時時間
            var image = File.ReadAllBytes(imagepath);

            // 呼叫通用文字識別, 圖片引數為本地圖片，可能會丟擲網路等異常，請使用try/catch捕獲
            //使用者向服務請求識別某張圖中的所有文字
            //var result = client.GeneralBasic(image);        //本地圖圖片
            //var result = client.GeneralBasicUrl(url);     //網路圖片

            //var result = client.General(image);           //本地圖片：通用文字識別（含位置資訊版）
            //var result = client.GeneralUrl(url);          //網路圖片：通用文字識別（含位置資訊版）

            //var result = client.GeneralEnhanced(image);   //本地圖片：呼叫通用文字識別（含生僻字版）
            //var result = client.GeneralEnhancedUrl(url);  //網路圖片：呼叫通用文字識別（含生僻字版）

            //var result = client.WebImage(image);          //本地圖片:使用者向服務請求識別一些背景複雜，特殊字型的文字。
            //var result = client.WebImageUrl(url);         //網路圖片:使用者向服務請求識別一些背景複雜，特殊字型的文字。

            string JsonString = client.Accurate(image).ToString();          //本地圖片：相對於通用文字識別該產品精度更高，但是識別耗時會稍長。
            string cut = JsonString.Substring(45, 10); // 擷取jason 文字串中 解析出的圖片文字 這段
            string CutResult = Regex.Replace(cut, "[^0-9A-Za-z]", ""); //去除掉符號空白...只留下字母&数字
            return CutResult;
        }
        /// <summary>
        /// Iron_圖片辨識
        /// </summary>
        /// <param name="imagepath">圖片路徑</param>
        /// <returns>字串</returns>
        public static string IronOCR(string imagepath)  
        {
            var Ocr = new IronTesseract();
            using (var Input = new OcrInput(imagepath))
            {
                Input.Deskew();  // use if image not straight
                Input.DeNoise(); // use if image contains digital noise
                var Result = Ocr.Read(Input);
                return Result.Text;
            }
        }
        #endregion


        #region 工具區
        /// <summary>
        /// 畫面定位 (selenium)
        /// </summary>
        /// <param name="driver">IWebDriver</param>
        /// <param name="element">指定元件 IWebElement</param>
        public static void Find_Element(IWebDriver driver, IWebElement element) 
        {
            Actions action = new Actions(driver);
            action.MoveToElement(element).Perform();
        }
        /// <summary>
        ///  畫面滾到目標位置
        /// </summary>
        /// <param name="driver">IWebDriver</param>
        /// <param name="element">指定元件 IWebElement</param>
        public static void SCrollToElement(IWebDriver driver, IWebElement element) 
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].scrollIntoView(true);", element); // Viewport對頂部對齊
            Actions act = new Actions(driver);
            act.MoveToElement(element).Perform();
        }
        /// <summary>
        /// 畫面向上/向下滾動 n pixel
        /// </summary>
        /// <param name="driver">IWebDriver</param>
        /// <param name="scroll_pixels">滾動pixel單位</param>
        public static void ScrollPageUpOrDown(IWebDriver driver, int scrollPixels) 
        {
            var js = (IJavaScriptExecutor)driver;
            js.ExecuteScript($"window.scrollTo(0,{scrollPixels})");
        }
        /// <summary>
        /// 建立資料夾 
        /// </summary>
        /// <param name="snapshotpath">資料夾路徑</param>
        public static void CreateFolder(string folderpath) 
        {
            try
            {
                if (Directory.Exists(folderpath))
                {
                    return;
                }
                else
                {
                    DirectoryInfo directiry = Directory.CreateDirectory(folderpath);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// 清空指定資料夾
        /// </summary>
        /// <param name="folderpath">資料夾路徑</param>
        public static void CleanUPFolder(string folderpath)
        {
            string[] need_to_clean_folder = Directory.GetFileSystemEntries(folderpath);
            foreach (var file in need_to_clean_folder)
            {
                File.Delete(file);
            }
        }
        /// <summary>
        /// 截圖當下畫面 (selenium)
        /// </summary>
        /// <param name="savepath">存檔路徑</param>
        /// <param name="driver">指定IwebDirver</param>
        public static void PageSnapshot(IWebDriver driver,string savepath)
        {
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(savepath, ScreenshotImageFormat.Png);
        }
        /// <summary>
        /// 元件截圖 (selenium)
        /// </summary>
        /// <param name="webElement">指定元件 IWebElement</param>
        /// <param name="savepath">存檔路徑</param>
        public static void ElementSnapshotshot(IWebElement element, string savepath)  
        {
            var elementScreenshot = (element as ITakesScreenshot).GetScreenshot();
            elementScreenshot.SaveAsFile(savepath);
        }
        /// <summary>
        /// 全螢幕截圖 (C#)
        /// </summary>
        /// <param name="savepath">存檔路徑</param>
        public static void FullScreenshot(string savepath) 
        {
            Bitmap myimage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myimage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            myimage.Save(savepath);
        }
        /// <summary>
        /// 亂數產生新式身分證字號
        /// </summary>
        /// <param name="sex">性別男(true or false)</param>
        /// <param name="city">城市索引(0,1,...)</param>
        /// <returns>身分證字號</returns>
        public static string CreateIDNumber(bool sex, int city)
        {
            //身分證開頭英文 (參考網址: https://ithelp.ithome.com.tw/articles/10202484)
            /*              
             (1)英文代號以下表轉換成數字 
       　　　A=10 台北市 city索引值(0)　　J=18 新竹縣　  city索引值(9)　　 S=26 高雄縣 city索引值
       　　　B=11 台中市 city索引值(1)　　K=19 苗栗縣　  city索引值(10)　  T=27 屏東縣 city索引值(16) 
       　　　C=12 基隆市 city索引值(2)　　L=20 台中縣　  city索引值     　     U=28 花蓮縣 city索引值(17) 
       　　　D=13 台南市 city索引值(3)　　M=21 南投縣　city索引值(11)       V=29 台東縣 city索引值(18) 
       　　　E=14 高雄市 city索引值(4)　　N=22 彰化縣　  city索引值(12)　  W=32 金門縣 city索引值(19) 
       　　　F=15 台北縣 city索引值(5)　　O=35 新竹市　 city索引值(13)　    X=30 澎湖縣 city索引值(20) 
       　　　G=16 宜蘭縣 city索引值(6)　　P=23 雲林縣　 city索引值(14)　    Y=31 陽明山 city索引值 
       　　　H=17 桃園縣 city索引值(7)　　Q=24 嘉義縣 　city索引值(15)　   Z=33 連江縣 city索引值(21)
       　　     I=34 嘉義市 city索引值(8)　　 R=25 台南縣 　city索引值
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

        }
        /// <summary>
        /// 亂數產生行動電話號碼
        /// </summary>
        /// <returns>行動電話號碼</returns>
        public static string CreateCellPhoneNumber() 
        {
            Random phone = new Random();
            string cellphonenumber = $"09{(phone.Next(100000000) + 100000000).ToString().Substring(1)}";
            return cellphonenumber;
        }
        /// <summary>
        /// 產生指定長度 " 頭(1位大or小寫英文) + 大or小寫英文&數字 " 隨機組合字串
        /// </summary>
        /// <param name="length">指定字串長度</param>
        /// <returns>英數字組合字串</returns>
        public static string CreateRandomString(int length)  
        {
            Random r = new Random();

            string code = "";
            switch (r.Next(0, 2))
            {
                case 0: code += (char)r.Next(65, 91); break;
                case 1: code += (char)r.Next(97, 123); break; //char代碼表:  https://www.cnblogs.com/tian_z/archive/2010/08/06/1793736.html
            }

            for (int i = 0; i < length; ++i)
                switch (r.Next(0, 3))
                {
                    case 0: code += r.Next(0, 10); break;
                    case 1: code += (char)r.Next(65, 91); break;
                    case 2: code += (char)r.Next(97, 123); break;
                }

            return code;
        }
        /// <summary>
        /// 產生指定長度隨機數字組合
        /// </summary>
        /// <param name="length">指定字串長度</param>
        /// <returns>數字組合字串</returns>
        public static string CreateRandomNumber(int length) 
        {
            Random r = new Random();

            string number = "";

            for (int i = 0; i < length; ++i)
            {
                number += r.Next(0, 10);
            }

            return number;
        }
        /// <summary>
        /// 亂數產生公司統一編號
        /// </summary>
        /// <returns>公司統一編號</returns>
        public static string CreateUniformNumber()  
        {
            int[] vat = new int[8];   // 用來存統一編號的 8 個數字
            int[] WeightedCount = new int[] { 1, 2, 1, 2, 1, 2, 4, 1 };   // 用來存統一編號 8 個數字依序要乘的權數
            string uniformNumber = "";

            Random ran = new Random();
            for (int i = 0; i < 7; i++)
            {
                vat[i] = ran.Next(0, 10);
                uniformNumber = uniformNumber + vat[i].ToString();
            }
            vat[7] = FindLastNumberInVATnumber();
            uniformNumber = uniformNumber + vat[7].ToString();
            return uniformNumber;

            int FindLastNumberInVATnumber()
            {
                int Amount = 0;
                for (int i = 0; i < 7; i++)
                {
                    Amount += (vat[i] * WeightedCount[i]) / 10 + (vat[i] * WeightedCount[i]) % 10;
                }
                return (10 - (Amount % 10)) % 10;
            }
        }
        /// <summary>
        /// 檢核是否為現行本國人身分證字號
        /// </summary>
        /// <param name="id">身分證字號</param>
        /// <returns>true or false</returns>
        public static bool CheckResidentID(string id)  
        {
            //除了檢查碼外每個數字的存放空間 
            int[] seed = new int[10];

            //建立字母陣列(A~Z)
            //A=10 B=11 C=12 D=13 E=14 F=15 G=16 H=17 J=18 K=19 L=20 M=21 N=22
            //P=23 Q=24 R=25 S=26 T=27 U=28 V=29 X=30 Y=31 W=32  Z=33 I=34 O=35            
            string[] charMapping = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "X", "Y", "W", "Z", "I", "O" };
            string target = id.Substring(0, 1); //取第一個英文數字
            for (int index = 0; index < charMapping.Length; index++)
            {
                if (charMapping[index] == target)
                {
                    index += 10;
                    //10進制的高位元放入存放空間   (權重*1)
                    seed[0] = index / 10;

                    //10進制的低位元*9後放入存放空間 (權重*9)
                    seed[1] = (index % 10) * 9;

                    break;
                }
            }
            for (int index = 2; index < 10; index++) //(權重*8~1)
            {   //將剩餘數字乘上權數後放入存放空間                
                seed[index] = Convert.ToInt32(id.Substring(index - 1, 1)) * (10 - index);
            }
            //檢查是否符合檢查規則，10減存放空間所有數字和除以10的餘數的個位數字是否等於檢查碼            
            //(10 - ((seed[0] + .... + seed[9]) % 10)) % 10 == 身分證字號的最後一碼   
            if ((10 - (seed.Sum() % 10)) % 10 != Convert.ToInt32(id.Substring(9, 1)))
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 檢核是否為中華民國外僑及大陸人士在台居留證證號(舊式+新式)
        /// </summary>
        /// <param name="id">身分證字號</param>
        /// <returns>true or false</returns>
        public static bool CheckForeignerID(string id)  
        {
            id = id.ToUpper();
            Regex regex = new Regex(@"^([A-Z])(A|B|C|D|8|9)(\d{8})$");
            Match match = regex.Match(id);


            if ("ABCD".IndexOf(match.Groups[2].Value) >= 0)
            {
                //舊式統號檢查
                return CheckOldResidentID(match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value);
            }
            else
            {
                //新式統號檢查 (2021/01/02 正式生效)
                return CheckNewResidentID(match.Groups[1].Value, match.Groups[2].Value + match.Groups[3].Value);
            }

            /// <summary>
            /// 舊式統號檢核
            /// </summary>
            /// <param name="firstLetter">第1碼英文字母(區域碼)</param>
            /// <param name="secondLetter">第2碼英文字母(性別碼)</param>
            /// <param name="num">第3~9流水號 + 第10碼檢查碼</param>
            /// <returns></returns>
            bool CheckOldResidentID(string firstLetter, string secondLetter, string num)
            {
                ///建立字母對應表(A~Z)
                ///A=10 B=11 C=12 D=13 E=14 F=15 G=16 H=17 J=18 K=19 L=20 M=21 N=22
                ///P=23 Q=24 R=25 S=26 T=27 U=28 V=29 X=30 Y=31 W=32  Z=33 I=34 O=35 
                string alphabet = "ABCDEFGHJKLMNPQRSTUVXYWZIO";
                string transferIdNo =
                    $"{alphabet.IndexOf(firstLetter) + 10}" +
                    $"{(alphabet.IndexOf(secondLetter) + 10) % 10}" +
                    $"{num}";
                int[] idNoArray = transferIdNo.ToCharArray()
                                              .Select(c => Convert.ToInt32(c.ToString()))
                                              .ToArray();

                int sum = idNoArray[0];
                int[] weight = new int[] { 9, 8, 7, 6, 5, 4, 3, 2, 1, 1 };
                for (int i = 0; i < weight.Length; i++)
                {
                    sum += weight[i] * idNoArray[i + 1];
                }
                return (sum % 10 == 0);
            }

            /// <summary>
            /// 新式統號檢核
            /// </summary>
            /// <param name="firstLetter">第1碼英文字母(區域碼)</param>
            /// <param name="num">第2碼(性別碼) + 第3~9流水號 + 第10碼檢查碼</param>
            /// <returns></returns>
            bool CheckNewResidentID(string firstLetter, string num)
            {
                ///建立字母對應表(A~Z)
                ///A=10 B=11 C=12 D=13 E=14 F=15 G=16 H=17 J=18 K=19 L=20 M=21 N=22
                ///P=23 Q=24 R=25 S=26 T=27 U=28 V=29 X=30 Y=31 W=32  Z=33 I=34 O=35 
                string alphabet = "ABCDEFGHJKLMNPQRSTUVXYWZIO";
                string transferIdNo = $"{(alphabet.IndexOf(firstLetter) + 10)}" +
                                      $"{num}";
                int[] idNoArray = transferIdNo.ToCharArray()
                                              .Select(c => Convert.ToInt32(c.ToString()))
                                              .ToArray();

                int sum = idNoArray[0];
                int[] weight = new int[] { 9, 8, 7, 6, 5, 4, 3, 2, 1, 1 };
                for (int i = 0; i < weight.Length; i++)
                {
                    sum += (weight[i] * idNoArray[i + 1]) % 10;
                }
                return (sum % 10 == 0);
            }


        }
        /// <summary>
        /// 檢核是否為符合規則的公司統一編號
        /// </summary>
        /// <param name="inputUniformNumber">公司統一編號</param>
        /// <returns>true or false</returns>
        public static bool CheckUniformNumber(string inputUniformNumber) 
        {
            if (inputUniformNumber == null)
            {
                return false;
            }
            Regex regex = new Regex(@"^\d{8}$");
            Match match = regex.Match(inputUniformNumber);
            if (!match.Success)
            {
                return false;
            }
            int[] idNoArray = inputUniformNumber.ToCharArray().Select(c => Convert.ToInt32(c.ToString())).ToArray();
            int[] weight = new int[] { 1, 2, 1, 2, 1, 2, 4, 1 };

            int subSum;     //小和
            int sum = 0;    //總和
            int sumFor7 = 1;
            for (int i = 0; i < idNoArray.Length; i++)
            {
                subSum = idNoArray[i] * weight[i];
                sum += (subSum / 10)   //商數
                     + (subSum % 10);  //餘數                
            }
            if (idNoArray[6] == 7)
            {
                //若第7碼=7，則會出現兩種數值都算對，因此要特別處理。
                sumFor7 = sum + 1;
            }
            return (sum % 10 == 0) || (sumFor7 % 10 == 0);
        }
        /// <summary>
        /// 檢查輸入字串是否含有數字
        /// </summary>
        /// <param name="inputString"></param>
        /// <returns>true or false</returns>
        public static bool CheckStringContainNumber(string inputString)  
        {
            if (Regex.Replace(inputString, @"[^0-9]+", "") != "")
            {
                return true;
            }
            else
                return false;
        }
        /// <summary>
        /// 強制砍process
        /// </summary>
        /// <param name="processname">process name  (e.g., chromedriver.exe. geckodriver.exe)</param>
        public static void KillProcess(string processname)
        {
            System.Diagnostics.ProcessStartInfo p;
            p = new System.Diagnostics.ProcessStartInfo("cmd.exe", "/C " + "taskkill /f /im " + processname);
            System.Diagnostics.Process proc = new System.Diagnostics.Process();
            proc.StartInfo = p;
            proc.Start();
            proc.WaitForExit();
            proc.Close();
        }
        /// <summary>
        /// 讀取Gmail最新信件內容
        /// </summary>
        /// <param name="UserAccount">使用者帳號</param>
        /// <param name="Password">使用者密碼</param>
        /// <returns>output 信件內容 to report</returns>
        public static string ReadGmailRecentMail(string UserAccount, string Password)  
        {
            string hostname = "imap.gmail.com";
            string username = UserAccount;
            string password = Password;
            using (ImapClient client = new ImapClient(hostname, 993, username, password, AuthMethod.Login, true))
            {
                var uids = client.Search(SearchCondition.All()); // 指定信件匣 "所有" 信件
                var last_message = client.GetMessages(uids).Last(); // 擷取信件匣內"最新"mail

                string from = last_message.From.Address.ToString().Trim();
                string from_displayname = last_message.From.DisplayName.ToString().Trim();
                string to = last_message.To.ToString().Trim();
                string date = last_message.Date().ToString().Trim();
                string subject = last_message.Subject.ToString().Trim();
                string body = last_message.Body.ToString().Trim();
                string bodyHTNL = last_message.IsBodyHtml.ToString().Trim();
                string attachment = last_message.Attachments.ToString().Trim();

                return $" 寄件人: {from_displayname} <{from}> \n 收件人: {to} \n 收件時間: {date} \n 主旨: {subject} \n 內文: \n\n {body}"; // report出信件內容
            }

        }
        #endregion
    }

}