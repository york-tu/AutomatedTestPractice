using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

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

        public static void SnapshotFullScreen(string savepath) //全螢幕截圖 (C#)
        {
            Bitmap myimage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myimage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            myimage.Save(savepath);
        }

        public static void CreateSnapshotFolder(string snapshotpath) // 產生snapshot folder
        {
            try
            {
                if (Directory.Exists(snapshotpath))
                {
                    return;
                }
                else
                {
                    DirectoryInfo directiry = Directory.CreateDirectory(snapshotpath);
                }
            }
            catch (Exception)
            {

                throw;
            }
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

        } // 亂數產生新式身分證字號

        public static string CreateCellPhoneNumber() // 亂數產生行動電話號碼
        {
            Random phone = new Random();
            string cellphonenumber = $"09{(phone.Next(100000000) + 100000000).ToString().Substring(1)}";
            return cellphonenumber;
        }

        public static bool CheckResidentID(string id) // 檢核是否為現行本國人身分證字號
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

        public static bool CheckForeignerID(string id) // 檢核是否為中華民國外僑及大陸人士在台居留證證號(舊式+新式)
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

        public static string CreateRandomString(int length) // 產生指定長度 " 頭(1位大or小寫英文) + 大or小寫英文&數字 " 隨機組合字串
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

        public static string CreateRandomNumber(int length) // 產生指定長度隨機數字組合
        {
            Random r = new Random();

            string number = "";

            for (int i = 0; i < length; ++i)
                {
                    number += r.Next(0, 10); 
                }

            return number;
        }

    }

    public class LaborReliefLoan_XPath
    {
        public static string name_column_Xpath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[1]/td[2]/input"; }
        public static string ID_column_XPath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[2]/td[2]/input"; }
        public static string cellphone_column_XPath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[3]/td[2]/input"; }
        public static string birthday_column_XPath() { return "//*[@id='birth']"; }
        public static string country_dropdownlist_XPath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[1]/li/span"; }
        public static string branch_dropdownlist_XPath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[5]/td[2]/div/ul[2]/li/span"; }
        public static string date_dropdownlist_XPath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[1]/li"; }
        public static string time_dropdownlist_XPath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[6]/td[2]/div/ul[2]/li"; }
        public static string i_have_read_button_XPath() { return "//*[@id='mainform']/div[9]/div[3]/div[2]/div/div[3]/table/tbody/tr[8]/td[2]/div/a"; }
        public static string submit_button_XPath() { return "//*[@id='submit']"; }
        public static string image_verify_code_column_XPath() { return "//*[@id='captchaValue']"; }


    }
}


