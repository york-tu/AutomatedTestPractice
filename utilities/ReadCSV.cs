using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using System.Globalization;
using System.IO;
using CSVHeader;
using CsvHelper.Configuration;

namespace CSVHeader
{
    public class UserDataList
    {
        public string NAME { get; set; }
        public string ID { get; set; }
        public string PHONE { get; set; }
        public string EMAIL { get; set; }
        public string BIRTHDAY { get; set; }
        public string CARDID { get; set; }
        public string MON { get; set; }
        public string YEAR { get; set; }
        public string CHECKCODE { get; set; }
        public string ACCOUNT { get; set; }
        public string PASSWORD { get; set; }


    }
    public class UserDataMap : ClassMap<UserDataList> //当字段名为中文时，那么可以定义一个映射匹配类，CSVHelper是自动根据你的类来自动映射匹配的
    {
        UserDataMap()
        {
            Map(m => m.NAME).Name("姓名");  //使用文件列名称指定映射
            Map(m => m.ID).Name("身分證");
            Map(m => m.PHONE).Name("電話");
            Map(m => m.EMAIL).Name("EMAIL");
            Map(m => m.BIRTHDAY).Name("西元出生年月日");
            Map(m => m.CARDID).Name("信用卡卡號");
            Map(m => m.MON).Name("有效年份");
            Map(m => m.YEAR).Name("有效月份");
            Map(m => m.CHECKCODE).Name("檢查碼");
            Map(m => m.ACCOUNT).Name("帳號");
            Map(m => m.PASSWORD).Name("密碼");

        }
    }
}
