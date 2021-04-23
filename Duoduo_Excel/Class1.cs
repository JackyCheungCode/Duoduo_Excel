using ExcelDna.Integration;
using GB2260;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace Duoduo_Excel
{
    public class Class1
    {
        [ExcelFunction(Description = "小写金额转大写")]
        public static string ToDx(string strAmount)
        {
            if (strAmount == "")
            {
                return null;
            }
            string functionReturnValue = null;

            bool IsNegative = false; // 是否是负数

            if (strAmount.Trim().Substring(0, 1) == "-")

            {

                // 是负数则先转为正数

                strAmount = strAmount.Trim().Remove(0, 1);

                IsNegative = true;

            }

            string strLower = null;

            string strUpart = null;

            string strUpper = null;

            int iTemp = 0;

            // 保留两位小数123.489→123.49　　123.4→123.4

            strAmount = Math.Round(double.Parse(strAmount), 2).ToString();

            if (strAmount.IndexOf(".") > 0)

            {
                if (strAmount.IndexOf(".") == strAmount.Length - 2)

                {
                    strAmount = strAmount + "0";
                }

            }

            else

            {
                strAmount = strAmount + ".00";
            }
            strLower = strAmount;
            iTemp = 1;
            strUpper = "";
            while (iTemp <= strLower.Length)

            {

                switch (strLower.Substring(strLower.Length - iTemp, 1))

                {

                    case ".":

                        strUpart = "圆";

                        break;

                    case "0":

                        strUpart = "零";

                        break;

                    case "1":

                        strUpart = "壹";

                        break;

                    case "2":

                        strUpart = "贰";

                        break;

                    case "3":

                        strUpart = "叁";

                        break;

                    case "4":

                        strUpart = "肆";

                        break;

                    case "5":

                        strUpart = "伍";

                        break;

                    case "6":

                        strUpart = "陆";

                        break;

                    case "7":

                        strUpart = "柒";

                        break;

                    case "8":

                        strUpart = "捌";

                        break;

                    case "9":

                        strUpart = "玖";

                        break;

                }
                switch (iTemp)
                {

                    case 1:

                        strUpart = strUpart + "分";

                        break;

                    case 2:

                        strUpart = strUpart + "角";

                        break;

                    case 3:

                        strUpart = strUpart + "";

                        break;

                    case 4:

                        strUpart = strUpart + "";

                        break;

                    case 5:

                        strUpart = strUpart + "拾";

                        break;

                    case 6:

                        strUpart = strUpart + "佰";

                        break;

                    case 7:

                        strUpart = strUpart + "仟";

                        break;

                    case 8:

                        strUpart = strUpart + "万";

                        break;

                    case 9:

                        strUpart = strUpart + "拾";

                        break;

                    case 10:

                        strUpart = strUpart + "佰";

                        break;

                    case 11:

                        strUpart = strUpart + "仟";

                        break;

                    case 12:

                        strUpart = strUpart + "亿";

                        break;

                    case 13:

                        strUpart = strUpart + "拾";

                        break;

                    case 14:

                        strUpart = strUpart + "佰";

                        break;

                    case 15:

                        strUpart = strUpart + "仟";

                        break;

                    case 16:

                        strUpart = strUpart + "万";

                        break;

                    default:

                        strUpart = strUpart + "";

                        break;

                }



                strUpper = strUpart + strUpper;

                iTemp = iTemp + 1;

            }

            strUpper = strUpper.Replace("零拾", "零");

            strUpper = strUpper.Replace("零佰", "零");

            strUpper = strUpper.Replace("零仟", "零");

            strUpper = strUpper.Replace("零零零", "零");

            strUpper = strUpper.Replace("零零", "零");

            strUpper = strUpper.Replace("零角零分", "整");

            strUpper = strUpper.Replace("零分", "整");

            strUpper = strUpper.Replace("零角", "零");

            strUpper = strUpper.Replace("零亿零万零圆", "亿圆");

            strUpper = strUpper.Replace("亿零万零圆", "亿圆");

            strUpper = strUpper.Replace("零亿零万", "亿");

            strUpper = strUpper.Replace("零万零圆", "万圆");

            strUpper = strUpper.Replace("零亿", "亿");

            strUpper = strUpper.Replace("零万", "万");

            strUpper = strUpper.Replace("零圆", "圆");

            strUpper = strUpper.Replace("零零", "零");


            // 对壹圆以下的金额的处理

            if (strUpper.Substring(0, 1) == "圆")

            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "零")

            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }

            if (strUpper.Substring(0, 1) == "角")
            {
                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }

            if (strUpper.Substring(0, 1) == "分")
            {

                strUpper = strUpper.Substring(1, strUpper.Length - 1);
            }
            if (strUpper.Substring(0, 1) == "整")
            {
                strUpper = "零圆整";
            }

            functionReturnValue = strUpper;
            if (IsNegative == true)
            {
                return "负" + functionReturnValue;
            }

            else
            {
                return functionReturnValue;
            }
        }

        [ExcelFunction(Description = "提取数字")]
        public static string Get_Num(string str) => System.Text.RegularExpressions.Regex.Replace(str, @"[^0-9]+", "");

        [ExcelFunction(Description = "码转换为米")]
        public static double MaToMi(double Ma) => Ma * 0.9144;

        [ExcelFunction(Description = "米转换为码")]
        public static double MiToMa(double Mi) => Mi / 0.9144;

        [ExcelFunction(Description = "中国银行汇率查需")]
        public static string BocHuilv(double money, DateTime date)
        {
            string dateOne = date.ToShortDateString().ToString().Replace("-", "");
            string url = "http://api.k780.com/?app=finance.rate_cnyquot_history&curno=USD&bankno=BOC&date=" + dateOne + "&appkey=25441&sign=93d579ace3e2b38a585e7e32bd37e0e7&format=json";
            string result = "";
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            Stream stream = resp.GetResponseStream();
            //获取内容
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                result = reader.ReadToEnd();
            }
            return result;
        }


        [ExcelFunction(Description = "用符号连接多个单元格的值")]
        public static string Link(Object[] args, string linkeChar)
        {
            string res = null;
            foreach (var item in args)
            {
                res += item + linkeChar;
            }
            return res;
        }

        [ExcelFunction(Description = "格式化成数值")]
        public static double ToDouble(Object str)
        {
            try
            {
                return Convert.ToDouble(str);
            }
            catch (Exception)
            {

                return 0;
            }
        }

        [ExcelFunction(Description = "提取省份")]
        public static string GetProvinces(string pro)
        {
            string res = null;
            foreach (var item in Getp())
            {
                //宁夏回族自治区，广西壮族自治区，直辖市,新疆维吾尔族自治区
                if (pro.Contains(item.Replace("省", "").Replace("回族自治区", "").Replace("市", "").Replace("壮族自治区", "").Replace("特别行政区", "")
                    .Replace("维吾尔族自治区", "")))
                {
                    res = item;
                    break;
                }
            }
            return res;
        }

        private static List<string> Getp()
        {
            List<string> provinces = new List<string>();
            GB2260.Gb2260 gb = Gb2260Factory.Create();

            foreach (var item in gb.Provinces)
            {
                provinces.Add(item.Name);
            }
            provinces.Add("台湾省");
            return provinces;
        }


        /// <summary>
        /// 拆分一个单元格
        /// </summary>
        /// <param name="str">传入的单元格的值</param>
        /// <param name="num">获取的部分</param>
        ///  <param name="spChar">用什么拆分</param>
        /// <returns></returns>
        [ExcelFunction(Description = "拆分值")]
        public static string ChaiFen(string str, string spChar, int num)
        {
            string[] strs = str.Split(spChar.ToCharArray()[0]);
            return strs[num];
        }
    }
}
