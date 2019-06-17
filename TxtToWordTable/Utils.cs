using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace TxtToWordTable
{
    public static class Utils
    {
        //对字符串做md5加密
        public static string GetMD5WithString(string input)
        {
            if (input == null)
            {
                return null;
            }
            MD5 md5Hash = MD5.Create();
            //将输入字符串转换为字节数组并计算哈希数据  
            byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input));
            //创建一个 Stringbuilder 来收集字节并创建字符串  
            StringBuilder sBuilder = new StringBuilder();
            //循环遍历哈希数据的每一个字节并格式化为十六进制字符串  
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            //返回十六进制字符串  
            return sBuilder.ToString();
        }


        /// <summary>
        /// 调用百度翻译API进行翻译
        /// 详情可参考http://api.fanyi.baidu.com/api/trans/product/apidoc
        /// </summary>
        /// <param name="q">待翻译字符</param>
        /// <param name="from">源语言</param>
        /// <param name="to">目标语言</param>
        /// <returns></returns>
        public static TranslationResult GetTranslationFromBaiduFanyi(string q, string from, string to)
        {
            //可以直接到百度翻译API的官网申请
            //此处的都是子丰随便写的，所以肯定是不能用的
            //一定要去申请，不然程序的翻译功能不能使用
            string appId = "20190616000307966";
            string password = "VJwJ2nmsOduSaNOVrYz2";

            // 转换编码
            UTF8Encoding utf8 = new UTF8Encoding();
            Byte[] encodedBytes = utf8.GetBytes(q);
            String new_q = utf8.GetString(encodedBytes);

            string jsonResult = String.Empty;
            //随机数
            string randomNum = System.DateTime.Now.Millisecond.ToString();
            //md5加密
            string md5Sign = GetMD5WithString(appId + q + randomNum + password);
            //url
            string url = String.Format("http://api.fanyi.baidu.com/api/trans/vip/translate?q={0}&from={1}&to={2}&appid={3}&salt={4}&sign={5}",
                new_q,
                from,
                to,
                appId,
                randomNum,
                md5Sign
                );
            WebClient wc = new WebClient();
            try
            {
                jsonResult = wc.DownloadString(url);
            }
            catch
            {
                jsonResult = string.Empty;
            }
            //解析json
            JavaScriptSerializer jss = new JavaScriptSerializer();
            TranslationResult result = jss.Deserialize<TranslationResult>(jsonResult);
            return result;
        }


        /// <summary>
        /// 将英文翻译为中文
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static string translation(string source)
        {
            TranslationResult result = GetTranslationFromBaiduFanyi(source, "en", "zh");
            //判断是否出错
            if (result.Error_code == null)
            {
                return result.Trans_result[0].Dst;
            }
            else
            {
                //检查appid和密钥是否正确
                return "翻译出错，错误码：" + result.Error_code + "，错误信息：" + result.Error_msg;
            }
        }


        public static void addTable(List<WordEntity> words)
        {
            //创建一个Document类实例，并添加section
            Document doc = new Document();
            Section section = doc.AddSection();

            //添加表格
            Table table = section.AddTable(true);

            //添加表格第1行
            TableRow row1 = table.AddRow();

            //添加第1个单元格到第1行
            TableCell cell1 = row1.AddCell();
            cell1.AddParagraph().AppendText("编码");

            //添加第2个单元格到第1行
            TableCell cell2 = row1.AddCell();
            cell2.AddParagraph().AppendText("英文描述");

            //添加第3个单元格到第1行
            TableCell cell3 = row1.AddCell();
            cell3.AddParagraph().AppendText("中文描述");



            foreach (WordEntity w in words)
            {
                //添加表格第2行
                TableRow row2 = table.AddRow(true, false);

                //添加第6个单元格到第2行
                TableCell cell6 = row2.AddCell();
                cell6.AddParagraph().AppendText(w.origin);

                //添加第7个单元格到第2行
                TableCell cell7 = row2.AddCell();
                cell7.AddParagraph().AppendText(w.en_origin);

                //添加第8个单元格到第2行
                TableCell cell8 = row2.AddCell();
                cell8.AddParagraph().AppendText(w.fanyi_res);
            }

            //table.AutoFit()
            
            table.AutoFit(AutoFitBehaviorType.AutoFitToWindow);

            //保存文档
            doc.SaveToFile(@"d:\Table.docx");
        }


    }
}
