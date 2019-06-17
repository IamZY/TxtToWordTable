using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TxtToWordTable
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"D:\validation";
            DirectoryInfo info = new DirectoryInfo(path);


            List<WordEntity> words = new List<WordEntity>();
            foreach (FileInfo file in info.GetFiles("*.txt"))
            {
                Console.WriteLine(file.Name + "开始...");

                using (StreamReader reader = new StreamReader(file.FullName))
                {
                    string str;
                    while ((str = reader.ReadLine()) != null)
                    {
                        if ("code|description".Equals(str)) { continue; }
                        string[] strs = str.Split(new char[] { '|' });
                        // 编码
                        string _origin = strs[0];
                        // 英文描述
                        string _en_origin = strs[1];
                        // 翻译后的结果
                        string _fanyi_res = Utils.translation(strs[1]);
                        Console.WriteLine(_origin + "----" + _en_origin + "----" + _fanyi_res);
                        words.Add(new WordEntity { origin = _origin, en_origin = _en_origin, fanyi_res = _fanyi_res });
                    }
                }


                Console.WriteLine(file.Name + "结束...");

            }

            Console.WriteLine("开始写入表格..........");
            Utils.addTable(words);


            Console.WriteLine("全部over...");


        }
    }
}
