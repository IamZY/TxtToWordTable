using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TxtToWordTable
{
    public class WordEntity
    {
        public string origin { get; set; }
        // 英文描述
        public string en_origin { get; set; }
        // 翻译后的结果
        public string fanyi_res { get; set; }
    }
}
