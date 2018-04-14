using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    public static class PaperProcessHelper
    {
        /// <summary>
        /// 利用冒号分隔段落成两个字符串
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static PaperParagraph DivideParagraph(string p)
        {
            if (string.IsNullOrEmpty(p))
                return null;
            int position = p.IndexOf(':');
            string prefix = p.Substring(0, position);
            string content = p.Substring(position + 1);
            return new PaperParagraph(prefix, content);
        }

        /// <summary>
        /// 把用;隔开的姓名字符串拆成多个姓名数组
        /// 例子 Ye, L (Ye, Ling); Fan, ZP (Fan, Zhipeng); Yu, B (Yu, Bo); Chang, J (Chang, Jia)
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        public static string[] DivideNames(string n)
        {
            if (string.IsNullOrEmpty(n))
                return null;

            string[] names = n.Split(';');
            return names;
        }

        /// <summary>
        /// 返回姓名字符串缩写
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetNameAbbr(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            int position = name.IndexOf("(");
            string result = name.Substring(0, position);

            return result.Trim();
        }
        /// <summary>
        /// 获取姓名字符串全称（包含括号）
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetNameFull(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            int position1 = name.IndexOf('(');
            int position2 = name.IndexOf(')');
            string result = name.Substring(position1, position2 - position1 + 1);

            return result.Trim();
        }


    }
}
