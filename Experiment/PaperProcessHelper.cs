using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace Experiment
{
    public static class PaperProcessHelper
    {

        public static ValidResult IsFormatOK(string filePath)
        {
            using (DocX doc = DocX.Load(filePath))
            {
                //判断是否都是有效段落
                foreach (var paragraph in doc.Paragraphs)
                {
                    if (!CheckFormat(paragraph.Text.Trim()))
                    {
                        return new ValidResult(false, "段落不是[空][包含:][包含附件][包含被引文献][包含引用文献][以第开头]");
                    }
                }

                //判断是否包含【作者】字段
                int count3 = doc.FindAll("作者").Count;
                if (count3 == 0)
                    return new ValidResult(false, "不包含作者段落");
                //判断被引文献和引用文献是否成对
                int count1 = doc.FindAll("被引文献").Count;
                int count2 = doc.FindAll("引用文献").Count;
                if (count1 == 0)
                    return new ValidResult(false, "不包含被引文献段落");
                if (count2 == 0)
                    return new ValidResult(false, "不包含引用文献段落");
                if (count1 != count2)
                    return new ValidResult(false, "被引文献和引用文献数目不相同");

                return new ValidResult(true, "");
            }
        }

        /// <summary>
        /// 检查段落是否符合要求
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private static bool CheckFormat(string p)
        {
            var assert1 = string.IsNullOrEmpty(p)
                || p.Contains("附件")
                || p.Contains("被引文献")
                || p.Contains("引用文献")
                || p.Contains(":")
                || p.StartsWith("第");
            return assert1;
        }

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
            string[] names = n.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
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
            string result;
            //姓名包含括号的情况
            if (name.Contains('('))
            {
                int position = name.IndexOf("(");
                result = name.Substring(0, position);
            }
            else
            {
                result = name;
            }
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
