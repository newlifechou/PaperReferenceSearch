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
            //一次读入内容
            List<string> lines = null;
            using (DocX doc = DocX.Load(filePath))
            {
                //读取所有非空行
                lines = doc.Paragraphs.Where(i => i.Text.Trim() != "")
                    .Select(i => i.Text.Trim()).ToList();
            }

            if (lines.Count == 0)
            {
                return new ValidResult(false, "文档不能为空");
            }

            //if (lines.Where(i => !CheckParagraphValid(i)).Count() > 0)
            //{
            //    return new ValidResult(false, "有段落不是[空] [包含：被引文献 引用文献] [以第开头]中一种");
            //}

            //判断是否包含【作者】字段
            int count_authors = lines.Where(i => i.Contains("作者")).Count();
            if (count_authors == 0)
                return new ValidResult(false, "包含0个作者段落");

            //判断被引文献和引用文献是否成对
            int count_be_ref = lines.Where(i => i.Contains("被引文献")).Count();
            int count_ref = lines.Where(i => i.Contains("引用文献")).Count();
            if (count_be_ref == 0)
                return new ValidResult(false, "包含0个被引文献段落");
            if (count_ref == 0)
                return new ValidResult(false, "包含0个引用文献段落");
            if (count_be_ref != count_ref)
                return new ValidResult(false, "被引文献和引用文献数目不配对");

            return new ValidResult(true, "");

        }

        /// <summary>
        /// 是否是可以处理的段落之一
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static bool CheckParagraphValid(string p)
        {
            var assert1 = string.IsNullOrEmpty(p)
                || p.Contains("附件")
                || p.Contains("被引文献")
                || p.Contains("引用文献")
                || p.StartsWith("第")
                || p.Contains(":");
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
            return new PaperParagraph(prefix.Trim(), content.Trim());
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

        public static bool IsEnglishLetter(char firstOne)
        {
            return char.IsLower(firstOne);
        }

        public static string AddPostfix(Dictionary<string, string> dict, string prefix)
        {
            if (dict.Count == 0) return "";
            int i = 1;
            string postfix = "";
            while (dict.ContainsKey(prefix))
            {
                postfix = new string('A', i);
                prefix += postfix;
                i++;
            }
            return postfix;
        }

        public static string CatAuthors(Dictionary<string,string> dict)
        {
            StringBuilder sb = new StringBuilder();
            var query = from p in dict
                        where p.Key.Contains("作者")
                        select p;
            foreach (var p in query)
            {
                sb.Append(p.Value);
                sb.Append(";");
            }
            return sb.ToString();
        }

    }
}
