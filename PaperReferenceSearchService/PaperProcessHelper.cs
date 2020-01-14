using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using System.Text.RegularExpressions;
using PaperReferenceSearchService.Model;

namespace PaperReferenceSearchService
{
    /// <summary>
    /// 自引他引文献作者处理辅助类
    /// </summary>
    public static class PaperProcessHelper
    {
        /// <summary>
        /// 检查文献格式是否正确
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
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
        /// 注意 姓名前后可能有空格需要trim
        /// 例子 Ye, L (Ye, Ling); Fan, ZP (Fan, Zhipeng); Yu, B (Yu, Bo); Chang, J (Chang, Jia)
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        public static string[] DivideNames(string n)
        {
            if (string.IsNullOrEmpty(n))
                return null;
            string[] rawNames = n.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> names = new List<string>();
            //去除姓名前后的空格
            rawNames.ToList().ForEach(i => names.Add(i.Trim()));
            return names.ToArray();
        }

        /// <summary>
        /// 获取没有缩写部分的姓名全名
        /// </summary>
        /// <param name="name"></param>
        /// <param name="withFirstBracket"></param>
        /// <returns></returns>
        public static string GetFullNameWithNoAbbr(string name, bool withFirstBracket)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            string result;
            //姓名包含括号的情况
            if (name.Contains('('))
            {
                int position = name.IndexOf("(");
                if (withFirstBracket)
                {
                    result = name.Substring(0, position + 1);
                }
                else
                {
                    result = name.Substring(0, position);
                }
            }
            else
            {
                result = name;
            }
            return result.Trim();
        }
        /// <summary>
        /// 获取只有缩写部分的姓名缩写名
        /// 2020-1-13 利用正则表达式重写
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetNameAbbr(string name, bool withFullBracket)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            if (withFullBracket)
            {
                string pantern = @"\([\w\W]+\)";
                var match = Regex.Match(name, pantern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    //System.Diagnostics.Debug.WriteLine(match.Value);
                    return match.Value;
                }
            }
            else
            {
                string pantern = @"\(([\w\W]+)\)";
                var match = Regex.Match(name, pantern, RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    //System.Diagnostics.Debug.WriteLine(match.Value);
                    return match.Groups[1].Value;
                }
            }
            return name;
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

        /// <summary>
        /// 判断是否是英语字母
        /// </summary>
        /// <param name="firstOne"></param>
        /// <returns></returns>
        public static bool IsEnglishLetter(char firstOne)
        {
            return char.IsLower(firstOne);
        }

        /// <summary>
        /// 添加后缀
        /// </summary>
        /// <param name="dict"></param>
        /// <param name="prefix"></param>
        /// <param name="symbol"></param>
        /// <returns></returns>
        public static string AddPostfix(Dictionary<string, string> dict, string prefix, char symbol = '*')
        {
            if (dict.Count == 0) return "";
            int i = 1;
            string postfix = "";
            while (dict.ContainsKey(prefix))
            {
                postfix = new string(symbol, i);
                prefix += postfix;
                i++;
            }
            return postfix;
        }
        /// <summary>
        /// 连接作者
        /// </summary>
        /// <param name="dict"></param>
        /// <returns></returns>
        public static string CatAuthors(Dictionary<string, string> dict)
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
