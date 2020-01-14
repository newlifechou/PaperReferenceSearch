using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearchService.Model
{
    /// <summary>
    /// 单独一个处理单元
    /// 包含一个被引文献
    /// 包含多个引用文献
    /// </summary>
    public class Paper
    {
        public Paper()
        {
            ReferenceType = PaperReferenceType.UnSet;
            Paragraphs = new Dictionary<string, string>();
            MatchedAuthors = new List<string>();
            NoneStandardInformation = "";
        }
        public List<string> MatchedAuthors { get; set; }
        public PaperReferenceType ReferenceType { get; set; }
        public Dictionary<string, string> Paragraphs { get; set; }

        public string NoneStandardInformation { get; set; }

        public void Add(string prefix, string content)
        {
            if (!Paragraphs.ContainsKey(prefix))
            {
                Paragraphs.Add(prefix, content);
            }
        }

    }
}
