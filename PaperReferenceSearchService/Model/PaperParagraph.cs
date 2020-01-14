using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearchService.Model
{
    /// <summary>
    /// 单独一行
    /// </summary>
    public class PaperParagraph
    {
        public PaperParagraph()
        {

        }
        public PaperParagraph(string p, string c)
        {
            Prefix = p;
            Content = c;
        }
        public string Prefix { get; set; }
        public string Content { get; set; }
    }


}
