using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearch.Model
{
    /// <summary>
    /// 文献引用
    /// </summary>
    public class ReferenceItem
    {
        /// <summary>
        /// 被引用文献
        /// </summary>
        public PaperIndex BeCitedPaper { get; set; }
        /// <summary>
        /// 引用文献
        /// </summary>
        public List<PaperIndex> RelatedPapers { get; set; }
    }
}
