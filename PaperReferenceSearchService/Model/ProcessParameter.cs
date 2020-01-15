using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearchService.Model
{
    public class ProcessParameter
    {
        public bool CanOpenOuputFolder { get; set; }
        public bool IsShowSelfReferenceTitleUnderLine { get; set; }
        public bool IsShowTotalStatistic { get; set; }
        public bool IsOnlyMatchFirstAuthor { get; set; }
        public bool IsOnlyMatchNameAbbr { get; set; }
        public bool IsShowMatchedAuthorHighlight { get; set; }
        public string InputFolder { get; set; }
        public string OutputFolder { get; set; }
    }
}
