using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaperReferenceSearch.Model
{

    public class PaperIndex
    {
        //作者编者
        public string Author { get; set; }
        //标题
        public string Title { get; set; }
        //来源出版物
        public string PublicSource { get; set; }
        //会议信息
        public string MeetingInformation { get; set; }
        //摘要
        public string Digest { get; set; }
        //引用的参考文献
        public string CitedPapers { get; set; }
        //文献类型
        public string PaperType { get; set; }
        //会议赞助方
        public string MeetingSupport { get; set; }
        //地址
        public string Address { get; set; }
        //被引用次
        public string CitedTimes { get; set; }
        //关键词
        public string Keywords { get; set; }
        //出版商信息
        public string Publisher { get; set; }
        //ISSN_ISBN
        public string ISSN_ISBN { get; set; }
        //引用的参考文献数
        public string CitedPaperTime { get; set; }
        //来源出版物缩写
        public string SourcePublishAbbr { get; set; }
        //页数
        public string PageCount { get; set; }
        //IDS
        public string IDS { get; set; }
        //语种
        public string LanguageType { get; set; }
        //WebOfService类别
        public string WebOfServiceType { get; set; }
        //研究方向
        public string ResearchDirection { get; set; }
        //基金资助方向
        public string FoundationSupportDirection { get; set; }
        //入藏号
        public string AccessionNumber { get; set; }
        //作者识别号
        public string AutherID { get; set; }
        //PubMedID
        public string PubMedID { get; set; }


    }
}
