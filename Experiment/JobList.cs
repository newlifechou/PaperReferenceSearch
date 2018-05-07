using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    /// <summary>
    /// 任务列表-代表一篇输入文档
    /// 包含多个处理单元
    /// </summary>
    public class JobList
    {
        public JobList()
        {
            Jobs = new List<JobUnit>();
        }
        public List<JobUnit> Jobs { get; set; }

        /// <summary>
        /// 该输入文档中，被引用文献数目，即任务单元的数目
        /// </summary>
        public int AllPaperCount
        {
            get
            {
                return Jobs.Count();
            }
        }
        /// <summary>
        /// 所有的引用文献
        /// </summary>
        public int AllReferenceCount
        {
            get
            {
                return Jobs.Sum(i => i.References.Count());
            }
        }

        /// <summary>
        /// 所有是自引的文献数目
        /// </summary>
        public int AllSelfReferenceCount
        {
            get
            {
                return Jobs.Sum(i => i.References.Where(k => k.ReferenceType == PaperReferenceType.Self).Count());
            }
        }
        /// <summary>
        /// 所有是他引的文献数目
        /// </summary>
        public int AllOtherReferenceCount
        {
            get
            {
                return Jobs.Sum(i => i.References.Where(k => k.ReferenceType == PaperReferenceType.Other).Count());
            }
        }
        /// <summary>
        /// 所有是未定的文献数目
        /// </summary>
        public int AllUnSetReferenceCount
        {
            get
            {
                return Jobs.Sum(i => i.References.Where(k => k.ReferenceType == PaperReferenceType.UnSet).Count());
            }
        }





    }
}
