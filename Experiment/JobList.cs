using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    /// <summary>
    /// 任务列表
    /// 包含多个处理单元
    /// </summary>
    public class JobList
    {
        public JobList()
        {
            Jobs = new List<JobUnit>();
        }
        public List<JobUnit> Jobs { get; set; }
    }
}
