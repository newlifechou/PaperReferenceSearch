using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    public class JobUnit
    {
        public JobUnit()
        {
            Master = new Paper();
            References = new List<Paper>();
        }
        public Paper Master { get; set; }
        public List<Paper> References { get; set; }
    }
}
