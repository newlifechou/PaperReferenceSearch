using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    public static class TestHelper
    {
        public static void TreeIt(Dictionary<string,string> dict)
        {
            foreach (var key in dict.Keys)
            {
                System.Diagnostics.Debug.WriteLine($"{key}->{dict[key]}");
            }
        }
    }
}
