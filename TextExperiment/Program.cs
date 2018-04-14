using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TextExperiment
{
    class Program
    {
        static void Main(string[] args)
        {
            //string line = @"标题: Histone Demethylases KDM4B and KDM6B Promote Osteogenic Differentiation of Human MSCs";
            string line = @"来源出版物: CELL STEM CELL  卷: 11  期: 1  页: 50-61  DOI: 10.1016/j.stem.2012.04.009  出版年: JUL 6 2012  ";
            int position = line.IndexOf(':');
            string prefix = line.Substring(0, position + 1);
            Console.WriteLine($"项目={prefix}");
            string content = line.Substring(position+1);
            Console.WriteLine($"内容={content.Trim()}");
            Console.Read();
        }
    }
}
