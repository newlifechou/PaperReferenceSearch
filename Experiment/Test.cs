using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Experiment
{
    public class Test
    {
        public void TestPaperProcessHelper()
        {
            string n = @"Ye, L (Ye, Ling); Fan, ZP (Fan, Zhipeng); Yu, B (Yu, Bo); Chang, J (Chang, Jia); Al Hezaimi, K (Al Hezaimi, Khalid); Zhou, XD (Zhou, Xuedong); Park, NH (Park, No-Hee); Wang, CY (Wang, Cun-Yu)";
            string[] names = PaperProcessHelper.DivideNames(n);
            foreach (var name in names)
            {
                Console.WriteLine(name.Trim());
                Console.WriteLine(PaperProcessHelper.GetNameAbbr(name.Trim(), true).Trim());
                Console.WriteLine(PaperProcessHelper.GetNameFull(name.Trim()).Trim());
            }
        }
    }
}
