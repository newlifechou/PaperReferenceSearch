using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using System.IO;

namespace Experiment
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.Combine(Environment.CurrentDirectory, "全例.docx");
            PaperProcess process = new PaperProcess();
            JobList jobList = process.Resolve(filePath);
            Console.WriteLine($"共检索到{jobList.Jobs.Count}篇论文");

            process.Analyse(jobList);
            Console.WriteLine("分析模式-作者姓名缩写");

            //输出
            foreach (var job in jobList.Jobs)
            {
                Console.WriteLine($"文献[{job.Master.Paragraphs["标题"].Trim()}]");
                //Console.WriteLine(job.Master.Paragraphs["作者"]);
                int self_ref_count = job.References.Where(i => i.ReferenceType != "自引").Count();
                Console.WriteLine($"共被引用{job.References.Count},自引有{job.References.Count - self_ref_count}，他引{self_ref_count},");
            }

            Console.WriteLine("done");
            Console.Read();
        }





    }
}
