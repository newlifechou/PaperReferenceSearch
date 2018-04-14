using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace Experiment
{
    public class PaperProcess
    {
        /// <summary>
        /// 检查
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool IsFormatOK(string filePath)
        {
            bool result = true;
            DocX doc = DocX.Load(filePath);
            int count1 = doc.FindAll("被引文献").Count;
            int count2 = doc.FindAll("引用文献").Count;
            doc.Dispose();
            result = count1 == count2;
            return result;
        }
        /// <summary>
        /// 解析word文档docx
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public JobList Resolve(string filePath)
        {
            DocX doc = DocX.Load(filePath);

            bool isMasterArea = false;
            bool isGuestArea = false;
            bool isGuestBlockStart = false;

            JobList worklist = new JobList();
            JobUnit tempJobUnit = new JobUnit();
            Paper tempPaper = null;
            PaperParagraph tempPara = null;

            foreach (var p in doc.Paragraphs)
            {
                string line = p.Text.Trim();
                //空行跳过
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }

                #region 判断区域位置
                //被引文献开始了
                if (line.Contains("被引文献"))
                {
                    //判断之前有没有Paper区块还没有添加
                    if (tempPaper != null && tempPaper.Paragraphs.Count > 0)
                    {
                        tempJobUnit.References.Add(tempPaper);
                        tempPaper = new Paper();
                    }

                    //判断之前有没有没添加的任务快
                    if (tempJobUnit != null && tempJobUnit.Master.Paragraphs.Count > 0)
                    {
                        worklist.Jobs.Add(tempJobUnit);
                        tempJobUnit = new JobUnit();
                    }


                    isMasterArea = true;
                    isGuestArea = false;
                    isGuestBlockStart = false;
                    continue;
                }
                //被引用文献结束，引用文献区域开始
                else if (line.Contains("引用文献"))
                {
                    isMasterArea = false;
                    isGuestArea = true;
                    isGuestBlockStart = false;
                    continue;
                }

                if (isGuestArea && line.StartsWith("第"))
                {
                    //新引用文献开始了
                    isGuestBlockStart = true;
                    //添加之前的区块
                    if (tempPaper != null && tempPaper.Paragraphs.Count > 0)
                    {
                        tempJobUnit.References.Add(tempPaper);
                    }
                    tempPaper = new Paper();
                    continue;
                }
                #endregion

                #region 区块处理
                //处理的被引用文献的行
                if (isMasterArea)
                {
                    //读取行到Master块
                    tempPara = PaperProcessHelper.DivideParagraph(line);
                    tempJobUnit.Master.Paragraphs.Add(tempPara.Prefix, tempPara.Content);
                }
                else if (isGuestArea && isGuestBlockStart)
                {
                    //处理的是引用文献的行
                    //读取行到临时块
                    tempPara = PaperProcessHelper.DivideParagraph(line);
                    if (tempPaper != null)
                    {
                        tempPaper.Paragraphs.Add(tempPara.Prefix, tempPara.Content);
                    }
                }

            }

            //添加末尾最后一个区块
            if (tempPaper != null && tempPaper.Paragraphs.Count > 0)
            {
                tempJobUnit.References.Add(tempPaper);
                tempPaper = new Paper();
            }

            //判断之前有没有没添加的任务快
            if (tempJobUnit != null && tempJobUnit.Master.Paragraphs.Count > 0)
            {
                worklist.Jobs.Add(tempJobUnit);
                tempJobUnit = new JobUnit();
            }

            #endregion

            doc.Dispose();

            return worklist;
        }

        /// <summary>
        /// 分析自引他引
        /// </summary>
        /// <param name="jobList"></param>
        public void Analyse(JobList jobList)
        {
            string p;
            //处理自引他引，对References进行标记
            foreach (var job in jobList.Jobs)
            {
                //找到Master的作者段落
                p = job.Master.Paragraphs["作者"];
                string[] names = PaperProcessHelper.DivideNames(p);

                foreach (var reference in job.References)
                {
                    p = reference.Paragraphs["作者"];
                    string ref_name = p;
                    int result = names.Where(i => ref_name
                                      .Contains(PaperProcessHelper.GetNameAbbr(i)))
                                      .Count();
                    if (result > 0)
                    {
                        reference.ReferenceType = PaperReferenceType.Self;
                    }
                    else
                    {
                        reference.ReferenceType = PaperReferenceType.Other;
                    }
                }

            }

        }

        /// <summary>
        /// 输出全部到文档
        /// </summary>
        /// <param name="jobList"></param>
        /// <param name="filePath"></param>
        public void OutputAll(JobList jobList, string newFilePath)
        {
            DocX doc = DocX.Create(newFilePath);
            Formatting formatting = new Formatting();
            formatting.FontFamily = new Font("宋体");
            formatting.Size = 9;

            Formatting formattingUnderLine = new Formatting();
            formatting.FontFamily = new Font("宋体");
            formattingUnderLine.Size = 9;
            formattingUnderLine.UnderlineStyle = UnderlineStyle.singleLine;

            Formatting formattingBGYellow = new Formatting();
            formatting.FontFamily = new Font("宋体");
            formattingBGYellow.Size = 9;
            formattingBGYellow.Highlight = Highlight.yellow;

            string tempPara;
            for (int i = 0; i < jobList.Jobs.Count; i++)
            {
                int ref_count = jobList.Jobs[i].References.Count;
                int ref_self_count = jobList.Jobs[i].References.Where(r =>
                                                r.ReferenceType == PaperReferenceType.Self).Count();
                int ref_other_count = ref_count - ref_self_count;
                tempPara = $"{i + 1}.被引文献:(被引{ref_count}自引{ref_self_count}他引{ref_other_count})";
                doc.InsertParagraph(tempPara, false, formatting);

                foreach (var paragraph in jobList.Jobs[i].Master.Paragraphs)
                {
                    doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                }

                doc.InsertParagraph("引用文献:", false, formatting);
                //遍历References
                for (int j = 0; j < jobList.Jobs[i].References.Count; j++)
                {
                    tempPara = $"第{j + 1}条，共{ref_count}条";
                    doc.InsertParagraph(tempPara, false, formatting);
                    PaperReferenceType referenceType = jobList.Jobs[i].References[j].ReferenceType;
                    foreach (var paragraph in jobList.Jobs[i].References[j].Paragraphs)
                    {
                        if (referenceType == PaperReferenceType.Self && paragraph.Key == "标题")
                        {
                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}",
                                false, formattingBGYellow);
                        }
                        else
                        {
                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                        }
                    }
                }

            }

            doc.Save();
            doc.Dispose();
        }

        public void OutputSelfReference(JobList jobList, string newFilePath)
        {

        }

        public void OutputOtherReference(JobList jobList, string newFilePath)
        {

        }



    }
}
