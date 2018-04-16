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
        /// 解析word文档docx
        /// 处理前确认文档基本规范
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public JobList Resolve(string filePath)
        {
            //存储文档的非空段落
            List<string> paragraphs = null;
            using (DocX doc = DocX.Load(filePath))
            {
                paragraphs = doc.Paragraphs.Where(i => !string.IsNullOrEmpty(i.Text.Trim()))
                    .Select(i => i.Text.Trim()).ToList();
            }

            bool isMasterArea = false;
            bool isGuestArea = false;
            bool isGuestBlockStart = false;

            JobList jobList = new JobList();
            JobUnit tempJobUnit = new JobUnit();
            Paper tempPaper = null;
            PaperParagraph tempPara = null;

            foreach (var line in paragraphs)
            {
                //无效行直接跳过不处理
                if (!PaperProcessHelper.CheckParagraphValid(line))
                {
                    System.Diagnostics.Debug.WriteLine("[解析]有无效行被跳过");
                    continue;
                }

                //空行和无效行跳过
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }
                if (line.Contains("附件"))
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
                        jobList.Jobs.Add(tempJobUnit);
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

                //除上面的特征行外，其余的有效行都应该包含英文冒号
                if (!line.Contains(":"))
                {
                    continue;
                }

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
                jobList.Jobs.Add(tempJobUnit);
                tempJobUnit = new JobUnit();
            }

            return jobList;

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
                if (job.Master.Paragraphs.ContainsKey("作者"))
                {
                    p = job.Master.Paragraphs["作者"];
                    string[] names = PaperProcessHelper.DivideNames(p);

                    foreach (var reference in job.References)
                    {
                        if (reference.Paragraphs.ContainsKey("作者"))
                        {
                            p = reference.Paragraphs["作者"];
                            string ref_name = p;
                            var query = names.Where(i => ref_name
                                              .Contains(PaperProcessHelper.GetNameAbbr(i)));

                            reference.MatchedAuthors = query.ToList();

                            //得到匹配到的数字
                            int result = query.Count();
                            reference.ReferenceType = result > 0 ? PaperReferenceType.Self : PaperReferenceType.Other;
                        }
                        else
                        {
                            reference.NoneStandardInformation = "此文献不包含作者段落";
                        }
                    }
                }
                else
                {
                    job.Master.NoneStandardInformation = "此文献不包含作者段落";
                }
            }

        }

        /// <summary>
        /// 输出全部到文档
        /// </summary>
        /// <param name="jobList"></param>
        /// <param name="filePath"></param>
        public void Output(JobList jobList, string newFilePath, OutputType outputType, bool isUnderLine)
        {
            using (DocX doc = DocX.Create(newFilePath))
            {
                //插入标题
                string page_title = "";
                switch (outputType)
                {
                    case OutputType.All:
                        page_title = "附件2 SCI-E引用 [自引+他引]";
                        break;
                    case OutputType.Self:
                        page_title = "附件2 SCI-E引用 [自引]";
                        break;
                    case OutputType.Other:
                        page_title = "附件2 SCI-E引用 [他引]";
                        break;
                    case OutputType.SelfWithMatchedAuthors:
                        page_title = "附件2 SCI-E引用 [自引-包含匹配到的作者]";
                        break;
                    case OutputType.Test:
                        page_title = "附件2 SCI-E引用 [调试信息，用于辅助核对分析]";
                        break;
                    default:
                        break;
                }

                doc.InsertParagraph(page_title, false, new Formatting() { Size = 14 });

                #region 预设样式
                Formatting formatting = new Formatting();
                formatting.FontFamily = new Font("宋体");
                formatting.Size = 9;

                Formatting formattingBold = new Formatting();
                formattingBold.FontFamily = new Font("宋体");
                formattingBold.Size = 9;
                formattingBold.Bold = true;

                Formatting formattingBoldBlue = new Formatting();
                formattingBoldBlue.FontFamily = new Font("宋体");
                formattingBoldBlue.Size = 9;
                formattingBoldBlue.Bold = true;
                formattingBoldBlue.FontColor = System.Drawing.Color.Blue;

                Formatting formattingBoldRed = new Formatting();
                formattingBoldRed.FontFamily = new Font("宋体");
                formattingBoldRed.Size = 9;
                formattingBoldRed.Bold = true;
                formattingBoldRed.FontColor = System.Drawing.Color.Red;

                Formatting formattingBoldRedBGYellow = new Formatting();
                formattingBoldRedBGYellow.FontFamily = new Font("宋体");
                formattingBoldRedBGYellow.Size = 9;
                formattingBoldRedBGYellow.Bold = true;
                formattingBoldRedBGYellow.FontColor = System.Drawing.Color.Red;
                formattingBoldRedBGYellow.Highlight = Highlight.yellow;

                Formatting formattingUnderLine = new Formatting();
                formatting.FontFamily = new Font("宋体");
                formattingUnderLine.Size = 9;
                formattingUnderLine.UnderlineStyle = UnderlineStyle.singleLine;

                Formatting formattingBGYellow = new Formatting();
                formatting.FontFamily = new Font("宋体");
                formattingBGYellow.Size = 9;
                formattingBGYellow.Highlight = Highlight.yellow;
                #endregion

                string tempPara = "";
                int master_Counter = 1;
                int reference_Counter = 1;
                foreach (var job in jobList.Jobs)
                {
                    //统计文献数目
                    int ref_count = job.References.Count;
                    int ref_self_count = job.References.Where(r =>
                                                    r.ReferenceType == PaperReferenceType.Self).Count();
                    int ref_other_count = job.References.Where(r =>
                                                    r.ReferenceType == PaperReferenceType.Other).Count();
                    int ref_unset = job.References.Where(r =>
                                                    r.ReferenceType == PaperReferenceType.UnSet).Count();
                    switch (outputType)
                    {
                        case OutputType.All:
                            tempPara = $"{master_Counter}.被引文献:(被引{ref_count}自引{ref_self_count}他引{ref_other_count})";
                            break;
                        case OutputType.Self:
                            tempPara = $"{master_Counter}.被引文献:(被引{ref_count}自引{ref_self_count})";
                            break;
                        case OutputType.Other:
                            tempPara = $"{master_Counter}.被引文献:(被引{ref_count}他引{ref_other_count})";
                            break;
                        case OutputType.SelfWithMatchedAuthors:
                            tempPara = $"{master_Counter}.被引文献:(被引{ref_count}自引{ref_self_count})";
                            break;
                        case OutputType.Test:
                            tempPara = $"{master_Counter}.被引文献:(被引{ref_count}自引{ref_self_count}他引{ref_other_count}未定{ref_unset})";
                            break;
                        default:
                            break;
                    }

                    doc.InsertParagraph(tempPara, false, formattingBold);

                    //测试信息输出
                    if (outputType == OutputType.Test)
                    {
                        if (job.Master.NoneStandardInformation != "")
                        {
                            doc.InsertParagraph($"[##此文献格式不规范{job.Master.NoneStandardInformation}]",
                                false, formattingBoldRedBGYellow);
                        }
                        else
                        {
                            doc.InsertParagraph($"[**此文献格式OK]", false, formattingBoldRed);
                        }
                    }

                    foreach (var paragraph in job.Master.Paragraphs)
                    {
                        doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                    }
                    master_Counter++;

                    string ref_title = "引用文献:";
                    switch (outputType)
                    {
                        case OutputType.All:
                            ref_title += "[全]";
                            break;
                        case OutputType.Self:
                            ref_title += "[自引]";
                            break;
                        case OutputType.Other:
                            ref_title += "[他引]";
                            break;
                        case OutputType.SelfWithMatchedAuthors:
                            ref_title += "[自引]";
                            break;
                        default:
                            break;
                    }

                    doc.InsertParagraph(ref_title, false, formattingBold);
                    #region 处理引用文献
                    //这里要对引用文献类型进行区分

                    reference_Counter = 1;
                    switch (outputType)
                    {
                        case OutputType.All:
                            #region Output_All
                            var misson_all = job.References;
                            if (misson_all.Count > 0)
                            {
                                foreach (var reference in misson_all)
                                {
                                    tempPara = $"第{reference_Counter}条，共{ref_count}条";
                                    doc.InsertParagraph(tempPara, false, formatting);
                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        //标记自引的标题
                                        if (reference.ReferenceType == PaperReferenceType.Self && paragraph.Key == "标题")
                                        {
                                            Formatting titleFormating = isUnderLine ? formattingUnderLine : formattingBGYellow;
                                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}",
                                                false, titleFormating);
                                        }
                                        else
                                        {
                                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                                        }
                                    }
                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, formatting);
                            }
                            #endregion
                            break;
                        case OutputType.Self:
                            #region Output_Self
                            var misson_self = job.References.Where(r => r.ReferenceType == PaperReferenceType.Self);
                            if (misson_self.Count() > 0)
                            {
                                foreach (var reference in misson_self)
                                {
                                    tempPara = $"第{reference_Counter}条，共{ref_self_count}条";
                                    doc.InsertParagraph(tempPara, false, formatting);
                                    foreach (var paragraph in reference.Paragraphs)
                                    {

                                        doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                                    }
                                    reference_Counter++;

                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, formatting);
                            }
                            #endregion
                            break;
                        case OutputType.Other:
                            #region Output_Other
                            var misson_other = job.References.Where(r => r.ReferenceType == PaperReferenceType.Other);
                            if (misson_other.Count() > 0)
                            {
                                foreach (var reference in misson_other)
                                {
                                    tempPara = $"第{reference_Counter}条，共{ref_other_count}条";
                                    doc.InsertParagraph(tempPara, false, formatting);
                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                                    }
                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, formatting);
                            }
                            #endregion
                            break;
                        case OutputType.SelfWithMatchedAuthors:
                            #region Output_SelfWithMatchedAuthors
                            var misson_self_with_authors = job.References.Where(r => r.ReferenceType == PaperReferenceType.Self);
                            if (misson_self_with_authors.Count() > 0)
                            {
                                foreach (var reference in misson_self_with_authors)
                                {
                                    tempPara = $"第{reference_Counter}条，共{ref_self_count}条";
                                    doc.InsertParagraph(tempPara, false, formatting);

                                    if (reference.MatchedAuthors.Count > 0)
                                    {
                                        doc.InsertParagraph($"匹配上的作者共{reference.MatchedAuthors.Count}人", false, formattingBoldBlue);
                                        string matched_authors = "";
                                        reference.MatchedAuthors.ForEach(i =>
                                        {
                                            matched_authors += i + ";";
                                        });
                                        doc.InsertParagraph($"[{matched_authors}]", false, formattingBoldBlue);
                                    }
                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        if (reference.ReferenceType == PaperReferenceType.Self)
                                        {
                                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                                        }
                                    }
                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, formatting);
                            }
                            #endregion
                            break;
                        case OutputType.Test:
                            #region OutputType_Test
                            var misson_test = job.References;
                            if (misson_test.Count() > 0)
                            {
                                foreach (var reference in misson_test)
                                {
                                    tempPara = $"第{reference_Counter}条，共{ref_count}条";

                                    doc.InsertParagraph(tempPara, false, formatting);

                                    //输出文献格式信息
                                    if (reference.NoneStandardInformation != "")
                                    {
                                        doc.InsertParagraph($"[##此文献格式不规范{reference.NoneStandardInformation}]",
                                            false, formattingBoldRedBGYellow);
                                    }
                                    else
                                    {
                                        doc.InsertParagraph($"[**此文献格式OK]", false, formattingBoldRed);
                                    }

                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, formatting);
                                    }
                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, formatting);
                            }
                            #endregion
                            break;
                        default:
                            break;
                    }

                    #endregion

                }

                doc.Save();

            }
        }



    }
}
