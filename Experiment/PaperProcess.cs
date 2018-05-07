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

                    if (tempJobUnit.Master.Paragraphs.ContainsKey(tempPara.Prefix))
                    {
                        System.Diagnostics.Debug.WriteLine($"重复键");
                        TestHelper.TreeIt(tempJobUnit.Master.Paragraphs);
                        tempJobUnit.Master.NoneStandardInformation += $"[{tempPara.Prefix}]行前缀重复";
                        //如果有重复，添加新后缀
                        tempPara.Prefix += PaperProcessHelper.AddPostfix(tempJobUnit.Master.Paragraphs,
                            tempPara.Prefix);
                    }
                    tempJobUnit.Master.Paragraphs.Add(tempPara.Prefix, tempPara.Content);
                }
                else if (isGuestArea && isGuestBlockStart)
                {
                    //处理的是引用文献的行
                    //读取行到临时块
                    tempPara = PaperProcessHelper.DivideParagraph(line);
                    if (tempPaper.Paragraphs.ContainsKey(tempPara.Prefix))
                    {
                        System.Diagnostics.Debug.WriteLine($"重复键");
                        TestHelper.TreeIt(tempPaper.Paragraphs);

                        tempPaper.NoneStandardInformation += $"[{tempPara.Prefix}]行前缀重复";
                        tempPara.Prefix += PaperProcessHelper.AddPostfix(tempPaper.Paragraphs,
                            tempPara.Prefix);
                    }
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
        public void Analyse(JobList jobList, bool isOnlyMatchFirstAuthor = false)
        {
            string p;
            //处理自引他引，对References进行标记
            foreach (var job in jobList.Jobs)
            {
                //找到Master的作者段落
                if (job.Master.Paragraphs.Keys.Where(i => i.Contains("作者")).Count() > 0)
                {
                    p = PaperProcessHelper.CatAuthors(job.Master.Paragraphs);
                    string[] temp_names = PaperProcessHelper.DivideNames(p);

                    List<string> names = new List<string>();
                    if (isOnlyMatchFirstAuthor && temp_names.Length > 0)
                    {
                        //只添加第一个作者到要匹配的列表中
                        names.Add(temp_names[0]);
                    }
                    else
                    {
                        names.AddRange(temp_names);
                    }
                    //处理该被引文献下面的每个引用文献
                    foreach (var reference in job.References)
                    {
                        if (reference.Paragraphs.Where(i => i.Key.Contains("作者")).Count() > 0)
                        {
                            //合并多个作者行
                            string ref_name_str = PaperProcessHelper.CatAuthors(reference.Paragraphs);
                            var match_names = names.Where(i => ref_name_str
                                              .Contains(PaperProcessHelper.GetNameAbbr(i)));

                            reference.MatchedAuthors = match_names.ToList();

                            //得到匹配到的数字，标记该引用文献属于自引还是他引
                            int result = match_names.Count();
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
        public void Output(JobList jobList, string newFilePath, OutputType outputType, bool isUnderLine = false)
        {
            using (DocX doc = DocX.Create(newFilePath))
            {

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

                //插入标题
                string tempLine = "";
                switch (outputType)
                {
                    case OutputType.All:
                        tempLine = "附件2 SCI-E引用 [自引+他引]";
                        break;
                    case OutputType.Self:
                        tempLine = "附件2 SCI-E引用 [自引]";
                        break;
                    case OutputType.Other:
                        tempLine = "附件2 SCI-E引用 [他引]";
                        break;
                    case OutputType.Other2:
                        tempLine = "附件2 SCI-E引用 [只有他引统计]";
                        break;
                    case OutputType.SelfWithMatchedAuthors:
                        tempLine = "附件2 SCI-E引用 [自引-包含匹配到的作者]";
                        break;
                    case OutputType.Test:
                        tempLine = "附件2 SCI-E引用 [调试信息，用于辅助核对分析]";
                        break;
                    default:
                        break;
                }

                doc.InsertParagraph(tempLine, false, new Formatting() { Size = 14 });

                //插入全局统计信息
                doc.InsertParagraph();
                tempLine = $"全局统计信息";
                doc.InsertParagraph(tempLine, false, formattingBoldBlue);
                tempLine = $"共处理：总被引文献={jobList.AllPaperCount},总引用文献={jobList.AllReferenceCount}";
                doc.InsertParagraph(tempLine, false, formattingBoldBlue);
                tempLine = $"结果：总自引={jobList.AllSelfReferenceCount}，总他引={jobList.AllOtherReferenceCount}，总未定={jobList.AllUnSetReferenceCount}";
                doc.InsertParagraph(tempLine, false, formattingBoldBlue);
                doc.InsertParagraph();
                //编号
                int master_Counter = 1;
                foreach (var job in jobList.Jobs)
                {

                    tempLine = $"{master_Counter}.被引文献:";
                    var p_reference_title = doc.InsertParagraph(tempLine, false, formattingBold);
                    string statistic = "";
                    switch (outputType)
                    {
                        case OutputType.All:
                            statistic = $"(被引{job.ReferenceCount}自引{job.SelfReferenceCount}他引{job.OtherReferenceCount})";
                            break;
                        case OutputType.Self:
                            statistic = $"(被引{job.ReferenceCount}自引{job.SelfReferenceCount})";
                            break;
                        case OutputType.Other:
                            statistic = $"(被引{job.ReferenceCount}他引{job.OtherReferenceCount})";
                            break;
                        case OutputType.Other2:
                            statistic = $"(他引{job.OtherReferenceCount})";
                            break;
                        case OutputType.SelfWithMatchedAuthors:
                            statistic = $"(被引{job.ReferenceCount}自引{job.SelfReferenceCount})";
                            break;
                        case OutputType.Test:
                            statistic = $"(被引{job.ReferenceCount}自引{job.SelfReferenceCount}他引{job.OtherReferenceCount})";
                            break;
                        default:
                            break;
                    }
                    p_reference_title.InsertText(tempLine.Length, statistic, false, formattingBoldBlue);

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
                            doc.InsertParagraph($"[**此文献格式OK]", false, formattingBoldBlue);
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
                        case OutputType.Other2:
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

                    int reference_Counter = 1;
                    switch (outputType)
                    {
                        case OutputType.All:
                            #region Output_All
                            var misson_all = job.References;
                            if (misson_all.Count > 0)
                            {
                                foreach (var reference in misson_all)
                                {
                                    tempLine = $"第{reference_Counter}条，共{misson_all.Count}条";
                                    doc.InsertParagraph(tempLine, false, formatting);
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
                                    tempLine = $"第{reference_Counter}条，共{misson_self.Count()}条";
                                    doc.InsertParagraph(tempLine, false, formatting);
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
                                    tempLine = $"第{reference_Counter}条，共{misson_other.Count()}条";
                                    doc.InsertParagraph(tempLine, false, formatting);
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
                                    tempLine = $"第{reference_Counter}条，共{misson_self_with_authors.Count()}条";
                                    doc.InsertParagraph(tempLine, false, formatting);

                                    if (reference.MatchedAuthors.Count > 0)
                                    {
                                        doc.InsertParagraph($"[匹配上的作者共{reference.MatchedAuthors.Count}人]",
                                            false, formattingBoldBlue);
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
                                    tempLine = $"第{reference_Counter}条，共{misson_test.Count}条";

                                    doc.InsertParagraph(tempLine, false, formatting);

                                    //输出文献格式信息
                                    if (reference.NoneStandardInformation != "")
                                    {
                                        doc.InsertParagraph($"[##此文献格式不规范{reference.NoneStandardInformation}]",
                                            false, formattingBoldRedBGYellow);
                                    }
                                    else
                                    {
                                        doc.InsertParagraph($"[**此文献格式OK]", false, formattingBoldBlue);
                                    }

                                    if (reference.MatchedAuthors.Count > 0)
                                    {
                                        doc.InsertParagraph($"[匹配上的作者共{reference.MatchedAuthors.Count}人]",
                                            false, formattingBoldBlue);
                                        string matched_authors = "";
                                        reference.MatchedAuthors.ForEach(i =>
                                        {
                                            matched_authors += i + ";";
                                        });
                                        doc.InsertParagraph($"[{matched_authors}]", false, formattingBoldBlue);
                                    }

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
