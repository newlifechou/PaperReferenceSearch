using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using CommonHelper;
using PaperReferenceSearchService.Model;
using System.Diagnostics;

namespace PaperReferenceSearchService
{
    /// <summary>
    /// 自引他引文献作者处理核心服务类
    /// </summary>
    public class PaperProcessService
    {

        public event EventHandler<string> UpdateStatusInformation;
        public event EventHandler<double> UpdateProgressBarValue;

        //处理参数
        public ProcessParameter Parameter;
        public PaperProcessService()
        {
            Parameter = new ProcessParameter();
        }

        private void UpdateStatus(string msg)
        {
            UpdateStatusInformation?.Invoke(this, msg);
        }

        private void UpdateProgress(double value)
        {
            UpdateProgressBarValue?.Invoke(this, value);
        }
        /// <summary>
        /// 执行主程序
        /// </summary>
        public void Run(IEnumerable<DataFile> jobs)
        {
            Stopwatch sw = new Stopwatch();
            sw.Reset();
            sw.Start();

            UpdateStatus("####开始处理格式规范的有效文件");
            string mainFolder = XSHelper.FileHelper.CreateFolder(Parameter.OutputFolder,
                                                                    XSHelper.FileHelper.GetNameByDate());

            int job_total_count = jobs.Count();
            int job_counter = 0;

            foreach (var file in jobs)
            {
                UpdateStatus($"####解析文档内容，请等待[{file.Name}]");
                var joblist = Resolve(file.FullName);
                UpdateStatus($"####分析引用类型，请等待[{file.Name}]");
                Analyse(joblist);


                UpdateStatus("####开始输出文档，请等待");

                var fileNameNoExtension = Path.GetFileNameWithoutExtension(file.FullName);

                //输出全类型
                var output_all = Path.Combine(mainFolder, $"{fileNameNoExtension}_全.docx");
                Output(joblist, output_all, OutputType.All);
                UpdateStatus($"输出文件:{output_all}");

                //输出自引类型
                var output_self = Path.Combine(mainFolder, $"{fileNameNoExtension}_自引.docx");
                Output(joblist, output_self, OutputType.Self);
                UpdateStatus($"输出文件:{output_self}");

                //输出他引类型
                var output_other = Path.Combine(mainFolder, $"{fileNameNoExtension}_他引.docx");
                Output(joblist, output_other, OutputType.Other);
                UpdateStatus($"输出文件:{output_other}");

                //输出他引类型-仅包含他引统计信息
                var output_other2 = Path.Combine(mainFolder, $"{fileNameNoExtension}_他引_仅包含他引统计.docx");
                Output(joblist, output_other2, OutputType.Other2);
                UpdateStatus($"输出文件:{output_other2}");

                //输出自引包含匹配到的作者列表
                var output_self_with_matched_authors =
                    Path.Combine(mainFolder, $"{fileNameNoExtension}_自引_包括匹配到的作者.docx");
                Output(joblist, output_self_with_matched_authors, OutputType.SelfWithMatchedAuthors);
                UpdateStatus($"输出文件:{output_self_with_matched_authors}");

                //输出测试类型
                var output_test = Path.Combine(mainFolder, $"{fileNameNoExtension}_全_调试.docx");
                Output(joblist, output_test, OutputType.Test);
                UpdateStatus($"输出文件:{output_test}");


                //记录进度
                job_counter++;
                int currentProgress = (int)(job_counter * 1.0f / job_total_count * 100);
                UpdateProgress(currentProgress);

                //Debug.WriteLine($"{currentProgress}-{job_counter}-{job_total_count}");
            }
            sw.Stop();

            UpdateStatus($"处理完毕，共处理{job_total_count}个文件，耗时{sw.ElapsedMilliseconds}ms");
            if (Parameter.CanOpenOuputFolder)
            {
                Process.Start(mainFolder);
            }
        }

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
                        //System.Diagnostics.Debug.WriteLine($"重复键");
                        tempJobUnit.Master.NoneStandardInformation += $"[{tempPara.Prefix}]行前缀重复 ";
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
                        //System.Diagnostics.Debug.WriteLine($"重复键");

                        tempPaper.NoneStandardInformation += $"[{tempPara.Prefix}]行前缀重复 ";
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
        public void Analyse(JobList jobList)
        {
            string master_name_str;
            //处理自引他引，对References进行标记
            foreach (var job in jobList.Jobs)
            {
                //找到Master的作者段落
                if (job.Master.Paragraphs.Keys.Where(i => i.Contains("作者")).Count() > 0)
                {
                    master_name_str = PaperProcessHelper.CatAuthors(job.Master.Paragraphs);
                    string[] all_master_names = PaperProcessHelper.DivideNames(master_name_str);

                    List<string> selectd_master_names = new List<string>();
                    if (Parameter.IsOnlyMatchFirstAuthor && all_master_names.Length > 0)
                    {
                        //只添加第一个作者到要匹配的列表中
                        selectd_master_names.Add(all_master_names[0].Trim());
                    }
                    else
                    {
                        all_master_names.ToList().ForEach(i => selectd_master_names.Add(i.Trim()));
                        //names.AddRange(temp_names);
                    }


                    //处理该被引文献下面的每个引用文献
                    foreach (var reference in job.References)
                    {
                        if (reference.Paragraphs.Where(i => i.Key.Contains("作者")).Count() > 0)
                        {
                            //合并多个作者行
                            string ref_name_str = PaperProcessHelper.CatAuthors(reference.Paragraphs);

                            string[] ref_names = PaperProcessHelper.DivideNames(ref_name_str);
                            foreach (var ref_name in ref_names)
                            {
                                var author = new Author { Name = ref_name, IsMatched = false };
                                reference.Authors.Add(author);
                            }

                            //遍历选定被引文献作者
                            foreach (var select_master_name in selectd_master_names)
                            {
                                string key_pattern = "";
                                bool isIncludeBracket = Parameter.IsIncludeBracket;
                                if (Parameter.IsOnlyMatchNameAbbr)
                                {
                                    //Pi, C (
                                    key_pattern = PaperProcessHelper.GetNameAbbr(select_master_name, isIncludeBracket);
                                }
                                else
                                {
                                    //(Pi, Chao)
                                    key_pattern = PaperProcessHelper.GetFullNameWithNoAbbr(select_master_name, isIncludeBracket);
                                }

                                //遍历每个文献的作者名，看是否匹配当前被引文献作者
                                foreach (var author in reference.Authors)
                                {

                                    if (author.Name.Contains(key_pattern))
                                    {
                                        author.IsMatched = true;
                                        //break;
                                    }
                                }
                            }

                            //存入匹配作者列表
                            reference.MatchedAuthors = reference.Authors.Where(i => i.IsMatched).Select(i => i.Name).ToList<string>();

                            //得到匹配到的数字，标记该引用文献属于自引还是他引
                            int result = reference.Authors.Where(i => i.IsMatched).Count();
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
        public void Output(JobList jobList, string newFilePath, OutputType outputType)
        {
            using (DocX doc = DocX.Create(newFilePath))
            {

                #region 预设样式
                Formatting normalFormat = new Formatting();
                normalFormat.FontFamily = new Font("宋体");
                normalFormat.Size = 9;

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

                Formatting formattingBluedBGLightGreen = new Formatting();
                formattingBluedBGLightGreen.FontFamily = new Font("宋体");
                formattingBluedBGLightGreen.Size = 9;
                formattingBluedBGLightGreen.FontColor = System.Drawing.Color.Blue;
                formattingBluedBGLightGreen.Highlight = Highlight.green;


                Formatting formattingUnderLine = new Formatting();
                formattingUnderLine.FontFamily = new Font("宋体");
                formattingUnderLine.Size = 9;
                formattingUnderLine.UnderlineStyle = UnderlineStyle.singleLine;

                Formatting formattingBGYellow = new Formatting();
                formattingBGYellow.FontFamily = new Font("宋体");
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
                        tempLine = "附件2 SCI-E引用 [他引-仅他引统计]";
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
                if (Parameter.IsShowTotalStatistic)
                {
                    doc.InsertParagraph();
                    tempLine = $"全局统计信息";
                    doc.InsertParagraph(tempLine, false, formattingBoldBlue);
                    tempLine = $"处理：总被引文献数={jobList.AllPaperCount},总引用文献数={jobList.AllReferenceCount}";
                    doc.InsertParagraph(tempLine, false, formattingBoldBlue);
                    tempLine = $"结果：总自引文献数={jobList.AllSelfReferenceCount}，总他引文献数={jobList.AllOtherReferenceCount}，总未定文献数={jobList.AllUnSetReferenceCount}";
                    doc.InsertParagraph(tempLine, false, formattingBoldBlue);

                    tempLine = "使用的匹配模式：";
                    if (Parameter.IsOnlyMatchNameAbbr)
                    {
                        tempLine += "使用姓名缩写";
                        if (Parameter.IsIncludeBracket)
                        {
                            tempLine += "(包含左括号)";
                        }
                        else
                        {
                            tempLine += "(不包含括号)";
                        }
                        tempLine += ";";
                    }
                    else
                    {
                        tempLine += "使用姓名全称(包含左右括号);";
                        if (Parameter.IsIncludeBracket)
                        {
                            tempLine += "(包含左右括号)";
                        }
                        else
                        {
                            tempLine += "(不包含括号)";
                        }
                        tempLine += ";";
                    }

                    if (Parameter.IsOnlyMatchFirstAuthor)
                    {
                        tempLine += "只匹配第一作者;";

                    }
                    else
                    {
                        tempLine += "匹配所有作者;";
                    }
                    doc.InsertParagraph(tempLine, false, formattingBoldBlue);


                    doc.InsertParagraph();
                }

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
                            statistic = $"(被引{job.ReferenceCount}自引{job.SelfReferenceCount}他引{job.OtherReferenceCount}未定{job.UnSetReferenceCount})";
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
                        doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, normalFormat);
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
                                    doc.InsertParagraph(tempLine, false, normalFormat);
                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        //标记自引的标题
                                        if (reference.ReferenceType == PaperReferenceType.Self && paragraph.Key == "标题")
                                        {
                                            Formatting titleFormating = Parameter.IsShowSelfReferenceTitleUnderLine ? formattingUnderLine : formattingBGYellow;
                                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}",
                                                false, titleFormating);
                                        }
                                        else
                                        {
                                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, normalFormat);
                                        }
                                    }
                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, normalFormat);
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
                                    doc.InsertParagraph(tempLine, false, normalFormat);
                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, normalFormat);
                                    }
                                    reference_Counter++;

                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, normalFormat);
                            }
                            #endregion
                            break;
                        case OutputType.Other:
                        case OutputType.Other2:
                            #region Output_Other
                            var misson_other = job.References.Where(r => r.ReferenceType == PaperReferenceType.Other);
                            if (misson_other.Count() > 0)
                            {
                                foreach (var reference in misson_other)
                                {
                                    tempLine = $"第{reference_Counter}条，共{misson_other.Count()}条";
                                    doc.InsertParagraph(tempLine, false, normalFormat);
                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, normalFormat);
                                    }
                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, normalFormat);
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
                                    doc.InsertParagraph(tempLine, false, normalFormat);

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
                                            //doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, normalFormat);
                                            #region 高亮输出匹配作者
                                            if (paragraph.Key == "作者")
                                            {
                                                //
                                                Formatting matchedAuthor = formattingBluedBGLightGreen;
                                                Paragraph p_temp = doc.InsertParagraph();
                                                p_temp.Append($"{paragraph.Key}:", normalFormat);

                                                //如果有匹配自引作者，突出显示该作者
                                                if (reference.MatchedAuthors.Count > 0 && Parameter.IsShowMatchedAuthorHighlight)
                                                {
                                                    for (int i = 0; i < reference.Authors.Count(); i++)
                                                    {
                                                        if (reference.Authors[i].IsMatched)
                                                        {
                                                            p_temp.Append($"{reference.Authors[i].Name}", matchedAuthor);
                                                        }
                                                        else
                                                        {
                                                            p_temp.Append($"{reference.Authors[i].Name}", normalFormat);
                                                        }
                                                        if (i < reference.Authors.Count - 1)
                                                        {
                                                            p_temp.Append($";", normalFormat);
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    p_temp.Append($"{paragraph.Value}", normalFormat);
                                                }
                                            }
                                            else
                                            {
                                                doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, normalFormat);
                                            }
                                            #endregion
                                        }
                                    }
                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, normalFormat);
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

                                    doc.InsertParagraph(tempLine, false, normalFormat);

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
                                        #region 输出匹配信息到每个引用文献前面
                                        List<string> uniqueMatchedStr = new List<string>();
                                        foreach (var item in reference.MatchedAuthors)
                                        {
                                            string temp_name;
                                            if (Parameter.IsOnlyMatchNameAbbr)
                                            {
                                                temp_name = PaperProcessHelper.GetNameAbbr(item,
                                                    Parameter.IsIncludeBracket);
                                            }
                                            else
                                            {
                                                temp_name = PaperProcessHelper.GetFullNameWithNoAbbr(item,
                                                    Parameter.IsIncludeBracket);
                                            }
                                            //去重
                                            if (!uniqueMatchedStr.Contains(temp_name))
                                            {
                                                uniqueMatchedStr.Add(temp_name);
                                            }
                                        }
                                        //输出匹配内容
                                        string p_line = "";
                                        foreach (var item in uniqueMatchedStr)
                                        {
                                            p_line += $"[{item}] ";
                                        }

                                        doc.InsertParagraph($"匹配用的字符串{uniqueMatchedStr.Count}个={p_line}", false, formattingBoldBlue);

                                        doc.InsertParagraph($"[匹配上的作者共{reference.MatchedAuthors.Count}人]",
                                            false, formattingBoldBlue);
                                        string matched_authors = "";
                                        reference.MatchedAuthors.ForEach(i =>
                                        {
                                            matched_authors += i + ";";
                                        });
                                        doc.InsertParagraph($"匹配上的全名=[{matched_authors}]", false, formattingBoldBlue);

                                        if (uniqueMatchedStr.Count != reference.MatchedAuthors.Count)
                                        {
                                            string matchtype = "";
                                            if (Parameter.IsOnlyMatchNameAbbr)
                                            {
                                                matchtype = "匹配缩写";
                                            }
                                            else
                                            {
                                                matchtype = "匹配全名";
                                            }
                                            doc.InsertParagraph($"###使用[{matchtype}],但匹配上作者不等于匹配字符串数目，请人工检查此段", false, formattingBoldRed);
                                        }
                                        #endregion
                                    }

                                    foreach (var paragraph in reference.Paragraphs)
                                    {
                                        //标记自引的标题
                                        if (reference.ReferenceType == PaperReferenceType.Self && paragraph.Key == "标题")
                                        {
                                            Formatting titleFormating = Parameter.IsShowSelfReferenceTitleUnderLine ? formattingUnderLine : formattingBGYellow;
                                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}",
                                                false, titleFormating);
                                        }
                                        else if (paragraph.Key == "作者")
                                        {
                                            //
                                            Formatting matchedAuthor = formattingBluedBGLightGreen;
                                            Paragraph p_temp = doc.InsertParagraph();
                                            p_temp.Append($"{paragraph.Key}:", normalFormat);

                                            //如果有匹配自引作者，突出显示该作者
                                            if (reference.MatchedAuthors.Count > 0 && Parameter.IsShowMatchedAuthorHighlight)
                                            {
                                                for (int i = 0; i < reference.Authors.Count(); i++)
                                                {
                                                    if (reference.Authors[i].IsMatched)
                                                    {
                                                        p_temp.Append($"{reference.Authors[i].Name}", matchedAuthor);
                                                    }
                                                    else
                                                    {
                                                        p_temp.Append($"{reference.Authors[i].Name}", normalFormat);
                                                    }
                                                    if (i < reference.Authors.Count - 1)
                                                    {
                                                        p_temp.Append($";", normalFormat);
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                p_temp.Append($"{paragraph.Value}", normalFormat);
                                            }
                                        }
                                        else
                                        {
                                            doc.InsertParagraph($"{paragraph.Key}:{paragraph.Value}", false, normalFormat);
                                        }
                                    }

                                    reference_Counter++;
                                }
                            }
                            else
                            {
                                doc.InsertParagraph("无", false, normalFormat);
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
