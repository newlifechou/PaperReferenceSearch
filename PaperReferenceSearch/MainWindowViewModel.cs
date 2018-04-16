using Experiment;
using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using PaperReferenceSearch.Model;
using PaperReferenceSearch.Service;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace PaperReferenceSearch
{
    public class MainWindowViewModel : ViewModelBase
    {
        public MainWindowViewModel()
        {
            InputFiles = new ObservableCollection<DataFile>();
            statusMessage = "";
            statusMsg = new StringBuilder();
            inputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "DataSample");
            outputPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            canOpenOutputFolder = true;
            IsUnderLine = false;
            CurrentProgress = 0;
            isStartEnable = true;

            LoadInputFiles();

            ChooseInputFolder = new RelayCommand(ActionChooseInputFolder, () => isStartEnable);
            ChooseOutputFolder = new RelayCommand(ActionChooseOutputFolder, () => isStartEnable);
            OpenDataFile = new RelayCommand<DataFile>(ActionOpenDataFile);
            Start = new RelayCommand(ActionStart, CanStart);
        }

        private void ActionStart()
        {
            //同步版本
            //DoJobTask(jobs);
            //异步版本
            Task task = new Task(DoJobTask);
            task.Start();
        }

        private void DoJobTask()
        {
            try
            {
                isStartEnable = false;
                if (InputFiles.Where(i => i.IsValid).Count() == 0)
                {
                    AppendStatusMessage("输入文件夹中的可以处理的文件个数为0");
                    return;
                }
                var jobs = InputFiles.Where(i => i.IsValid);
                Stopwatch sw = new Stopwatch();
                sw.Reset();
                sw.Start();
                PaperProcess service = new PaperProcess();
                AppendStatusMessage("####开始处理格式规范的有效文件");
                var mainFolder = Path.Combine(OutputPath, DateTime.Now.ToString("yyMMdd"));
                if (Directory.Exists(mainFolder))
                {
                    Directory.Delete(mainFolder, true);
                }
                Directory.CreateDirectory(mainFolder);

                int job_total_count = jobs.Count();
                int job_counter = 0;

                foreach (var file in jobs)
                {
                    var joblist = service.Resolve(file.FullName);

                    service.Analyse(joblist);

                    var fileNameNoExtension = Path.GetFileNameWithoutExtension(file.FullName);

                    //输出全类型
                    var output_all = Path.Combine(mainFolder, $"{fileNameNoExtension}_全.docx");
                    service.Output(joblist, output_all, OutputType.All, IsUnderLine);
                    AppendStatusMessage($"输出文件:{output_all}");
                    //输出自引类型
                    var output_self = Path.Combine(mainFolder, $"{fileNameNoExtension}_自引.docx");
                    service.Output(joblist, output_self, OutputType.Self, IsUnderLine);
                    AppendStatusMessage($"输出文件:{output_self}");
                    //输出他引类型
                    var output_other = Path.Combine(mainFolder, $"{fileNameNoExtension}_他引.docx");
                    service.Output(joblist, output_other, OutputType.Other, IsUnderLine);
                    AppendStatusMessage($"输出文件:{output_other}");
                    //输出自引包含匹配到的作者列表
                    var output_self_with_matched_authors =
                        Path.Combine(mainFolder, $"{fileNameNoExtension}_自引_包括匹配到的作者.docx");
                    service.Output(joblist, output_self_with_matched_authors, OutputType.SelfWithMatchedAuthors, IsUnderLine);
                    AppendStatusMessage($"输出文件:{output_self_with_matched_authors}");

                    //输出测试类型
                    var output_test = Path.Combine(mainFolder, $"{fileNameNoExtension}_调试输出.docx");
                    service.Output(joblist, output_test, OutputType.Test, IsUnderLine);
                    AppendStatusMessage($"输出文件:{output_test}");


                    //记录进度
                    job_counter++;
                    CurrentProgress = (int)(job_counter * 1.0f / job_total_count * 100);
                    Debug.WriteLine($"{CurrentProgress}-{job_counter}-{job_total_count}");
                }
                sw.Stop();
                AppendStatusMessage($"处理完毕，共处理{job_total_count}个文件，耗时{sw.ElapsedMilliseconds}ms");
                AppendStatusMessage("输出类型[全]=包含自引和他引，自引文献标题被标记");
                AppendStatusMessage("输出类型[自引]，仅包含自引文献");
                AppendStatusMessage("输出类型[他引]，仅包含他引文献");
                AppendStatusMessage("输出类型[自引-包含匹配作者信息]，仅包含自引文献");
                AppendStatusMessage("输出类型[调试]，包含自引他引文献，输出每篇文献的处理情况");
                AppendStatusMessage("选择需要的，其余忽略");
                if (CanOpenOutputFolder)
                {
                    Process.Start(mainFolder);
                }
            }

            catch (Exception ex)
            {
                AppendStatusMessage(ex.Message);
            }
            finally
            {
                isStartEnable = true;
            }
        }

        private bool CanStart()
        {
            return InputFiles.Count > 0 && isStartEnable;
        }

        private void ActionOpenDataFile(DataFile file)
        {
            try
            {
                System.Diagnostics.Process.Start(file.FullName);
            }
            catch (Exception ex)
            {
                AppendStatusMessage(ex.Message);
            }
        }

        private void ActionChooseOutputFolder()
        {
            string description = "请选择输出文件夹用来存放输出文档";
            string folderPath = XSHelper.ShowFolderBrowserDialog(description);
            if (string.IsNullOrEmpty(folderPath))
            {
                return;
            }
            else
            {
                OutputPath = folderPath;
                AppendStatusMessage($"设置输出位置为{folderPath}");
            }
        }

        private void ActionChooseInputFolder()
        {
            string description = "请选择存放要处理文档的文件夹\r\n确保格式是docx且符合规范要求";
            string folderPath = XSHelper.ShowFolderBrowserDialog(description);
            if (string.IsNullOrEmpty(folderPath))
                return;

            InputPath = folderPath;
            if (!Directory.Exists(folderPath))
            {
                return;
            }
            else
            {
                InputPath = folderPath;
                AppendStatusMessage($"设置输入数据文件夹为{InputPath}");
                CurrentProgress = 0;
                LoadInputFiles();
            }

        }

        private void LoadInputFiles()
        {
            statusMsg.Clear();
            if (Directory.Exists(InputPath))
            {
                InputFiles.Clear();
                foreach (var fileName in Directory.GetFiles(InputPath, "*.docx"))
                {
                    if (!File.Exists(fileName))
                        continue;
                    FileInfo tempFile = new FileInfo(fileName);
                    DataFile tempData = new DataFile()
                    {
                        Name = tempFile.Name,
                        FullName = tempFile.FullName
                    };
                    var valid_result = PaperProcessHelper.IsFormatOK(tempFile.FullName);
                    tempData.IsValid = valid_result.IsValid;
                    tempData.ValidInformation = valid_result.ValidInformation;

                    InputFiles.Add(tempData);
                    AppendStatusMessage($"添加了{fileName}");
                }
                int valid_count = InputFiles.Where(i => i.IsValid).Count();
                AppendStatusMessage($"共添加了{InputFiles.Count}个输入文件,有效文件共{valid_count}个");
                AppendStatusMessage("程序仅辅助判定文件是否有效");
                AppendStatusMessage("程序只处理有效文件，无效文件将跳过");

            }
        }


        #region 公开属性
        public ObservableCollection<DataFile> InputFiles { get; set; }

        private string statusMessage;
        public string StatusMessage
        {
            get
            {
                return statusMessage;
            }
            set
            {
                statusMessage = value;
                RaisePropertyChanged(nameof(StatusMessage));
            }
        }

        private string inputPath;
        public string InputPath
        {
            get
            {
                return inputPath;
            }
            set
            {
                inputPath = value;
                RaisePropertyChanged(nameof(InputPath));
            }
        }
        private string outputPath;
        public string OutputPath
        {
            get
            {
                return outputPath;
            }
            set
            {
                outputPath = value;
                RaisePropertyChanged(nameof(OutputPath));
            }
        }

        public bool canOpenOutputFolder;
        public bool CanOpenOutputFolder
        {
            get
            {
                return canOpenOutputFolder;
            }
            set
            {
                canOpenOutputFolder = value;
                RaisePropertyChanged(nameof(CanOpenOutputFolder));
            }
        }

        public bool isUnderLine;
        public bool IsUnderLine
        {

            get
            {
                return isUnderLine;
            }
            set
            {
                isUnderLine = value;
                RaisePropertyChanged(nameof(IsUnderLine));
            }
        }

        public int currentProgress;
        public int CurrentProgress
        {
            get
            {
                return currentProgress;
            }
            set
            {
                currentProgress = value;
                RaisePropertyChanged(nameof(CurrentProgress));
            }
        }
        #endregion

        #region 私有变量
        private bool isStartEnable;
        private StringBuilder statusMsg;
        private void AppendStatusMessage(string msg)
        {
            //statusMsg.AppendLine(msg);
            statusMsg.Insert(0, $"{msg}\r\n");
            StatusMessage = statusMsg.ToString();
        }
        #endregion

        #region 公开命令
        public RelayCommand ChooseInputFolder { get; set; }
        public RelayCommand ChooseOutputFolder { get; set; }
        public RelayCommand<DataFile> OpenDataFile { get; set; }
        public RelayCommand Start { get; set; }
        #endregion

    }
}
