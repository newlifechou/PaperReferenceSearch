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

            LoadInputFiles();

            ChooseInputFolder = new RelayCommand(ActionChooseInputFolder);
            ChooseOutputFolder = new RelayCommand(ActionChooseOutputFolder);
            OpenDataFile = new RelayCommand<DataFile>(ActionOpenDataFile);
            Start = new RelayCommand(ActionStart, CanStart);
        }

        private void ActionStart()
        {
            if (InputFiles.Where(i => i.ValidateState).Count() == 0)
            {
                AppendStatusMessage("输入文件夹中的可以处理的文件个数为0");
                return;
            }

            var jobs = InputFiles.Where(i => i.ValidateState);

            try
            {
                Stopwatch sw = new Stopwatch();
                sw.Reset();
                sw.Start();
                PaperProcess service = new PaperProcess();
                AppendStatusMessage("####开始处理格式规范的有效文件");
                var mainFolder = Path.Combine(OutputPath, DateTime.Now.ToString("yyMMdd"));
                foreach (var file in jobs)
                {
                    var joblist = service.Resolve(file.FullName);

                    service.Analyse(joblist);

                    var fileNameNoExtension = Path.GetFileNameWithoutExtension(file.FullName);
                    if (!Directory.Exists(mainFolder))
                    {
                        Directory.CreateDirectory(mainFolder);
                    }

                    var output_all = Path.Combine(mainFolder, $"{fileNameNoExtension}_All.docx");
                    service.Output(joblist, output_all, OutputType.All);
                    AppendStatusMessage($"输出文件:{output_all}");

                    var output_self = Path.Combine(mainFolder, $"{fileNameNoExtension}_Self.docx");
                    service.Output(joblist, output_self, OutputType.Self);
                    AppendStatusMessage($"输出文件:{output_self}");

                    var output_other = Path.Combine(mainFolder, $"{fileNameNoExtension}_Other.docx");
                    service.Output(joblist, output_other, OutputType.Other);
                    AppendStatusMessage($"输出文件:{output_other}");
                }
                sw.Stop();
                AppendStatusMessage($"处理完毕，共耗时{sw.ElapsedMilliseconds}ms");

                if (CanOpenOutputFolder)
                {
                    Process.Start(mainFolder);
                }

            }
            catch (Exception ex)
            {
                AppendStatusMessage(ex.Message);
            }

        }

        private bool CanStart()
        {
            return InputFiles.Count > 0;
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
                LoadInputFiles();
            }

        }

        private void LoadInputFiles()
        {
            statusMsg.Clear();
            if (Directory.Exists(InputPath))
            {
                PaperProcess service = new PaperProcess();
                InputFiles.Clear();
                foreach (var fileName in Directory.GetFiles(InputPath, "*.docx"))
                {
                    if (!File.Exists(fileName))
                        continue;
                    FileInfo tempFile = new FileInfo(fileName);
                    DataFile tempData = new DataFile()
                    {
                        Name = tempFile.Name,
                        FullName = tempFile.FullName,
                        ValidateState = service.IsFormatOK(fileName)
                    };
                    InputFiles.Add(tempData);
                    AppendStatusMessage($"添加了{fileName}");
                }
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
        #endregion

        #region 私有变量
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
