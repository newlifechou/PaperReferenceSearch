using PaperReferenceSearchService;
using PaperReferenceSearchService.Model;
using CommonHelper;

using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
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
        //状态信息管理
        private StatusHelper statusHelper = new StatusHelper();

        public MainWindowViewModel()
        {
            InputFiles = new ObservableCollection<DataFile>();
            statusMessage = "";
            inputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "DataSample");
            outputPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            canOpenOutputFolder = true;
            IsShowSelfReferenceTitleUnderLine = false;
            IsShowTotalStatistic = true;
            IsOnlyMatchFirstAuthor = false;
            IsOnlyMatchNameAbbr = true;
            IsShowMatchedAuthorHighlight = true;

            CurrentProgress = 0;
            isStartEnable = true;
            ChooseInputFolder = new RelayCommand(ActionChooseInputFolder, () => isStartEnable);
            RefreshInputFolder = new RelayCommand(ActionRefreshInputFolder, () => isStartEnable);
            ChooseOutputFolder = new RelayCommand(ActionChooseOutputFolder, () => isStartEnable);
            OpenDataFile = new RelayCommand<DataFile>(ActionOpenDataFile);
            Start = new RelayCommand(ActionStart, CanStart);

            statusHelper.ClearStatus();
            //载入欢迎信息
            string welcome = XSHelper.FileHelper.GetFullFileName("Startup.txt", XSHelper.FileHelper.GetCurrentFolderPath());
            UpdateStatusMessage(XSHelper.FileHelper.ReadText(welcome) ?? "");
            //首次载入文件
            LoadInputFiles();
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
                //2020-1-14将业务代码集中到Service中去
                isStartEnable = false;
                statusHelper.ClearStatus();

                if (InputFiles.Where(i => i.IsValid).Count() == 0)
                {
                    UpdateStatusMessage("输入文件夹中的可以处理的文件个数为0");
                    return;
                }
                var jobs = InputFiles.Where(i => i.IsValid);


                ProcessParameter parameter = new ProcessParameter();
                parameter.InputFolder = InputPath;
                parameter.OutputFolder = OutputPath;
                parameter.IsOnlyMatchFirstAuthor = IsOnlyMatchFirstAuthor;
                parameter.IsOnlyMatchNameAbbr = IsOnlyMatchNameAbbr;
                parameter.IsShowSelfReferenceTitleUnderLine = IsShowSelfReferenceTitleUnderLine;
                parameter.IsShowTotalStatistic = IsShowTotalStatistic;
                parameter.CanOpenOuputFolder = CanOpenOutputFolder;
                parameter.IsShowMatchedAuthorHighlight = IsShowMatchedAuthorHighlight;

                PaperProcessService service = new PaperProcessService();
                service.Parameter = parameter;
                service.UpdateProgressBarValue += Service_UpdateProgressBarValue;
                service.UpdateStatusInformation += Service_UpdateStatusInformation;


                service.Run(jobs);


            }

            catch (Exception ex)
            {
                UpdateStatusMessage($"[Exeption]:{ex.Message}");
            }
            finally
            {
                isStartEnable = true;
            }
        }

        #region 消息处理方法
        private void Service_UpdateStatusInformation(object sender, string e)
        {
            StatusMessage = statusHelper.InsertStatus(e);
        }

        private void Service_UpdateProgressBarValue(object sender, double e)
        {
            if (e >= 0 && e <= 100)
            {
                CurrentProgress = e;
            }
        }

        private void UpdateStatusMessage(string msg)
        {
            StatusMessage = statusHelper.InsertStatus(msg);
        }
        #endregion

        private bool CanStart()
        {
            return InputFiles.Count > 0 && isStartEnable;
        }

        private void ActionOpenDataFile(DataFile file)
        {
            try
            {
                XSHelper.FileHelper.OpenFile(file.FullName);
            }
            catch (Exception ex)
            {
                UpdateStatusMessage($"[Exeption]:{ex.Message}");
            }
        }

        private void ActionChooseOutputFolder()
        {
            string description = "请选择输出文件夹用来存放输出文档";

            PathParameter dialogResult = XSHelper.FileHelper.ShowFolderBrowserDialog(description);
            if (!dialogResult.HasSelected) return;
            if (string.IsNullOrEmpty(dialogResult.SelectPath))
            {
                UpdateStatusMessage($"选择的输出数据文件夹{dialogResult.SelectPath}不存在");
                return;
            }
            else
            {
                OutputPath = dialogResult.SelectPath;
                UpdateStatusMessage($"设置输出位置为{OutputPath}");
            }
        }

        private void ActionRefreshInputFolder()
        {
            if (!string.IsNullOrEmpty(InputPath))
            {
                CurrentProgress = 0;
                LoadInputFiles();
            }
        }

        private void ActionChooseInputFolder()
        {
            string description = "请选择存放要处理文档的文件夹\r\n确保格式是docx且符合规范要求";
            PathParameter dialogResult = XSHelper.FileHelper.ShowFolderBrowserDialog(description);
            if (!dialogResult.HasSelected)
                return;

            if (!Directory.Exists(dialogResult.SelectPath))
            {
                UpdateStatusMessage($"选择的输入数据文件夹{dialogResult.SelectPath}不存在");
                return;
            }
            else
            {
                InputPath = dialogResult.SelectPath;
                UpdateStatusMessage($"设置输入数据文件夹为{InputPath}");
                CurrentProgress = 0;

                statusHelper.ClearStatus();
                LoadInputFiles();
            }

        }

        private void LoadInputFiles()
        {
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
                    var validation_result = PaperProcessHelper.IsFormatOK(tempFile.FullName);
                    tempData.IsValid = validation_result.IsValid;
                    tempData.ValidInformation = validation_result.ValidInformation;

                    InputFiles.Add(tempData);
                    UpdateStatusMessage($"输入文件{fileName}");
                }
                int valid_count = InputFiles.Where(i => i.IsValid).Count();
                UpdateStatusMessage($"共添加了{InputFiles.Count}个输入文件,有效文件共{valid_count}个");
                UpdateStatusMessage("###无效文件处理时将跳过");

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

        private bool canOpenOutputFolder;
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

        private bool isShowSelfReferenceTitleUnderLine;
        public bool IsShowSelfReferenceTitleUnderLine
        {

            get
            {
                return isShowSelfReferenceTitleUnderLine;
            }
            set
            {
                isShowSelfReferenceTitleUnderLine = value;
                RaisePropertyChanged(nameof(IsShowSelfReferenceTitleUnderLine));
            }
        }

        private bool isShowTotalStatistic;
        public bool IsShowTotalStatistic
        {
            get
            {
                return isShowTotalStatistic;
            }
            set
            {
                isShowTotalStatistic = value;
                RaisePropertyChanged(nameof(IsShowTotalStatistic));
            }
        }
        private bool isShowMatchedAuthor;
        public bool IsShowMatchedAuthorHighlight
        {
            get
            {
                return isShowMatchedAuthor;
            }
            set
            {
                isShowMatchedAuthor = value;
                RaisePropertyChanged(nameof(IsShowMatchedAuthorHighlight));
            }
        }
        private bool isOnlyMatchFirstAuthor;
        public bool IsOnlyMatchFirstAuthor
        {
            get
            {
                return isOnlyMatchFirstAuthor;
            }
            set
            {
                isOnlyMatchFirstAuthor = value;
                RaisePropertyChanged(nameof(IsOnlyMatchFirstAuthor));
            }
        }

        private bool isOnlyMatchNameAbbr;
        public bool IsOnlyMatchNameAbbr
        {
            get
            {
                return isOnlyMatchNameAbbr;
            }
            set
            {
                isOnlyMatchNameAbbr = value;
                RaisePropertyChanged(nameof(IsOnlyMatchNameAbbr));
            }
        }
        private double currentProgress;
        public double CurrentProgress
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
        #endregion

        #region 公开命令
        public RelayCommand ChooseInputFolder { get; set; }
        public RelayCommand RefreshInputFolder { get; set; }
        public RelayCommand ChooseOutputFolder { get; set; }
        public RelayCommand<DataFile> OpenDataFile { get; set; }
        public RelayCommand Start { get; set; }
        #endregion

    }
}
