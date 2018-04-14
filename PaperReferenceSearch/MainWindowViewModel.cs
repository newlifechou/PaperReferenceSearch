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

            LoadInputFiles();

            ChooseInputFolder = new RelayCommand(ActionChooseInputFolder);
            ChooseOutputFolder = new RelayCommand(ActionChooseOutputFolder);
            OpenDataFile = new RelayCommand<DataFile>(ActionOpenDataFile);
        }

        private void ActionOpenDataFile(DataFile file)
        {
            try
            {
                System.Diagnostics.Process.Start(file.FullName);
            }
            catch (Exception ex)
            {
                AddStatusMessage(ex.Message);
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
                AddStatusMessage($"设置输出位置为{folderPath}");
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
                AddStatusMessage($"设置输入数据文件夹为{InputPath}");
                LoadInputFiles();
            }

        }

        private void LoadInputFiles()
        {
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
                    AddStatusMessage($"添加了{fileName}");
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


        #endregion

        #region 私有变量
        private StringBuilder statusMsg;
        private void AddStatusMessage(string msg)
        {
            statusMsg.AppendLine(msg);
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
