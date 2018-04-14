using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Collections.ObjectModel;
using PaperReferenceSearch.Model;
using Experiment;



namespace PaperReferenceSearch
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            InputFiles = new ObservableCollection<DataFile>();

            this.DataContext = this;
        }

        public ObservableCollection<DataFile> InputFiles { get; set; }

        private void BtnInputFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;
            string folderPath = dialog.SelectedPath;
            TxtInputPath.Text = folderPath;


            if (!Directory.Exists(folderPath))
                return;
            PaperProcess process = new PaperProcess();
            InputFiles.Clear();
            foreach (var fileName in Directory.GetFiles(folderPath, "*.docx"))
            {
                if (!File.Exists(fileName))
                    continue;
                FileInfo tempFile = new FileInfo(fileName);
                DataFile tempData = new DataFile()
                {
                    FileName = tempFile.Name,
                    ValidateState = process.IsFormatOK(fileName)
                };
                InputFiles.Add(tempData);
            }

        }
    }
}
