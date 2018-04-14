using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PaperReferenceSearch.Service
{
    public static class XSHelper
    {
        /// <summary>
        /// 显示文件夹选择框
        /// </summary>
        /// <param name="description"></param>
        /// <returns></returns>
        public static string ShowFolderBrowserDialog(string description)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.RootFolder = Environment.SpecialFolder.Desktop;
            dialog.Description = description;
            dialog.ShowNewFolderButton = false;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return null;

            return dialog.SelectedPath;
        }
    }
}
