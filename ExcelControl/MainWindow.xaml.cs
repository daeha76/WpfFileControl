using ExcelControl.Model;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Window = System.Windows.Window;

namespace ExcelControl
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private FileInfo[] fileInfo;
        private Dictionary<string, string> excelDictionary;
        private static string destFolder;
        private void GetFileInfo_Click(object sender, RoutedEventArgs e)
        {
            string pictureFolder = GetFolderString("사진을 불러올");
            if (pictureFolder == null) return;

            destFolder = GetFolderString("사진을 저장할");
            if (destFolder == null) return;

            DirectoryInfo directoryInfo = new DirectoryInfo(pictureFolder);
            fileInfo = directoryInfo.GetFiles();
            DirectoryInfo[] directoryInfos = directoryInfo.GetDirectories();

            foreach (var directory in directoryInfos)
            {
                ProcessDirectory(directory);
            }

            MessageBox.Show("모든 파일의 목록을 가져왔습니다.");
        }

        private void ProcessDirectory(DirectoryInfo directory)
        {
            var files = directory.GetFiles();
            fileInfo = fileInfo.Concat(files).ToArray();

            var subDirs = directory.GetDirectories();
            foreach (DirectoryInfo subDir in subDirs)
            {
                ProcessDirectory(subDir);
            }
        }

        private string GetFolderString(string text)
        {
            string description = $"{text} 폴더를 선택하세요.";
            MessageBox.Show(description);

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog
            {
                Description = description,
                ShowNewFolderButton = true,
            };

            if (folderBrowserDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                MessageBox.Show($"{text} 폴더를 지정하지 않았습니다.");
                return null;
            }

            return folderBrowserDialog.SelectedPath;
        }

        private void LoadExcelData(string filePath)
        {
            // Create an Excel Application object
            Application excelApp = new Application();

            // Open the Excel workbook
            Workbook workbook = excelApp.Workbooks.Open(filePath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Sheets[1];

            // Get the used range of the worksheet
            Range usedRange = worksheet.UsedRange;

            // Define the starting row and column
            int startRow = 7;
            int keyColumn = 19;
            int dataColumn = 5;

            // Create a dictionary to store the data
            excelDictionary = new Dictionary<string, string>();

            // Iterate through each row from the starting row to the last used row
            for (int row = startRow; row <= usedRange.Rows.Count; row++)
            {
                // Get the key and data values from the specified columns
                string key = (usedRange.Cells[row, keyColumn] as Range)?.Value?.ToString();
                string data = (usedRange.Cells[row, dataColumn] as Range)?.Value?.ToString();

                // Add the key-value pair to the dictionary
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(data))
                {
                    excelDictionary[key] = data;
                }
            }

            workbook.Close();
            excelApp.Quit();
            MessageBox.Show("엑셀의 내용을 가져왔습니다.");
        }

        private void LoadExcel_Click(object sender, RoutedEventArgs e)
        {
            if (fileInfo == null)
            {
                MessageBox.Show("증명사진 폴더를 먼저 선택해 주세요.");
                return;
            }

            // Replace "filePath" with the actual path to your Excel file
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter ="Excel Files|*.xlsx;*.xls",
                Title = "Select an Excel File",
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                LoadExcelData(filePath);
            }
        }

        private void CopyFile_Click(object sender, RoutedEventArgs e)
        {
            if (excelDictionary == null)
            {
                MessageBox.Show("Load Excel를 먼저 실행하세요.");
                return;
            }

            foreach (var dic in excelDictionary)
            {
                var fileList = fileInfo.Where(f => f.Name.Contains(dic.Value)).ToList();
                if (fileList.Count == 1)
                {
                    fileList.FirstOrDefault().FileCopy(destFolder);
                }
                else
                {
                    LoadImageFiles(fileList, dic);
                }
            }
        }

        public ObservableCollection<ImageFileInfo> ImageFiles { get; set; }
        private void LoadImageFiles(List<FileInfo> fileInfos, KeyValuePair<string,string> dic)
        {
            ImageFiles = new ObservableCollection<ImageFileInfo>();

            foreach (FileInfo fileInfo in fileInfos)
            {
                string imagePath = fileInfo.FullName;
                string fileName = fileInfo.Name;
                string filePath = fileInfo.DirectoryName;
                string lastModified = fileInfo.LastWriteTime.ToString();

                ImageFileInfo imageFileInfo = new ImageFileInfo()
                {
                    ImagePath = imagePath,
                    FileName = fileName,
                    FilePath = filePath,
                    LastModified = lastModified,
                    Key = dic.Key,
                    SearchName = dic.Value,
                    DestFolder = destFolder,
                };

                ImageFiles.Add(imageFileInfo);
            }

            ImageSelectionWindow imageSelectionWindow = new ImageSelectionWindow(ImageFiles);
            bool? result = imageSelectionWindow.ShowDialog();
        }
    }
}
