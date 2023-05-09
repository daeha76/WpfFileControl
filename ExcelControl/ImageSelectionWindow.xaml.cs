using ExcelControl.Model;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Windows;

namespace ExcelControl
{
    /// <summary>
    /// ImageSelectionWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ImageSelectionWindow : Window
    {
        public ImageSelectionWindow()
        {
            InitializeComponent();
        }

        public ImageSelectionWindow(ObservableCollection<ImageFileInfo> imageFileInfos)
        {
            InitializeComponent();
            imageListBox.ItemsSource = imageFileInfos;

            SearchName = imageFileInfos.FirstOrDefault().SearchName;
            SearchKey = imageFileInfos.FirstOrDefault().Key;
        }

        public string SearchName { get; set; }
        public string SearchKey { get; set; }
        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (imageListBox.SelectedItems.Count > 0)
            {
                foreach (ImageFileInfo selectedItem in imageListBox.SelectedItems)
                {
                    string sourceFilePath = selectedItem.ImagePath;
                    string destinationFilePath = Path.Combine(selectedItem.DestFolder, selectedItem.FileName);

                    try
                    {
                        File.Copy(sourceFilePath, destinationFilePath, true);
                        MessageBox.Show($"{selectedItem.FileName}을(를) 복사하였습니다.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"복사 중 오류가 발생하였습니다: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("이미지를 선택해주세요.");
            }
            Close();
        }
    }
}
