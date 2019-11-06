using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SGCCExcelOp;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Btn_Page1_Click(object sender, RoutedEventArgs e)
        {
            changePage.Visibility = Visibility.Hidden;
            dailyPage.Visibility = Visibility.Visible;
            safeMeetPage.Visibility = Visibility.Hidden;
            monthCount.Visibility = Visibility.Hidden;
            morePage.Visibility = Visibility.Hidden;
        }

        private void Btn_Page2_Click(object sender, RoutedEventArgs e)
        {
            changePage.Visibility = Visibility.Hidden;
            dailyPage.Visibility = Visibility.Hidden;
            safeMeetPage.Visibility = Visibility.Visible;
            monthCount.Visibility = Visibility.Hidden;
            morePage.Visibility = Visibility.Hidden;
        }

        private void Btn_Page3_Click(object sender, RoutedEventArgs e)
        {
            changePage.Visibility = Visibility.Hidden;
            dailyPage.Visibility = Visibility.Hidden;
            safeMeetPage.Visibility = Visibility.Hidden;
            monthCount.Visibility = Visibility.Visible;
            morePage.Visibility = Visibility.Hidden;
        }

        private void Btn_Page4_Click(object sender, RoutedEventArgs e)
        {
            changePage.Visibility = Visibility.Hidden;
            dailyPage.Visibility = Visibility.Hidden;
            safeMeetPage.Visibility = Visibility.Hidden;
            monthCount.Visibility = Visibility.Hidden;
            morePage.Visibility = Visibility.Visible;
        }

        private void e5000_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            e5000Path.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void e3000_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            e3000Path.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void jisuanWork_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            planWork.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void anquan_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            eleDanger.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void lihuiSave_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择生成EXCEL文件夹";
                dialog.ShowNewFolderButton = false;
                dialog.RootFolder = Environment.SpecialFolder.Desktop;
                dialog.ShowDialog();
                if (dialog.SelectedPath == string.Empty)
                {
                    lihuiSave.Text = "请选择目录！";
                    MessageBox.Show("请选择需要存储的位置", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                lihuiSave.Text = dialog.SelectedPath;
            }

            if (string.IsNullOrEmpty(safeMeetDate.Text))
            {
                MessageBox.Show("请选择日期", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var date = safeMeetDate.Text.Replace('/', '-');
            var safeMeetPath = AppDomain.CurrentDomain.BaseDirectory + @"\SafeMeeting\";
            var filesList = Directory.GetFiles(safeMeetPath);
            if (filesList.Length > 0)
            {
                foreach (var file in filesList)
                {
                    if (file.Contains(date))
                    {
                        File.Copy(file, lihuiSave.Text + "\\" + Path.GetFileName(file), true);
                    }
                }
            }

            MessageBox.Show("生成完毕", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        private void monthPath_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            monthPath.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void saveCount_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择生成EXCEL文件夹";
                dialog.ShowNewFolderButton = false;
                dialog.RootFolder = Environment.SpecialFolder.Desktop;
                dialog.ShowDialog();
                if (dialog.SelectedPath == string.Empty)
                {
                    saveCountPath.Text = "请选择目录！";
                    MessageBox.Show("请选择需要存储的位置", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                saveCountPath.Text = dialog.SelectedPath;
            }

            if (string.IsNullOrEmpty(countDate.Text))
            {
                MessageBox.Show("请选择日期", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var date = countDate.Text.Replace('/', '-');
            var countPath = AppDomain.CurrentDomain.BaseDirectory + @"\DianliangCount\";
            var filesList = Directory.GetFiles(countPath);
            if (filesList.Length > 0)
            {
                foreach (var file in filesList)
                {
                    if (file.Contains(date))
                    {
                        File.Copy(file, saveCountPath.Text + "\\" + Path.GetFileName(file), true);
                    }
                }
            }

            MessageBox.Show("生成完毕", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        private void save_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择生成EXCEL文件夹";
                dialog.ShowNewFolderButton = false;
                dialog.RootFolder = Environment.SpecialFolder.Desktop;
                dialog.ShowDialog();
                if (dialog.SelectedPath == string.Empty)
                {
                    savePath.Text = "请选择目录！";
                    MessageBox.Show("请选择需要存储的位置", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                savePath.Text = dialog.SelectedPath;
            }

            if (string.IsNullOrEmpty(dianliangDate.Text))
            {
                MessageBox.Show("请选择日期", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var date = dianliangDate.Text.Replace('/', '-');
            var dianLiangPath = AppDomain.CurrentDomain.BaseDirectory + @"\DianLiang\";
            var filesList = Directory.GetFiles(dianLiangPath);
            if(filesList.Length > 0)
            {
                foreach(var file in filesList)
                {
                    if(file.Contains(date))
                    {
                        File.Copy(file, savePath.Text + "\\" + Path.GetFileName(file), true);
                    }
                }
            }

            MessageBox.Show("生成完毕", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
            return;

          


            
            //if (string.IsNullOrWhiteSpace(guangfuZuigao.Text))
            //{
            //    MessageBox.Show("请输入光伏最高负荷", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}
            //if (string.IsNullOrWhiteSpace(guangfuzuigaotime.Text))
            //{
            //    MessageBox.Show("请输入光伏最高负荷出现时间", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //if (string.IsNullOrWhiteSpace(quanwangzuigaofuhe.Text))
            //{
            //    MessageBox.Show("请输入全网负荷峰值", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //if (string.IsNullOrWhiteSpace(quanwangzuigaotime.Text))
            //{
            //    MessageBox.Show("请输入全网负荷峰值出现时间", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //if (string.IsNullOrWhiteSpace(jiaxingzuigaofuhe.Text))
            //{
            //    MessageBox.Show("请输入嘉兴负荷峰值", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //if (string.IsNullOrWhiteSpace(jiaxingzuigaotime.Text))
            //{
            //    MessageBox.Show("请输入嘉兴负荷峰值出现时间", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //if (string.IsNullOrWhiteSpace(guanfuInput.Text))
            //{
            //    MessageBox.Show("请输入光伏电量", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //if (string.IsNullOrWhiteSpace(e5000Path.Text))
            //{
            //    MessageBox.Show("请选择E5000文件", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //if (string.IsNullOrWhiteSpace(e3000Path.Text))
            //{
            //    MessageBox.Show("请选择E3000文件", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}

            //var excelop = new ExcelOpService();
            ////E5000表格读取
            //var e5000ExtName = Path.GetExtension(e5000Path.Text);//判断EXCEL文件格式
            //ISheet e5000Sheet;
            //if (string.IsNullOrEmpty(e5000ExtName) || (e5000ExtName.ToUpper() != ".XLS" && e5000ExtName.ToUpper() != ".XLSX"))
            //{
            //    MessageBox.Show("输入的文件格式不正确", "错误信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}
            //else if (e5000ExtName.ToUpper() == ".XLS")
            //{
            //    using (var file = new FileStream(e5000Path.Text, FileMode.Open, FileAccess.Read))
            //    {
            //        var e5000Book = new HSSFWorkbook(file);
            //        e5000Sheet = e5000Book.GetSheetAt(0);
            //        file.Close();
            //    }
            //}
            //else
            //{
            //    using (var file = new FileStream(e5000Path.Text, FileMode.Open, FileAccess.Read))
            //    {

            //        var e5000Book = new XSSFWorkbook(file);
            //        e5000Sheet = e5000Book.GetSheetAt(0);
            //        file.Close();
            //    }
            //}
            //// var e5000Model = excelop.GetE5000Model(e5000Sheet, guanfuInput.Text);//将E5000的数据转换为对应的dto  
            ////textBox2.Text += "读取E5000数据";
            ////E3000表格读取
            //var e3000ExtName = System.IO.Path.GetExtension(e3000Path.Text);//判断EXCEL文件格式
            //ISheet e3000Sheet;

            //if (string.IsNullOrEmpty(e3000ExtName) || (e3000ExtName.ToUpper() != ".XLS" && e3000ExtName.ToUpper() != ".XLSX"))
            //{
            //    MessageBox.Show("输入的文件格式不正确", "错误信息", MessageBoxButton.OK, MessageBoxImage.Warning);
            //    return;
            //}
            //else if (e3000ExtName.ToUpper() == ".XLS")
            //{
            //    using (var file = new FileStream(e3000Path.Text, FileMode.Open, FileAccess.Read))
            //    {

            //        var e3000Book = new HSSFWorkbook(file);
            //        e3000Sheet = e3000Book.GetSheetAt(0);
            //        file.Close();
            //    }
            //}
            //else
            //{
            //    using (var file = new FileStream(e3000Path.Text, FileMode.Open, FileAccess.Read))
            //    {

            //        var e3000Book = new XSSFWorkbook(file);
            //        e3000Sheet = e3000Book.GetSheetAt(0);
            //        file.Close();
            //    }
            //}
            //var e3000Model = excelop.GetE3000Model(e3000Sheet);//将E3000的数据转换为对应的dto  
            ////textBox2.Text += "读取E3000数据";
            ////模板读取
            //var modelPath = AppDomain.CurrentDomain.BaseDirectory + @"\ExcelModel\DailyElectricityModel.xlsx";
            //XSSFWorkbook endBook = null;
            //using (var file = new FileStream(modelPath, FileMode.Open, FileAccess.Read))
            //{
            //    endBook = new XSSFWorkbook(file);
            //    file.Close();
            //}
            //var endSheet = endBook.GetSheetAt(0);
            ////textBox2.Text += "读取模板";

            ////endSheet = excelop.SetE5000Value(e5000Model, endSheet);//将E5000的数据赋值到模板中

            ////endSheet = excelop.SetE3000Value(e3000Model, endSheet, guangfuZuigao.Text, guangfuzuigaotime.Text,
            ////quanwangzuigaofuhe.Text, quanwangzuigaotime.Text, jiaxingzuigaofuhe.Text, jiaxingzuigaotime.Text);//将E3000的数据赋值到模板中
            ////textBox2.Text += "对模板进行赋值";
            //var targetPath = savePath.Text + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
            //using (FileStream fileStream = File.Create(targetPath))
            //{
            //    endBook.Write(fileStream);
            //    fileStream.Close();
            //}
            ////textBox2.Text += "文件已生成";
            //MessageBox.Show("文件已生成，请到对应目录下寻找", "提示信息", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
