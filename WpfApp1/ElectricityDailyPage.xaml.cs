using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SGCCExcelOp;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Shapes;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Path = System.IO.Path;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for ElectricityDailyPage.xaml
    /// </summary>
    public partial class ElectricityDailyPage : Page
    {
        public ElectricityDailyPage()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            textBox1.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            textBox2.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.ShowDialog();
            textBox3.Text = ofd.FileName;
            ofd.RestoreDirectory = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择生成EXCEL文件夹";
                dialog.ShowNewFolderButton = false;
                dialog.RootFolder = Environment.SpecialFolder.Desktop;
                dialog.ShowDialog();
                if (dialog.SelectedPath == string.Empty)
                {
                    textBox4.Text = "请选择目录！";
                    MessageBox.Show("请选择需要存储的位置", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                textBox4.Text = dialog.SelectedPath;

            }

            if (string.IsNullOrWhiteSpace(guangfuZuigao.Text))
            {
                MessageBox.Show("请输入光伏最高负荷", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(guangfuzuigaotime.Text))
            {
                MessageBox.Show("请输入光伏最高负荷出现时间", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(quanwangzuigaofuhe.Text))
            {
                MessageBox.Show("请输入全网负荷峰值", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(quanwangzuigaotime.Text))
            {
                MessageBox.Show("请输入全网负荷峰值出现时间", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(jiaxingzuigaofuhe.Text))
            {
                MessageBox.Show("请输入嘉兴负荷峰值", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(jiaxingzuigaotime.Text))
            {
                MessageBox.Show("请输入嘉兴负荷峰值出现时间", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(guanfuInput.Text))
            {
                MessageBox.Show("请输入光伏电量", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("请选择E5000文件", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("请选择E3000文件", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(textBox4.Text))
            {
                MessageBox.Show("请选择生成文件存储位置", "警告信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var excelop = new ExcelOpService();
            //E5000表格读取
            var e5000Path = textBox2.Text;
            var e5000ExtName = Path.GetExtension(e5000Path);//判断EXCEL文件格式
            ISheet e5000Sheet;
            if (string.IsNullOrEmpty(e5000ExtName) || (e5000ExtName.ToUpper() != ".XLS" && e5000ExtName.ToUpper() != ".XLSX"))
            {
                MessageBox.Show("输入的文件格式不正确", "错误信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (e5000ExtName.ToUpper() == ".XLS")
            {
                using (var file = new FileStream(e5000Path, FileMode.Open, FileAccess.Read))
                {

                    var e5000Book = new HSSFWorkbook(file);
                    e5000Sheet = e5000Book.GetSheetAt(0);
                    file.Close();
                }
            }
            else
            {
                using (var file = new FileStream(e5000Path, FileMode.Open, FileAccess.Read))
                {

                    var e5000Book = new XSSFWorkbook(file);
                    e5000Sheet = e5000Book.GetSheetAt(0);
                    file.Close();
                }
            }
            var e5000Model = excelop.GetE5000Model(e5000Sheet, guanfuInput.Text);//将E5000的数据转换为对应的dto  
            //textBox2.Text += "读取E5000数据";
            //E3000表格读取
            var e3000Path = textBox3.Text;
            var e3000ExtName = Path.GetExtension(e3000Path);//判断EXCEL文件格式
            ISheet e3000Sheet;

            if (string.IsNullOrEmpty(e3000ExtName) || (e3000ExtName.ToUpper() != ".XLS" && e3000ExtName.ToUpper() != ".XLSX"))
            {
                MessageBox.Show("输入的文件格式不正确", "错误信息", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            else if (e3000ExtName.ToUpper() == ".XLS")
            {
                using (var file = new FileStream(e3000Path, FileMode.Open, FileAccess.Read))
                {

                    var e3000Book = new HSSFWorkbook(file);
                    e3000Sheet = e3000Book.GetSheetAt(0);
                    file.Close();
                }
            }
            else
            {
                using (var file = new FileStream(e3000Path, FileMode.Open, FileAccess.Read))
                {

                    var e3000Book = new XSSFWorkbook(file);
                    e3000Sheet = e3000Book.GetSheetAt(0);
                    file.Close();
                }
            }
            var e3000Model = excelop.GetE3000Model(e3000Sheet);//将E3000的数据转换为对应的dto  
            //textBox2.Text += "读取E3000数据";
            //模板读取
            var modelPath = AppDomain.CurrentDomain.BaseDirectory + @"\ExcelModel\DailyElectricityModel.xlsx";
            XSSFWorkbook endBook = null;
            using (var file = new FileStream(modelPath, FileMode.Open, FileAccess.Read))
            {
                endBook = new XSSFWorkbook(file);
                file.Close();
            }
            var endSheet = endBook.GetSheetAt(0);
            //textBox2.Text += "读取模板";

            endSheet = excelop.SetE5000Value(e5000Model, endSheet);//将E5000的数据赋值到模板中

            endSheet = excelop.SetE3000Value(e3000Model, endSheet, guangfuZuigao.Text, guangfuzuigaotime.Text,
                quanwangzuigaofuhe.Text, quanwangzuigaotime.Text, jiaxingzuigaofuhe.Text, jiaxingzuigaotime.Text);//将E3000的数据赋值到模板中
            //textBox2.Text += "对模板进行赋值";
            var targetPath = textBox4.Text + "\\" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
            using (FileStream fileStream = File.Create(targetPath))
            {
                endBook.Write(fileStream);
                fileStream.Close();
            }
            //textBox2.Text += "文件已生成";
            MessageBox.Show("文件已生成，请到对应目录下寻找", "提示信息", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
