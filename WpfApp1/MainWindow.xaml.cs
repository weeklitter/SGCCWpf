using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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

        private ElectricityDailyPage electricityDailyPage;

        private SafeMeetingPage safeMeetingPage;

        private void Btn_Page1_Click(object sender, RoutedEventArgs e)
        {
            if (electricityDailyPage == null)
            {
                Console.WriteLine("first Page1");
                electricityDailyPage = new ElectricityDailyPage();
            }
            Console.WriteLine(" Page1");
            PowerPage.Content = new Frame()
            {

                Content = electricityDailyPage
            };
        }

        private void Btn_Page2_Click(object sender, RoutedEventArgs e)
        {
            if (safeMeetingPage == null)
            {
                safeMeetingPage = new SafeMeetingPage();
            }

            PowerPage.Content = new Frame()
            {
                Content = safeMeetingPage
            };
        }
    }
}
