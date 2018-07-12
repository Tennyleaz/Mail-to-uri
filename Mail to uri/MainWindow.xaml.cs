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

namespace Mail_to_uri
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            tbContent.Text += "\n第二行...";
        }

        private void btnSend_Click(object sender, RoutedEventArgs e)
        {
            /// mailto:manish@simplygraphix.com?
            /// subject=Feedback for webdevelopersnotes.com
            /// &body=The Tips and Tricks section is great 
            /// &cc=anotheremailaddress@anotherdomain.com
            /// &bcc = onemore@anotherdomain.com
            ///           

            string strAddress = tbEmail.Text;
            string strSubject = tbSubject.Text;
            string strBody = tbContent.Text;
            string strCC = tbCC.Text;
            string strBCC = tbBCC.Text;

            char[] separators = { ';' };
            List<string> vAddress= new List<string>();            
            vAddress.AddRange(strAddress.Split(separators, StringSplitOptions.RemoveEmptyEntries));

            List<string> vCC = new List<string>();
            vCC.AddRange(strCC.Split(separators, StringSplitOptions.RemoveEmptyEntries));

            List<string> vBCC = new List<string>();
            vBCC.AddRange(strBCC.Split(separators, StringSplitOptions.RemoveEmptyEntries));

            string strResult = MailToUtility.SendMailToURI(vAddress, vCC, vBCC, strSubject, strBody);
            if (!string.IsNullOrEmpty(strResult))
                MessageBox.Show(strResult);
        }

        private void btnTest_Click(object sender, RoutedEventArgs e)
        {
            /// Windows 7：HKEY_CLASSES_ROOT\mailto\ 可能不存在，IE需要設定Google工具列或是類似的工具，才能開啟mailto:連結
            /// Windows 8：HKEY_CLASSES_ROOT\mailto\ 會存在，但是下面不一定有東西
            /// Windows 10：HKEY_CLASSES_ROOT\mailto\shell\open\command\ 會存在，預設值可能是 C:\windows\system32\rundll32.exe

            bool b = MailToUtility.IsMailClientInstalled();
            if (!b)
            {
                b = MailToUtility.TestOpenCommand();
                if (!b)
                    MessageBox.Show("no email client detected");
                else
                    MessageBox.Show("email client detected!");
            }
            else
                MessageBox.Show("email client detected!");
        }
    }
}
