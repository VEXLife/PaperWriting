using System.Diagnostics;
using System.Text;
using System.Windows.Controls;

namespace PaperWriting
{
    /// <summary>
    /// About.xaml 的交互逻辑
    /// </summary>
    public partial class About : Page
    {
        public About()
        {
            InitializeComponent();
            var version_text = new StringBuilder()
                .Append("加载项版本：")
                .AppendLine(System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString())
                .Append("Microsoft Word产品版本：")
                .Append(System.Windows.Forms.Application.ProductVersion);
            this.versionLabel.Text = version_text.ToString();
        }

        private void To_Icons8(object sender, System.Windows.RoutedEventArgs e)
        {
            Process.Start("https://icons8.com");
        }

        private void To_Feathers(object sender, System.Windows.RoutedEventArgs e)
        {
            Process.Start("https://feathersicon.com");
        }
    }
}
