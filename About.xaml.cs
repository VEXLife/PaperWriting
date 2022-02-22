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

        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            Process.Start(e.Uri.ToString());
            e.Handled = true;
        }
    }
}
