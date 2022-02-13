using System;
using System.Collections.ObjectModel;
using System.Windows.Controls;

namespace PaperWriting
{

    /// <summary>
    /// MainSettings.xaml 的交互逻辑。
    /// </summary>
    public partial class MainSettings : Page
    {
        public MainSettings()
        {
            InitializeComponent();
        }

        ~MainSettings()
        {
            Properties.Settings.Default.Save();
        }
    }

    /// <summary>
    /// 下拉菜单项。
    /// </summary>
    class TargetPositionComboBoxItems : ObservableCollection<string>
    {
        public TargetPositionComboBoxItems()
        {
            Add("上方");
            Add("下方");
        }
    }

    /// <summary>
    /// 描述插入方位的枚举。
    /// </summary>
    [Serializable]
    public enum TargetPosition
    {
        Above,Below
    }
}
