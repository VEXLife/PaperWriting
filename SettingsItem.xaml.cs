using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace PaperWriting
{
    /// <summary>
    /// SettingsItem.xaml 的交互逻辑
    /// </summary>
    public partial class SettingsItem : UserControl
    {
        public SettingsItem()
        {
            InitializeComponent();
            this.DataContext = this;
        }

        public static readonly DependencyProperty SettingItemValueProperty = DependencyProperty.Register("SettingItemValue", typeof(object), typeof(UserControl));
        public string SettingItemName { get; set; }
        public string SettingItemTip { get; set; }
        public object SettingItemValue
        {
            get
            {
                return GetValue(SettingItemValueProperty);
            }
            set
            {
                SetValue(SettingItemValueProperty, value);
            }
        }
        public SettingItemTypeEnum SettingItemType { get; set; } = SettingItemTypeEnum.Text;
        public dynamic SettingItemOptions { get; set; }
    }

    public enum SettingItemTypeEnum
    {
        RichText, Text, Option
    }

    public class SettingItemSelector : DataTemplateSelector
    {
        public DataTemplate RichTextTmpl { get; set; }
        public DataTemplate TextTmpl { get; set; }
        public DataTemplate OptionTmpl { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item == null) return null;
            switch ((SettingItemTypeEnum)item)
            {
                case SettingItemTypeEnum.RichText:
                    return RichTextTmpl;
                case SettingItemTypeEnum.Text:
                    return TextTmpl;
                case SettingItemTypeEnum.Option:
                    return OptionTmpl;
                default:
                    return base.SelectTemplate(item, container);
            }
        }
    }

    public class SettingOptionValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch (value.ToString())
            {
                case "Above":
                    return "上方";
                case "Below":
                    return "下方";
                default:
                    return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            switch (value)
            {
                case "上方":
                    return TargetPosition.Above;
                case "下方":
                    return TargetPosition.Below;
                default:
                    return DependencyProperty.UnsetValue;
            }
        }
    }
}
