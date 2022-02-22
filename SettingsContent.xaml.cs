using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace PaperWriting
{
    /// <summary>
    /// SettingsContent.xaml 的交互逻辑
    /// </summary>
    public partial class SettingsContent : UserControl
    {
        private readonly Point originalPoint = new Point(0, 0);
        private readonly DoubleAnimationUsingKeyFrames animation = new DoubleAnimationUsingKeyFrames();
        public Dictionary<string, object> Pages { get; } = new Dictionary<string, object>()
        {
            {"settings", new MainSettings() },
            {"info", new About() }
        }; // 页面切换的内容字典

        public SettingsContent()
        {
            InitializeComponent();
            animation.KeyFrames.Add(new EasingDoubleKeyFrame()
            {
                KeyTime = KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.15)),
                Value = 1,
                EasingFunction = new CubicEase()
                {
                    EasingMode = EasingMode.EaseInOut
                }
            }); // 避免反复初始化新的关键帧
            NavBar_Selected(settings, null);
        }

        private void menu_btn_LostFocus(object sender, RoutedEventArgs e)
        {
            menu_btn.IsChecked = false;
        }

        private void NavBar_Selected(object sender, RoutedEventArgs e)
        {
            // 实现导航栏长方形移动的动画
            var senderControl = sender as ListBoxItem;
            animation.KeyFrames[0].Value = senderControl.TranslatePoint(originalPoint, nav_btn_container).Y;
            nav_rect_trans.BeginAnimation(TranslateTransform.YProperty, animation);
            ContentFrame.Content = Pages[senderControl.Name];
        }
    }
}
