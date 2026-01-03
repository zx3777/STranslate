using STranslate.ViewModels;
using STranslate.Core;
using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace STranslate.Views;

public partial class MouseHookIconWindow : Window
{
    private string _currentText = "";
    private readonly DispatcherTimer _hideTimer;

    public MouseHookIconWindow()
    {
        InitializeComponent();

        // 初始化自动隐藏定时器（3秒后自动消失）
        _hideTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        _hideTimer.Tick += (s, e) => HideWindow();

        // 鼠标移入暂停计时，移出重新计时，防止用户想点的时候图标消失
        this.MouseEnter += (s, e) => _hideTimer.Stop();
        this.MouseLeave += (s, e) => _hideTimer.Start();
    }

    /// <summary>
    /// 在指定位置显示图标
    /// </summary>
    /// <param name="point">鼠标的物理屏幕坐标</param>
    /// <param name="text">选中的文本</param>
    public void ShowAt(Point point, string text)
    {
        _currentText = text;

        // ★★★ 核心修复：获取当前屏幕的 DPI 缩放比例 ★★★
        var dpiScale = VisualTreeHelper.GetDpi(this);
        // 如果当前窗口未显示，可能获取不到 DPI，尝试获取主窗口的 DPI 作为后备
        if (dpiScale.PixelsPerDip == 1.0 && Application.Current.MainWindow != null)
        {
            dpiScale = VisualTreeHelper.GetDpi(Application.Current.MainWindow);
        }

        // ★★★ 核心修复：将物理像素(Pixel)转换为逻辑像素(DIP) ★★★
        // 公式：逻辑坐标 = 物理坐标 / DPI缩放比例
        // 并添加 (+10, +10) 的偏移量，避免图标直接出现在鼠标尖端遮挡点击
        this.Left = (point.X / dpiScale.DpiScaleX) + 10;
        this.Top = (point.Y / dpiScale.DpiScaleY) + 10;

        // --- 边界检查 (防止图标跑出屏幕外) ---
        var screenWidth = SystemParameters.VirtualScreenWidth;
        var screenHeight = SystemParameters.VirtualScreenHeight;
        
        // 如果超出右边界，改显示在鼠标左侧
        if (this.Left + this.Width > screenWidth) 
            this.Left = (point.X / dpiScale.DpiScaleX) - this.Width - 10;
            
        // 如果超出下边界，改显示在鼠标上方
        if (this.Top + this.Height > screenHeight) 
            this.Top = (point.Y / dpiScale.DpiScaleY) - this.Height - 10;

        this.Show();
        this.Opacity = 1;
        _hideTimer.Start();
        
        // 播放淡入动画 (需要在 xaml 中定义 key 为 FadeIn 的动画)
        (this.Resources["FadeIn"] as System.Windows.Media.Animation.Storyboard)?.Begin();
    }

    private void HideWindow()
    {
        _hideTimer.Stop();
        this.Hide();
    }

    private void TranslateBtn_Click(object sender, RoutedEventArgs e)
    {
        HideWindow();
        
        // 通过 DataContext 获取 ViewModel 来执行翻译命令
        // DataContext 是在 MainWindowViewModel 初始化时赋值的
        if (DataContext is MainWindowViewModel vm)
        {
            vm.ExecuteTranslate(Utilities.LinebreakHandler(_currentText, vm.Settings.LineBreakHandleType));
        }
    }
}
