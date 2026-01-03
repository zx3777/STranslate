using STranslate.ViewModels;
using STranslate.Core;
using System.Windows;
using System.Windows.Threading;

namespace STranslate.Views;

public partial class MouseHookIconWindow : Window
{
    private string _currentText = "";
    private readonly DispatcherTimer _hideTimer;

    public MouseHookIconWindow()
    {
        InitializeComponent();

        // 初始化自动隐藏定时器（例如3秒后自动消失）
        _hideTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(3)
        };
        _hideTimer.Tick += (s, e) => HideWindow();

        // 鼠标移入暂停计时，移出重新计时
        this.MouseEnter += (s, e) => _hideTimer.Stop();
        this.MouseLeave += (s, e) => _hideTimer.Start();
    }

    public void ShowAt(Point point, string text)
    {
        _currentText = text;
        
        // 设置位置（在鼠标右下方一点）
        this.Left = point.X + 10;
        this.Top = point.Y + 10;

        // 确保在屏幕内
        var screenWidth = SystemParameters.VirtualScreenWidth;
        var screenHeight = SystemParameters.VirtualScreenHeight;
        if (this.Left + this.Width > screenWidth) this.Left = point.X - this.Width - 10;
        if (this.Top + this.Height > screenHeight) this.Top = point.Y - this.Height - 10;

        this.Show();
        this.Opacity = 1;
        _hideTimer.Start();
        
        // 播放淡入动画
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
        // 执行翻译
        if (DataContext is MainWindowViewModel vm)
        {
            vm.ExecuteTranslate(Utilities.LinebreakHandler(_currentText, vm.Settings.LineBreakHandleType));
        }
    }
}
