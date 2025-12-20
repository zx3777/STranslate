using ScreenGrab;
using System.Drawing;

namespace STranslate.Core;

public class Screenshot : IScreenshot
{
    private readonly Settings _settings;

    public Screenshot(Settings settings)
    {
        _settings = settings;
    }

    public Bitmap? GetScreenshot()
    {
        if (ScreenGrabber.IsCapturing)
            return default;
        var bitmap = ScreenGrabber.CaptureDialog(isAuxiliary: _settings.ShowScreenshotAuxiliaryLines);
        if (bitmap == null)
            return default;
        return bitmap;
    }

    public async Task<Bitmap?> GetScreenshotAsync()
    {
        if (ScreenGrabber.IsCapturing)
            return default;
        var bitmap = await ScreenGrabber.CaptureAsync(isAuxiliary: _settings.ShowScreenshotAuxiliaryLines);
        if (bitmap == null)
            return default;
        return bitmap;
    }
}
