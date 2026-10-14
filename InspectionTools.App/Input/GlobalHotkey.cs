using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Interop;

namespace InspectionTools.App.Input;

/// <summary>
/// 全体のホットキー（他のソフトにフォーカスがあっても効くキー）。Windows の RegisterHotKey で登録し、押されたら Pressed を起こす。
/// 登録したキーはシステム全体で奪われ、他のソフトには届かなくなる
/// </summary>
public sealed class GlobalHotkey : IDisposable {

    private const int HotkeyId = 1;
    private const int WmHotkey = 0x0312;
    private const uint ModAlt = 0x1;
    private const uint ModControl = 0x2;
    private const uint ModShift = 0x4;
    private const uint ModWin = 0x8;
    private const uint ModNoRepeat = 0x4000;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private HwndSource? _source;
    private IntPtr _windowHandle;
    private bool _isRegistered;

    /// <summary>いま登録されているか</summary>
    public bool IsRegistered => _isRegistered;

    /// <summary>ホットキーが押された</summary>
    public event Action? Pressed;

    /// <summary>キーを Windows の修飾キーの値と仮想キーコードに直す（直せないキーなら false）</summary>
    public static bool TryConvert(KeyChord chord, out uint modifiers, out uint virtualKey) {
        modifiers = 0;
        if (chord.Modifiers.HasFlag(ModifierKeys.Alt)) {
            modifiers |= ModAlt;
        }
        if (chord.Modifiers.HasFlag(ModifierKeys.Control)) {
            modifiers |= ModControl;
        }
        if (chord.Modifiers.HasFlag(ModifierKeys.Shift)) {
            modifiers |= ModShift;
        }
        if (chord.Modifiers.HasFlag(ModifierKeys.Windows)) {
            modifiers |= ModWin;
        }

        virtualKey = (uint)KeyInterop.VirtualKeyFromKey(chord.Key);
        return virtualKey != 0;
    }

    /// <summary>ウィンドウにホットキーを登録する。登録できないときは、理由を error に入れて false を返す（他のソフトが同じキーを使っているときなど）</summary>
    public bool TryRegister(IntPtr windowHandle, KeyChord chord, out string error) {
        if (_isRegistered) {
            throw new InvalidOperationException("すでに登録されています。");
        }
        if (!TryConvert(chord, out var modifiers, out var virtualKey)) {
            error = $"キー '{chord}' は、全体のホットキーにできません。";
            return false;
        }

        // 押しっぱなしで繰り返し呼ばれないように、MOD_NOREPEAT を付ける
        if (!RegisterHotKey(windowHandle, HotkeyId, modifiers | ModNoRepeat, virtualKey)) {
            error = $"キー '{chord}' を全体のホットキーに登録できませんでした（他のソフトが同じキーを使っている可能性があります）: {new Win32Exception(Marshal.GetLastWin32Error()).Message}";
            return false;
        }

        _isRegistered = true;
        _source = HwndSource.FromHwnd(windowHandle);
        _source?.AddHook(WndProc);
        _windowHandle = windowHandle;
        error = "";
        return true;
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) {
        if (msg == WmHotkey && wParam.ToInt32() == HotkeyId) {
            handled = true;
            Pressed?.Invoke();
        }
        return IntPtr.Zero;
    }

    /// <summary>登録を解除する（他のソフトで、そのキーが打てるようになる）。登録していなければ何もしない</summary>
    public void Unregister() {
        if (!_isRegistered) {
            return;
        }
        _source?.RemoveHook(WndProc);
        UnregisterHotKey(_windowHandle, HotkeyId);
        _isRegistered = false;
    }

    public void Dispose() => Unregister();
}
