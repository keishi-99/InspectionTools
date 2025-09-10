using System.Runtime.InteropServices;
using System.Windows.Input;

namespace WinFormsLibrary {

    internal static partial class NativeMethods {

        [LibraryImport("user32.dll")]
        public static partial IntPtr GetForegroundWindow();

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static partial bool SetForegroundWindow(IntPtr hWnd);

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static partial bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static partial bool IsIconic(IntPtr hWnd); // ウィンドウが最小化されているかチェック

        // StringBuilder の代わりに char[] を使用し、Out 属性を指定します。
        // 呼び出し側で char[] を用意し、結果を string に変換する必要があります。
        [LibraryImport("user32.dll", EntryPoint = "GetWindowTextW", StringMarshalling = StringMarshalling.Utf16)]
        public static partial int GetWindowText(IntPtr hWnd, [Out] char[] lpString, int nMaxCount);

        // SendMessage の IntPtr lParam オーバーロードは変更なし
        [LibraryImport("user32.dll", EntryPoint = "SendMessageW", StringMarshalling = StringMarshalling.Utf16)]
        public static partial IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        // StringBuilder の代わりに char[] を使用し、In, Out 属性を指定します。
        // これにより、バッファとして使用される文字配列がマーシャリングされます。
        [LibraryImport("user32.dll", EntryPoint = "SendMessageW", StringMarshalling = StringMarshalling.Utf16)]
        public static partial IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, [In, Out] char[] lParam);

        // bool の戻り値に MarshalAs(UnmanagedType.Bool) を追加します。
        [LibraryImport("user32.dll", EntryPoint = "PostMessageW", StringMarshalling = StringMarshalling.Utf16)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static partial bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [LibraryImport("user32.dll", EntryPoint = "FindWindowW", StringMarshalling = StringMarshalling.Utf16)]
        public static partial IntPtr FindWindow(string? lpClassName, string lpWindowName);

        [LibraryImport("user32.dll", EntryPoint = "FindWindowExW", StringMarshalling = StringMarshalling.Utf16)]
        public static partial IntPtr FindWindowEx(IntPtr hWnd, IntPtr hwndChildafter, string? lpClassName, string lpWindowName);

        // bool の戻り値に MarshalAs(UnmanagedType.Bool) を追加します。
        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)] // 追加
        public static partial bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        // bool の戻り値に MarshalAs(UnmanagedType.Bool) を追加します。
        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)] // 追加
        public static partial bool SetFocus(IntPtr hWnd);

        [LibraryImport("user32.dll")]
        public static partial IntPtr GetFocus();

        // bool の戻り値に MarshalAs(UnmanagedType.Bool) を追加します。
        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)] // 追加
        public static partial bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        // bool の戻り値に MarshalAs(UnmanagedType.Bool) を追加します。
        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)] // 追加
        public static partial bool UnregisterHotKey(IntPtr hWnd, int id);
    }

    /// <summary>
    /// NativeMethods クラスの機能を外部のアセンブリに公開するためのパブリックなラッパークラスです。
    /// このクラスは、NativeMethods と同じアセンブリ内に配置する必要があります。
    /// </summary>
    public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);
    public static class Win32Wrapper {
        // GetForegroundWindow のラッパー
        public static IntPtr GetForegroundWindow() {
            return NativeMethods.GetForegroundWindow();
        }

        // SetForegroundWindow のラッパー
        public static bool SetForegroundWindow(IntPtr hWnd) {
            return NativeMethods.SetForegroundWindow(hWnd);
        }

        // ShowWindow のラッパー
        public static bool ShowWindow(IntPtr hWnd, int nCmdShow) {
            return NativeMethods.ShowWindow(hWnd, nCmdShow);
        }

        // IsIconic のラッパー
        public static bool IsIconic(IntPtr hWnd) {
            return NativeMethods.IsIconic(hWnd);
        }

        /// <summary>
        /// 指定されたウィンドウのテキストを取得します。
        /// </summary>
        /// <param name="hWnd">ウィンドウのハンドル。</param>
        /// <returns>ウィンドウのテキスト。取得できなかった場合は空文字列。</returns>
        public static string GetWindowText(IntPtr hWnd) {
            // バッファサイズを定義します。必要に応じて調整してください。
            const int MaxCount = 256;
            var buffer = new char[MaxCount];
            // 内部の NativeMethods を呼び出します。
            var length = NativeMethods.GetWindowText(hWnd, buffer, MaxCount);

            if (length is > 0 and < MaxCount) {
                // 取得した文字数分だけ文字列を構築します。
                return new string(buffer, 0, length);
            }
            return string.Empty;
        }

        /// <summary>
        /// 指定されたウィンドウまたはコントロールにメッセージを送信します。
        /// lParam がポインタとして扱われるメッセージ（例: WM_COMMAND, WM_LBUTTONDOWN）に使用します。
        /// </summary>
        /// <param name="hWnd">メッセージを受信するウィンドウのハンドル。</param>
        /// <param name="msg">送信するメッセージ。</param>
        /// <param name="wParam">メッセージ固有の情報。</param>
        /// <param name="lParam">メッセージ固有の情報。</param>
        /// <returns>メッセージ処理の結果。</returns>
        public static IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam) {
            return NativeMethods.SendMessage(hWnd, msg, wParam, lParam);
        }

        /// <summary>
        /// 指定されたウィンドウまたはコントロールにテキストメッセージを送信します (例: WM_SETTEXT)。
        /// </summary>
        /// <param name="hWnd">メッセージを受信するウィンドウのハンドル。</param>
        /// <param name="msg">送信するメッセージ (例: WM_SETTEXT)。</param>
        /// <param name="wParam">メッセージ固有の情報。</param>
        /// <param name="text">送信するテキスト。</param>
        /// <returns>メッセージ処理の結果。</returns>
        public static IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string text) {
            // 文字列をネイティブメモリにマーシャリングし、そのポインタを lParam として渡します。
            // Marshal.StringToHGlobalUni は Unicode (UTF-16) 文字列をネイティブメモリにコピーします。
            var ptr = Marshal.StringToHGlobalUni(text);
            try {
                return NativeMethods.SendMessage(hWnd, msg, wParam, ptr);
            } finally {
                // ネイティブメモリを解放することを忘れないでください。
                Marshal.FreeHGlobal(ptr);
            }
        }

        /// <summary>
        /// 指定されたウィンドウまたはコントロールからテキストメッセージを取得します (例: WM_GETTEXT)。
        /// </summary>
        /// <param name="hWnd">メッセージを受信するウィンドウのハンドル。</param>
        /// <param name="msg">送信するメッセージ (例: WM_GETTEXT)。</param>
        /// <param name="wParam">メッセージ固有の情報。</param>
        /// <param name="maxLen">取得するテキストの最大長 (終端NULL文字を含む)。</param>
        /// <returns>取得したテキスト。</returns>
        public static string GetMessageText(IntPtr hWnd, uint msg, int maxLen) // wParam パラメーターを削除
        {
            if (maxLen <= 0) {
                return string.Empty;
            }

            var buffer = new char[maxLen];
            // 内部の NativeMethods を呼び出します。
            // WM_GETTEXT の wParam はバッファサイズなので、maxLen を渡します。
            var result = NativeMethods.SendMessage(hWnd, msg, maxLen, buffer);

            // WM_GETTEXT の場合、戻り値はコピーされた文字数 (終端NULL文字を除く) です。
            var length = (int)result;
            return length > 0 && length < maxLen ? new string(buffer, 0, length) : string.Empty;
        }

        // PostMessage のラッパー
        public static bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam) {
            return NativeMethods.PostMessage(hWnd, msg, wParam, lParam);
        }

        // FindWindow のラッパー
        public static IntPtr FindWindow(string? lpClassName, string lpWindowName) {
            return NativeMethods.FindWindow(lpClassName, lpWindowName);
        }

        // FindWindowEx のラッパー
        public static IntPtr FindWindowEx(IntPtr hWnd, IntPtr hwndChildafter, string? lpClassName, string lpWindowName) {
            return NativeMethods.FindWindowEx(hWnd, hwndChildafter, lpClassName, lpWindowName);
        }

        // EnumChildWindows のラッパー
        public static bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam) {
            return NativeMethods.EnumChildWindows(hWndParent, lpEnumFunc, lParam);
        }

        // SetFocus のラッパー
        public static bool SetFocus(IntPtr hWnd) {
            return NativeMethods.SetFocus(hWnd);
        }

        // GetFocus のラッパー
        public static IntPtr GetFocus() {
            return NativeMethods.GetFocus();
        }

        // RegisterHotKey のラッパー
        public static bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk) {
            return NativeMethods.RegisterHotKey(hWnd, id, fsModifiers, vk);
        }

        // UnregisterHotKey のラッパー
        public static bool UnregisterHotKey(IntPtr hWnd, int id) {
            return NativeMethods.UnregisterHotKey(hWnd, id);
        }

        // ホットキーを表すクラス
        public class Hotkey(uint modifier, System.Windows.Input.Key key, Action action) {
            private static int s_nextId = 1;

            public int Id { get; } = s_nextId++;
            public uint Modifier { get; set; } = modifier;
            public int VirtualKey { get; set; } = KeyInterop.VirtualKeyFromKey(key); // Key -> VKコード
            public Action Action { get; set; } = action ?? throw new ArgumentNullException(nameof(action));
        }

        public const uint ModNone = 0x0;
        public const uint ModAlt = 0x1;
        public const uint ModControl = 0x2;
        public const uint ModShift = 0x4;
        public const uint ModWin = 0x8;
        public const uint WmGetText = 0x000D;
        public const uint EmSetSel = 0x00B1;
        public const uint BmClick = 0x00F5;
        public const uint WmCommand = 0x0111;
        public const uint WmLButtonDown = 0x0201;
        public const uint WmLButtonUp = 0x0202;
        public const uint WmHotKey = 0x312;
        public const uint WmClear = 0x0303;

        public const int SwRestore = 9;    // 最小化されたウィンドウを復元し、アクティブにする
        public const int SwShowNomal = 1; // 通常のサイズと位置で表示し、アクティブにする
        public const int SwShow = 5;       // ウィンドウを現在のサイズと位置で表示する

        public const Key HotkeyMinus = Key.OemMinus;              // ( - )用
        public const Key HotkeyTilde = Key.Oem7;                  // ( ^ )用
        public const Key HotkeyAtsign = Key.OemTilde;             // ( @ )用
        public const Key HotkeyBracketL = Key.OemOpenBrackets;    // ( [ )用
        public const Key HotKeyemiColon = Key.OemPlus;           // ( ; )用
        public const Key HotkeyColon = Key.OemSemicolon;          // ( : )用
        public const Key HotkeyBracketR = Key.OemCloseBrackets;   // ( ] )用
        public const Key HotkeyComma = Key.OemComma;              // ( , )用
        public const Key HotkeyPeriod = Key.OemPeriod;            // ( . )用
        public const Key HotkeySlash = Key.OemQuestion;           // ( / )用
        public const Key HotkeyBackslash = Key.OemBackslash;      // ( \ )用

        public const Key HotkeyD1 = Key.D1;      // ( 1 )用
        public const Key HotkeyD2 = Key.D2;      // ( 2 )用
        public const Key HotkeyD3 = Key.D3;      // ( 3 )用
        public const Key HotkeyD4 = Key.D4;      // ( 4 )用
        public const Key HotkeyD5 = Key.D5;      // ( 5 )用
        public const Key HotkeyD6 = Key.D6;      // ( 6 )用
        public const Key HotkeyD7 = Key.D7;      // ( 7 )用
        public const Key HotkeyD8 = Key.D8;      // ( 8 )用
        public const Key HotkeyD9 = Key.D9;      // ( 9 )用
        public const Key HotkeyD0 = Key.D0;      // ( 0 )用

        public const Key HotkeyNum0 = Key.NumPad0;     // ( num0 )用
        public const Key HotkeyNum1 = Key.NumPad1;     // ( num1 )用
        public const Key HotkeyNum2 = Key.NumPad2;     // ( num2 )用
        public const Key HotkeyNum3 = Key.NumPad3;     // ( num3 )用
        public const Key HotkeyNum4 = Key.NumPad4;     // ( num4 )用
        public const Key HotkeyNum5 = Key.NumPad5;     // ( num5 )用
        public const Key HotkeyNum6 = Key.NumPad6;     // ( num6 )用
        public const Key HotkeyNum7 = Key.NumPad7;     // ( num7 )用
        public const Key HotkeyNum8 = Key.NumPad8;     // ( num8 )用
        public const Key HotkeyNum9 = Key.NumPad9;     // ( num9 )用

        public const Key HotkeyNumAdd = Key.Add;              // ( num+ )用
        public const Key HotkeyNumSubtract = Key.Subtract;    // ( num- )用
        public const Key HotkeyNumDivide = Key.Divide;        // ( num/ )用
        public const Key HotkeyNumMultiply = Key.Multiply;    // ( num* )用
    }
}
