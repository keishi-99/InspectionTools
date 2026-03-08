using System.Runtime.InteropServices;
using VisaComLib;

namespace InspectionTools.Common {
    internal static partial class NativeMethods {
        [LibraryImport("ausb.dll", EntryPoint = "ausb_start")]
        public static partial int start(uint dwTmout);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_open")]
        public static partial int open(ref uint hDev, uint dwMyid);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_write")]
        public static partial int ausbWrite32(uint hDev, uint pWrtBuffer, uint dwCount);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_write")]
        public static partial int ausbWrite64(uint hDev, ulong pWrtBuffer, uint dwCount);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_read")]
        public static partial int ausbRead32(uint hDev, uint pRdBuffer, uint dwCount, ref uint pRdCnt);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_read")]
        public static partial int ausbRead64(uint hDev, ulong pRdBuffer, uint dwCount, ref uint pRdCnt);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_close")]
        public static partial int close(uint dDev);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_end")]
        public static partial int end();
        [LibraryImport("ausb.dll", EntryPoint = "ausb_clear")]
        public static partial int clear(uint hDev);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_trigger")]
        public static partial int trigger(uint hDev);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_readstb")]
        public static partial int readstb(uint hDev, ref int lngSTB);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_timeout")]
        public static partial int timeout(uint tmout);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_local")]
        public static partial int local(uint hDev);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_llo")]
        public static partial int llo(uint hDev);
        [LibraryImport("ausb.dll", EntryPoint = "ausb_reset")]
        public static partial int reset(uint hDev);

        // コマンド文字列をASCIIバイト列に変換してUSBデバイスに書き込む
        public static int Write(uint hDev, string strCmd) {
            ulong pBuf;

            byte[] pBuffer = System.Text.Encoding.ASCII.GetBytes(strCmd);
            uint cmdCnt = (uint)pBuffer.Length;
            Array.Resize(ref pBuffer, pBuffer.Length + 1); // null terminator

            var ret = 0;
            unsafe {
                fixed (byte* p = &pBuffer[0]) {
                    pBuf = (ulong)p;
                    ret = IntPtr.Size == 4
                        ? ausbWrite32(hDev, (uint)pBuf, cmdCnt)
                        : ausbWrite64(hDev, pBuf, cmdCnt);
                }
            }
            return ret;
        }

        public static int Read(uint hDev, ref string readDt, ref uint rdCnt, uint lngCnt = 256) {
            int ret;
            ulong pBuf;
            byte[] pBuffer;
            pBuffer = new byte[lngCnt + 1];
            unsafe {
                fixed (byte* p = &pBuffer[0]) {
                    pBuf = (ulong)p;
                    ret = IntPtr.Size == 4 ? ausbRead32(hDev, (uint)pBuf, lngCnt, ref rdCnt) : ausbRead64(hDev, pBuf, lngCnt, ref rdCnt);

                    readDt = "";
                    if (ret == 0) {
                        var tmps = System.Text.Encoding.Default.GetString(pBuffer);
                        readDt = tmps.TrimEnd(['\r', '\n', '\0']);
                    }
                }
            }
            return ret;
        }
    }

    public static class AusbWrapper {
        // USBドライバを初期化する
        public static int Start(uint dwTmout) {
            return NativeMethods.start(dwTmout);
        }
        // 指定IDのUSBデバイスをオープンする
        public static int Open(ref uint hDev, uint dwMyid) {
            return NativeMethods.open(ref hDev, dwMyid);
        }
        // コマンド文字列をUSBデバイスに書き込む
        public static int Write(uint hDev, string strCmd) {
            return NativeMethods.Write(hDev, strCmd);
        }
        // USBデバイスからデータを読み取る
        public static int Read(uint hDev, ref string readDt, ref uint rdCnt, uint lngCnt = 256) {
            return NativeMethods.Read(hDev, ref readDt, ref rdCnt, lngCnt);
        }
        // 指定USBデバイスをクローズする
        public static int Close(uint hDev) {
            return NativeMethods.close(hDev);
        }
        // USBドライバを終了する
        public static int End() {
            return NativeMethods.end();
        }
    }

    public class USBDeviceManager : IDisposable {
        private readonly FormattedIO488 _dev;
        private ResourceManager? _resourceManager;
        private IMessage? _io;  // IMessageの保持

        private bool _disposed = false; // Disposeが既に呼ばれたかどうかのフラグ

        public USBDeviceManager() {
            _dev = new FormattedIO488();
        }

        // VISAアドレスを指定してデバイスに接続する
        public void OpenDev(string visaaddress) {
            try {
                _resourceManager = new ResourceManager();
                _io = (IMessage)_resourceManager.Open(visaaddress, AccessMode.NO_LOCK);
                _dev.IO = _io;
                _dev.IO.Timeout = 20000;
            } catch (Exception ex) {
                throw new ApplicationException("接続中にエラーが発生しました。", ex);
            }
        }

        // デバイスにコマンドを送信する
        public void OutputDev(string cmd) {
            _dev.WriteString(cmd);
        }

        // デバイスからデータを読み取る
        public string InputDev() {
            try {
                return _dev.ReadString();
            } catch (Exception ex) {
                throw new ApplicationException("データ取得中にエラーが発生しました。", ex);
            }
        }

        // デバイスの接続をクローズする
        public void CloseDev() {
            if (_io != null) {
                try {
                    _io.Close();
                } finally {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_io);
                    _io = null;
                }
            }

            if (_resourceManager != null) {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_resourceManager);
                _resourceManager = null;
            }
        }

        // IDisposableの実装
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // マネージドリソースを解放する
        protected virtual void Dispose(bool disposing) {
            if (!_disposed) {
                if (disposing) {
                    CloseDev();
                }
                _disposed = true;
            }
        }

        // ファイナライザ
        ~USBDeviceManager() {
            Dispose(false);
        }
    }
}