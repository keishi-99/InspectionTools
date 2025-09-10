using System.Runtime.InteropServices;
using VisaComLib;

namespace WinFormsLibrary {
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

        public static int Write(uint hDev, string strCmd) {
            uint cmdCnt;
            int i;
            ulong pBuf;
            Console.WriteLine(strCmd);
            cmdCnt = (uint)strCmd.Length;

            byte[] pBuffer;
            pBuffer = new byte[cmdCnt + 1];

            for (i = 0; i < cmdCnt; i++) {
                pBuffer[i] = (byte)strCmd[i];
            }
            var ret = 0;
            unsafe {
                fixed (byte* p = &pBuffer[0]) {
                    pBuf = (ulong)p;
                }

                ret = IntPtr.Size == 4 ? ausbWrite32(hDev, (uint)pBuf, cmdCnt) : ausbWrite64(hDev, pBuf, cmdCnt);
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
        public static int Start(uint dwTmout) {
            return NativeMethods.start(dwTmout);
        }
        public static int Open(ref uint hDev, uint dwMyid) {
            return NativeMethods.open(ref hDev, dwMyid);
        }
        public static int Write(uint hDev, string strCmd) {
            return NativeMethods.Write(hDev, strCmd);
        }
        public static int Read(uint hDev, ref string readDt, ref uint rdCnt, uint lngCnt = 256) {
            return NativeMethods.Read(hDev, ref readDt, ref rdCnt, lngCnt);
        }
        public static int Close(uint hDev) {
            return NativeMethods.close(hDev);
        }
        public static int End() {
            return NativeMethods.end();
        }
    }

    public class USBDeviceManager : IDisposable {
        private readonly FormattedIO488 _dev;
        private IMessage? _io;  // IMessageの保持
        public bool disposed = false; // Disposeが既に呼ばれたかどうかのフラグ

        public USBDeviceManager() {
            this._dev = new FormattedIO488();
        }

        public void OpenDev(string visaaddress) {
            try {
                var resourceManager = new ResourceManager();
                this._io = (IMessage)resourceManager.Open(visaaddress, AccessMode.NO_LOCK);
                this._dev.IO = this._io;
                this._dev.IO.Timeout = 20000;
            } catch (Exception ex) {
                throw new ApplicationException("接続中にエラーが発生しました。", ex);
            }
        }

        public void OutputDev(string cmd) {
            this._dev.WriteString(cmd);
        }

        public string InputDev() {
            try {
                return this._dev.ReadString();
            } catch (Exception ex) {
                throw new ApplicationException("データ取得中にエラーが発生しました。", ex);
            }
        }

        public void CloseDev() {
            if (this._io != null) {
                this._io.Close();
                this._io = null;
            }
        }

        // IDisposableの実装
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing) {
            if (!disposed) {
                if (disposing) {
                    // マネージドリソースの解放
                    CloseDev(); // 既にCloseDevを実装済み
                }

                // アンマネージドリソースの解放 (必要があれば)

                disposed = true;
            }
        }

        // デストラクタ (必要があれば)
        ~USBDeviceManager() {
            Dispose(false);
        }
    }
}
