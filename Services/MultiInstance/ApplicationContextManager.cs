using Microsoft.Office.Core;
using NLog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services
{
    /// <summary>
    /// 複数PowerPointインスタンス対応：アプリケーションコンテキスト管理
    /// 商用レベルの複数ウィンドウ対応を実現
    /// </summary>
    public class ApplicationContextManager : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        private readonly object lockObject = new object();
        private readonly Dictionary<int, PowerPoint.Application> applicationCache;
        private PowerPoint.Application currentApplication;
        private IntPtr lastActiveWindow = IntPtr.Zero;
        private System.Threading.Timer windowCheckTimer;
        private bool disposed = false;

        // Win32 API宣言
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        public ApplicationContextManager()
        {
            applicationCache = new Dictionary<int, PowerPoint.Application>();

            // フォールバック：最初は現在のアプリケーションを使用
            currentApplication = Globals.ThisAddIn.Application;

            // 5秒間隔でアクティブウィンドウをチェック（商用レベルのレスポンス性）
            windowCheckTimer = new System.Threading.Timer(CheckActiveWindow, null,
                TimeSpan.FromSeconds(1), TimeSpan.FromSeconds(5));

            logger.Info("ApplicationContextManager initialized for commercial multi-instance support");
        }

        /// <summary>
        /// 現在アクティブなPowerPointアプリケーションを取得（商用メイン機能）
        /// </summary>
        public PowerPoint.Application CurrentApplication
        {
            get
            {
                lock (lockObject)
                {
                    // アクティブアプリケーションの有効性チェック
                    if (!IsApplicationValid(currentApplication))
                    {
                        logger.Debug("Current application invalid, searching for active instance");
                        RefreshCurrentApplication();
                    }

                    return currentApplication ?? Globals.ThisAddIn.Application;
                }
            }
        }

        /// <summary>
        /// アクティブウィンドウ変更チェック（バックグラウンド監視）
        /// </summary>
        private void CheckActiveWindow(object state)
        {
            try
            {
                var foregroundWindow = GetForegroundWindow();

                // 前回と同じウィンドウなら処理スキップ
                if (foregroundWindow == lastActiveWindow)
                    return;

                if (IsPowerPointWindow(foregroundWindow))
                {
                    var processId = GetWindowProcessId(foregroundWindow);
                    if (processId > 0)
                    {
                        UpdateCurrentApplicationByProcess((int)processId);
                        lastActiveWindow = foregroundWindow;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Error in active window check (non-critical)");
            }
        }

        /// <summary>
        /// プロセスIDからPowerPointウィンドウかどうかを判定
        /// </summary>
        private bool IsPowerPointWindow(IntPtr hwnd)
        {
            try
            {
                var processId = GetWindowProcessId(hwnd);
                if (processId == 0) return false;

                var process = Process.GetProcessById((int)processId);
                return string.Equals(process.ProcessName, "POWERPNT", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// ウィンドウハンドルからプロセスIDを取得
        /// </summary>
        private uint GetWindowProcessId(IntPtr hwnd)
        {
            GetWindowThreadProcessId(hwnd, out uint processId);
            return processId;
        }

        /// <summary>
        /// プロセスIDに基づいて現在のアプリケーションを更新
        /// </summary>
        private void UpdateCurrentApplicationByProcess(int processId)
        {
            lock (lockObject)
            {
                try
                {
                    // キャッシュから取得
                    if (applicationCache.ContainsKey(processId))
                    {
                        var cachedApp = applicationCache[processId];
                        if (IsApplicationValid(cachedApp))
                        {
                            if (currentApplication != cachedApp)
                            {
                                currentApplication = cachedApp;
                                logger.Debug($"Switched to cached application (PID: {processId})");
                            }
                            return;
                        }
                        else
                        {
                            applicationCache.Remove(processId);
                        }
                    }

                    // 新しいアプリケーションインスタンスを検索・登録
                    var newApp = FindApplicationByProcess(processId);
                    if (newApp != null)
                    {
                        applicationCache[processId] = newApp;
                        currentApplication = newApp;
                        logger.Info($"Registered and switched to new application (PID: {processId})");
                    }
                }
                catch (Exception ex)
                {
                    logger.Warn(ex, $"Failed to update application for process {processId}");
                }
            }
        }

        /// <summary>
        /// プロセスIDからPowerPointアプリケーションを検索
        /// </summary>
        private PowerPoint.Application FindApplicationByProcess(int targetProcessId)
        {
            try
            {
                // 現在のアプリケーションが該当プロセスかチェック
                var currentProcess = Process.GetCurrentProcess();
                if (targetProcessId == currentProcess.Id)
                {
                    return Globals.ThisAddIn.Application;
                }

                // ROT (Running Object Table) からPowerPointアプリケーションを検索
                return GetApplicationFromROT(targetProcessId);
            }
            catch (Exception ex)
            {
                logger.Debug(ex, $"Failed to find application for process {targetProcessId}");
                return null;
            }
        }

        /// <summary>
        /// ROTからPowerPointアプリケーションを取得（商用レベル実装）
        /// </summary>
        private PowerPoint.Application GetApplicationFromROT(int processId)
        {
            try
            {
                // COM相互運用による安全なROTアクセス
                var rotTable = GetRunningObjectTable(0, out IRunningObjectTable rot);
                if (rotTable != 0) return null;

                rot.EnumRunning(out IEnumMoniker enumMoniker);
                enumMoniker.Reset();

                var monikers = new IMoniker[1];
                IntPtr numFetched = IntPtr.Zero;

                while (enumMoniker.Next(1, monikers, numFetched) == 0)
                {
                    var moniker = monikers[0];
                    if (moniker == null) continue;

                    try
                    {
                        rot.GetObject(moniker, out object obj);
                        if (obj is PowerPoint.Application app)
                        {
                            // プロセスIDの一致確認（簡易版）
                            var appProcess = GetApplicationProcess(app);
                            if (appProcess?.Id == processId)
                            {
                                return app;
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(moniker);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "ROT access failed, using fallback method");
                return null;
            }
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        /// <summary>
        /// PowerPointアプリケーションからプロセス情報を取得
        /// </summary>
        private Process GetApplicationProcess(PowerPoint.Application app)
        {
            try
            {
                // PowerPointアプリケーションのウィンドウハンドルからプロセスID取得
                if (app.ActiveWindow != null)
                {
                    var hwnd = (IntPtr)app.ActiveWindow.HWND;
                    var processId = GetWindowProcessId(hwnd);
                    if (processId > 0)
                    {
                        return Process.GetProcessById((int)processId);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Failed to get application process");
            }

            return null;
        }

        /// <summary>
        /// PowerPointアプリケーションの有効性チェック
        /// </summary>
        private bool IsApplicationValid(PowerPoint.Application app)
        {
            if (app == null) return false;

            try
            {
                // プロパティアクセスで有効性確認
                var _ = app.Version;
                return true;
            }
            catch (COMException)
            {
                return false;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Application validity check failed");
                return false;
            }
        }

        /// <summary>
        /// 現在のアプリケーション参照を更新
        /// </summary>
        private void RefreshCurrentApplication()
        {
            try
            {
                // アクティブなPowerPointプロセスを検索
                var powerPointProcesses = Process.GetProcessesByName("POWERPNT");

                foreach (var process in powerPointProcesses)
                {
                    var app = FindApplicationByProcess(process.Id);
                    if (app != null && IsApplicationValid(app))
                    {
                        currentApplication = app;
                        applicationCache[process.Id] = app;
                        logger.Debug($"Refreshed current application to PID {process.Id}");
                        return;
                    }
                }

                // フォールバック：元のアプリケーション
                currentApplication = Globals.ThisAddIn.Application;
                logger.Debug("Fallback to original application");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to refresh current application");
                currentApplication = Globals.ThisAddIn.Application;
            }
        }

        /// <summary>
        /// 手動でアプリケーションコンテキストを切り替え（高度な使用用途）
        /// </summary>
        public void SwitchToApplication(PowerPoint.Application application)
        {
            if (application != null && IsApplicationValid(application))
            {
                lock (lockObject)
                {
                    currentApplication = application;
                    logger.Info("Manual application context switch performed");
                }
            }
        }

        /// <summary>
        /// デバッグ情報取得（トラブルシューティング用）
        /// </summary>
        public string GetDebugInfo()
        {
            try
            {
                var info = new System.Text.StringBuilder();
                info.AppendLine("=== ApplicationContextManager Debug Info ===");

                var processes = Process.GetProcessesByName("POWERPNT");
                info.AppendLine($"PowerPoint Processes: {processes.Length}");

                lock (lockObject)
                {
                    info.AppendLine($"Cached Applications: {applicationCache.Count}");
                    info.AppendLine($"Current Application Valid: {IsApplicationValid(currentApplication)}");

                    if (currentApplication != null)
                    {
                        try
                        {
                            info.AppendLine($"Current Application Version: {currentApplication.Version}");
                            info.AppendLine($"Current Presentations: {currentApplication.Presentations.Count}");
                        }
                        catch (Exception ex)
                        {
                            info.AppendLine($"Current Application Error: {ex.Message}");
                        }
                    }
                }

                return info.ToString();
            }
            catch (Exception ex)
            {
                return $"Debug info error: {ex.Message}";
            }
        }

        public void Dispose()
        {
            if (disposed) return;

            try
            {
                windowCheckTimer?.Dispose();

                lock (lockObject)
                {
                    // キャッシュクリア（COMオブジェクトは自動GC）
                    applicationCache.Clear();
                    currentApplication = null;
                }

                logger.Info("ApplicationContextManager disposed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error disposing ApplicationContextManager");
            }

            disposed = true;
        }
    }
}