using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using NLog;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Utils
{
    /// <summary>
    /// COMオブジェクトの安全な解放とUNDO境界管理を行うヘルパークラス
    /// </summary>
    public static class ComHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// COMオブジェクトを安全に解放します
        /// </summary>
        /// <param name="comObject">解放するCOMオブジェクト</param>
        public static void ReleaseComObject(object comObject)
        {
            if (comObject == null) return;

            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    int refCount = Marshal.ReleaseComObject(comObject);
                    logger.Debug($"COM object released. Reference count: {refCount}");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to release COM object");
            }
        }

        /// <summary>
        /// 複数のCOMオブジェクトを一括で解放します
        /// </summary>
        /// <param name="comObjects">解放するCOMオブジェクトの配列</param>
        public static void ReleaseComObjects(params object[] comObjects)
        {
            if (comObjects == null) return;

            foreach (var obj in comObjects)
            {
                ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// 🆕 UNDO境界管理付きでCOMオブジェクトを使用する処理を安全に実行します
        /// </summary>
        /// <typeparam name="T">戻り値の型</typeparam>
        /// <param name="action">実行する処理</param>
        /// <param name="undoEntryName">UNDO操作名</param>
        /// <param name="comObjects">処理後に解放するCOMオブジェクト</param>
        /// <returns>処理結果</returns>
        public static T ExecuteWithComCleanup<T>(Func<T> action, string undoEntryName = null, params object[] comObjects)
        {
            try
            {
                // 🆕 UNDO境界を開始
                if (!string.IsNullOrEmpty(undoEntryName))
                {
                    StartNewUndoEntry(undoEntryName);
                }

                return action();
            }
            finally
            {
                ReleaseComObjects(comObjects);
            }
        }

        /// <summary>
        /// 🆕 UNDO境界管理付きでCOMオブジェクトを使用する処理を安全に実行します（戻り値なし）
        /// </summary>
        /// <param name="action">実行する処理</param>
        /// <param name="undoEntryName">UNDO操作名</param>
        /// <param name="comObjects">処理後に解放するCOMオブジェクト</param>
        public static void ExecuteWithComCleanup(Action action, string undoEntryName = null, params object[] comObjects)
        {
            try
            {
                // 🆕 UNDO境界を開始
                if (!string.IsNullOrEmpty(undoEntryName))
                {
                    StartNewUndoEntry(undoEntryName);
                }

                action();
            }
            finally
            {
                ReleaseComObjects(comObjects);
            }
        }

        /// <summary>
        /// 🆕 従来版（後方互換性維持）- UNDO境界なし
        /// </summary>
        /// <param name="action">実行する処理</param>
        /// <param name="comObjects">処理後に解放するCOMオブジェクト</param>
        public static void ExecuteWithComCleanup(Action action, params object[] comObjects)
        {
            ExecuteWithComCleanup(action, undoEntryName: null, comObjects);
        }

        /// <summary>
        /// 🆕 新しいUNDOエントリを開始します
        /// </summary>
        /// <param name="undoEntryName">UNDO操作の名前（ログ用）</param>
        public static void StartNewUndoEntry(string undoEntryName)
        {
            try
            {
                // 🆕 複数インスタンス対応：現在アクティブなアプリケーション取得
                var application = GetCurrentActiveApplication();
                if (application != null)
                {
                    application.StartNewUndoEntry();
                    logger.Debug($"Started new undo entry for multi-instance: {undoEntryName}");
                }
                else
                {
                    logger.Warn("No active PowerPoint application found for undo entry");
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to start undo entry: {undoEntryName}");
            }
        }

        /// <summary>
        /// 🆕 現在アクティブなPowerPointアプリケーション取得
        /// </summary>
        private static PowerPoint.Application GetCurrentActiveApplication()
        {
            try
            {
                // ApplicationContextManagerが利用可能な場合はそれを使用
                var contextManager = Globals.ThisAddIn.ApplicationContextManager;
                if (contextManager != null)
                {
                    return contextManager.CurrentApplication;
                }

                // フォールバック：従来の方法
                return Globals.ThisAddIn.Application;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Failed to get current active application, using fallback");
                return Globals.ThisAddIn.Application;
            }
        }

        /// <summary>
        /// 🆕 大量操作時のUNDO境界分割処理
        /// </summary>
        /// <param name="itemCount">処理するアイテム数</param>
        /// <param name="currentIndex">現在の処理インデックス</param>
        /// <param name="batchSize">バッチサイズ（デフォルト：50）</param>
        /// <param name="undoEntryBaseName">UNDO操作の基本名</param>
        public static void SplitUndoEntryIfNeeded(int itemCount, int currentIndex, int batchSize, string undoEntryBaseName)
        {
            // 大量操作の場合、適切な間隔でUNDO境界を分割
            if (itemCount > batchSize && currentIndex > 0 && currentIndex % batchSize == 0)
            {
                var batchNumber = (currentIndex / batchSize) + 1;
                var undoEntryName = $"{undoEntryBaseName} (バッチ{batchNumber})";
                StartNewUndoEntry(undoEntryName);

                logger.Debug($"Split UNDO entry at index {currentIndex}: '{undoEntryName}'");
            }
        }

        /// <summary>
        /// GCを強制実行してメモリを解放します
        /// </summary>
        public static void ForceGarbageCollection()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            logger.Debug("Forced garbage collection completed");
        }
    }
}