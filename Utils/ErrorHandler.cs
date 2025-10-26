using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NLog;

namespace PowerPointEfficiencyAddin.Utils
{
    /// <summary>
    /// エラーハンドリングを統一管理するクラス
    /// </summary>
    public static class ErrorHandler
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 処理を安全に実行し、エラーが発生した場合は適切に処理します
        /// </summary>
        /// <param name="action">実行する処理</param>
        /// <param name="operationName">操作名（ログ用）</param>
        /// <param name="showUserMessage">ユーザーにメッセージを表示するか</param>
        /// <returns>処理が成功したかどうか</returns>
        public static bool ExecuteSafely(Action action, string operationName, bool showUserMessage = true)
        {
            try
            {
                action();
                logger.Info($"Operation completed successfully: {operationName}");
                return true;
            }
            catch (COMException comEx)
            {
                HandleComException(comEx, operationName, showUserMessage);
                return false;
            }
            catch (Exception ex)
            {
                HandleGeneralException(ex, operationName, showUserMessage);
                return false;
            }
        }

        /// <summary>
        /// 戻り値のある処理を安全に実行します
        /// </summary>
        /// <typeparam name="T">戻り値の型</typeparam>
        /// <param name="func">実行する処理</param>
        /// <param name="operationName">操作名（ログ用）</param>
        /// <param name="defaultValue">エラー時のデフォルト値</param>
        /// <param name="showUserMessage">ユーザーにメッセージを表示するか</param>
        /// <returns>処理結果またはデフォルト値</returns>
        public static T ExecuteSafely<T>(Func<T> func, string operationName, T defaultValue = default(T), bool showUserMessage = true)
        {
            try
            {
                T result = func();
                logger.Info($"Operation completed successfully: {operationName}");
                return result;
            }
            catch (COMException comEx)
            {
                HandleComException(comEx, operationName, showUserMessage);
                return defaultValue;
            }
            catch (Exception ex)
            {
                HandleGeneralException(ex, operationName, showUserMessage);
                return defaultValue;
            }
        }

        /// <summary>
        /// COMExceptionを処理します
        /// </summary>
        /// <param name="comEx">COMException</param>
        /// <param name="operationName">操作名</param>
        /// <param name="showUserMessage">ユーザーにメッセージを表示するか</param>
        private static void HandleComException(COMException comEx, string operationName, bool showUserMessage)
        {
            string errorMessage = GetComErrorMessage(comEx);
            logger.Error(comEx, $"COM error in {operationName}: {errorMessage}");

            if (showUserMessage)
            {
                ShowErrorMessage($"操作中にエラーが発生しました: {operationName}", errorMessage);
            }
        }

        /// <summary>
        /// 一般例外を処理します
        /// </summary>
        /// <param name="ex">Exception</param>
        /// <param name="operationName">操作名</param>
        /// <param name="showUserMessage">ユーザーにメッセージを表示するか</param>
        private static void HandleGeneralException(Exception ex, string operationName, bool showUserMessage)
        {
            logger.Error(ex, $"General error in {operationName}: {ex.Message}");

            if (showUserMessage)
            {
                ShowErrorMessage($"予期しないエラーが発生しました: {operationName}", ex.Message);
            }
        }

        /// <summary>
        /// COMExceptionから適切なエラーメッセージを取得します
        /// </summary>
        /// <param name="comEx">COMException</param>
        /// <returns>エラーメッセージ</returns>
        private static string GetComErrorMessage(COMException comEx)
        {
            switch ((uint)comEx.HResult)
            {
                case 0x800A03EC: // Selection is empty
                    return "図形が選択されていません。操作する図形を選択してください。";
                case 0x800A01A8: // Object doesn't support this property or method
                    return "選択されたオブジェクトではこの操作を実行できません。";
                case 0x800A000D: // Type mismatch
                    return "操作に適さないオブジェクトが選択されています。";
                case 0x80004005: // Unspecified error
                    return "PowerPointとの通信でエラーが発生しました。";
                default:
                    return $"COM エラー (0x{comEx.HResult:X8}): {comEx.Message}";
            }
        }

        /// <summary>
        /// ユーザーにエラーメッセージを表示します
        /// </summary>
        /// <param name="title">タイトル</param>
        /// <param name="message">メッセージ</param>
        private static void ShowErrorMessage(string title, string message)
        {
            MessageBox.Show(
                message,
                title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }

        /// <summary>
        /// 選択状態を検証します
        /// </summary>
        /// <param name="selectionCount">選択数</param>
        /// <param name="minRequired">最小必要数</param>
        /// <param name="maxAllowed">最大許可数（0は無制限）</param>
        /// <param name="operationName">操作名</param>
        /// <returns>選択状態が有効かどうか</returns>
        public static bool ValidateSelection(int selectionCount, int minRequired, int maxAllowed, string operationName)
        {
            if (selectionCount < minRequired)
            {
                string message = minRequired == 1
                    ? "図形を選択してください。"
                    : $"最低{minRequired}つの図形を選択してください。";

                ShowErrorMessage($"{operationName} - 選択エラー", message);
                logger.Warn($"Insufficient selection for {operationName}: {selectionCount} (required: {minRequired})");
                return false;
            }

            if (maxAllowed > 0 && selectionCount > maxAllowed)
            {
                ShowErrorMessage(
                    $"{operationName} - 選択エラー",
                    $"選択できる図形は最大{maxAllowed}つまでです。"
                );
                logger.Warn($"Too many selections for {operationName}: {selectionCount} (max: {maxAllowed})");
                return false;
            }

            return true;
        }
    }
}