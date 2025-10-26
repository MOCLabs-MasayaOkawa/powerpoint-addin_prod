using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace PowerPointEfficiencyAddin.Utils
{
    /// <summary>
    /// ログ設定を管理するクラス
    /// </summary>
    public static class LoggerConfig
    {
        private static bool _isInitialized = false;

        /// <summary>
        /// ログ設定を初期化します
        /// </summary>
        public static void Initialize()
        {
            if (_isInitialized) return;

            try
            {
                var config = new LoggingConfiguration();

                // ログファイルのパス設定
                string logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointEfficiencyAddin",
                    "Logs"
                );

                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // ファイルターゲットの設定（シンプル版）
                var fileTarget = new FileTarget("fileTarget")
                {
                    FileName = Path.Combine(logDirectory, "PowerPointAddin-${shortdate}.log"),
                    Layout = "${longdate} ${level:uppercase=true:padding=-5} ${logger:shortName=true} ${message} ${exception:format=tostring}",
                    Encoding = System.Text.Encoding.UTF8
                };

                // デバッグターゲットの設定（開発時用）
                var debugTarget = new DebugTarget("debugTarget")
                {
                    Layout = "${time} ${level:uppercase=true:padding=-5} ${logger:shortName=true} ${message}"
                };

                // ルール設定
                config.AddTarget(fileTarget);
                config.AddTarget(debugTarget);

#if DEBUG
                config.AddRuleForAllLevels(debugTarget);
                config.AddRuleForAllLevels(fileTarget);
#else
                config.AddRuleForOneLevel(LogLevel.Info, fileTarget);
                config.AddRuleForOneLevel(LogLevel.Warn, fileTarget);
                config.AddRuleForOneLevel(LogLevel.Error, fileTarget);
                config.AddRuleForOneLevel(LogLevel.Fatal, fileTarget);
#endif

                LogManager.Configuration = config;
                _isInitialized = true;

                var logger = LogManager.GetCurrentClassLogger();
                logger.Info("PowerPoint Efficiency Addin logging initialized");
            }
            catch (Exception ex)
            {
                // ログ初期化に失敗した場合でもアドインは動作させる
                System.Diagnostics.Debug.WriteLine($"Failed to initialize logging: {ex.Message}");
            }
        }

        /// <summary>
        /// ログ設定をシャットダウンします
        /// </summary>
        public static void Shutdown()
        {
            if (_isInitialized)
            {
                LogManager.Shutdown();
                _isInitialized = false;
            }
        }

        /// <summary>
        /// パフォーマンス測定を行うログ記録器を作成します
        /// </summary>
        /// <param name="operationName">操作名</param>
        /// <returns>パフォーマンス測定器</returns>
        public static IDisposable MeasurePerformance(string operationName)
        {
            return new PerformanceMeasurer(operationName);
        }

        /// <summary>
        /// パフォーマンス測定を行うクラス
        /// </summary>
        private class PerformanceMeasurer : IDisposable
        {
            private readonly string _operationName;
            private readonly DateTime _startTime;
            private readonly Logger _logger;
            private bool _disposed = false;

            public PerformanceMeasurer(string operationName)
            {
                _operationName = operationName;
                _startTime = DateTime.Now;
                _logger = LogManager.GetCurrentClassLogger();
                _logger.Debug($"Performance measurement started: {_operationName}");
            }

            public void Dispose()
            {
                if (!_disposed)
                {
                    var elapsed = DateTime.Now - _startTime;
                    _logger.Debug($"Performance measurement completed: {_operationName} - {elapsed.TotalMilliseconds:F2}ms");
                    _disposed = true;
                }
            }
        }
    }
}